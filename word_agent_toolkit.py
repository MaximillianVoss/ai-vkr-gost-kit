from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


REPO_ROOT = Path(__file__).resolve().parent
DEFAULT_SPEC_PATH = Path(__file__).resolve().with_name("ai-client-vkr-gost-spec.json")
WORD_FORMAT_DOCX = 16
WORD_FORMAT_PDF = 17

ALIGNMENTS = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def load_spec(spec_path: Path | None) -> dict[str, Any]:
    current_path = spec_path or DEFAULT_SPEC_PATH
    with current_path.open("r", encoding="utf-8") as file:
        return json.load(file)


def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def absolute_path(path: Path) -> str:
    return str(path.resolve())


def import_word_client():
    import win32com.client  # type: ignore[import]

    return win32com.client


def create_word_application():
    client = import_word_client()
    word = client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    word.AutomationSecurity = 3
    return word


def save_json(data: dict[str, Any], output_path: Path | None) -> None:
    payload = json.dumps(data, ensure_ascii=False, indent=2)
    if output_path is None:
        print(payload)
        return
    ensure_parent(output_path)
    output_path.write_text(payload, encoding="utf-8")


def read_package_styles(document_path: Path) -> dict[str, Any]:
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with ZipFile(document_path) as archive:
        styles_xml = archive.read("word/styles.xml")
    root = ET.fromstring(styles_xml)

    styles_by_type: dict[str, list[dict[str, str]]] = {
        "paragraph": [],
        "character": [],
        "table": [],
        "numbering": [],
    }

    for style in root.findall("w:style", ns):
        style_type = style.get(f"{{{ns['w']}}}type", "unknown")
        style_id = style.get(f"{{{ns['w']}}}styleId", "")
        name_element = style.find("w:name", ns)
        style_name = name_element.get(f"{{{ns['w']}}}val", style_id) if name_element is not None else style_id
        target = styles_by_type.setdefault(style_type, [])
        target.append({"style_id": style_id, "name": style_name})

    for values in styles_by_type.values():
        values.sort(key=lambda item: item["name"].lower())

    return {
        "document_path": str(document_path),
        "styles": styles_by_type,
    }


def create_docx_from_template(template_path: Path, output_path: Path) -> None:
    ensure_parent(output_path)
    word = create_word_application()
    document = None
    try:
        document = word.Documents.Add(absolute_path(template_path))
        document.SaveAs(absolute_path(output_path), FileFormat=WORD_FORMAT_DOCX)
    finally:
        if document is not None:
            document.Close(False)
        word.Quit()


def refresh_fields(docx_path: Path) -> None:
    word = create_word_application()
    document = None
    try:
        document = word.Documents.Open(absolute_path(docx_path))
        document.Fields.Update()
        for toc in document.TablesOfContents:
            toc.Update()
        document.Save()
    finally:
        if document is not None:
            document.Close(False)
        word.Quit()


def export_pdf(docx_path: Path, pdf_path: Path) -> None:
    ensure_parent(pdf_path)
    word = create_word_application()
    document = None
    try:
        document = word.Documents.Open(absolute_path(docx_path))
        document.Fields.Update()
        for toc in document.TablesOfContents:
            toc.Update()
        document.ExportAsFixedFormat(absolute_path(pdf_path), WORD_FORMAT_PDF)
    finally:
        if document is not None:
            document.Close(False)
        word.Quit()


def get_style_name(document: Document, candidates: list[str], fallback: str) -> str:
    existing_names = {style.name for style in document.styles}
    for candidate in candidates:
        if candidate in existing_names:
            return candidate
    return fallback


def get_existing_style(document: Document, candidates: list[str]) -> Any | None:
    existing_names = {style.name for style in document.styles}
    for candidate in candidates:
        if candidate in existing_names:
            return document.styles[candidate]
    return None


def clear_document(document: Document) -> None:
    body = document.element.body
    for element in list(body):
        if element.tag != qn("w:sectPr"):
            body.remove(element)


def ensure_font_family(style, font_name: str) -> None:
    style.font.name = font_name
    rpr = style._element.get_or_add_rPr()
    if rpr.rFonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rpr.rFonts.set(qn("w:ascii"), font_name)
    rpr.rFonts.set(qn("w:hAnsi"), font_name)
    rpr.rFonts.set(qn("w:eastAsia"), font_name)
    rpr.rFonts.set(qn("w:cs"), font_name)


def clear_style_numbering(style) -> None:
    paragraph_properties = style._element.find(qn("w:pPr"))
    if paragraph_properties is None:
        return
    for child in list(paragraph_properties):
        if child.tag == qn("w:numPr"):
            paragraph_properties.remove(child)


def apply_paragraph_style(
    style,
    *,
    font_name: str,
    font_size_pt: float,
    bold: bool | None = None,
    alignment: str | None = None,
    line_spacing: float | None = None,
    first_line_indent_cm: float | None = None,
) -> None:
    clear_style_numbering(style)
    ensure_font_family(style, font_name)
    style.font.size = Pt(font_size_pt)
    if bold is not None:
        style.font.bold = bold

    paragraph_format = style.paragraph_format
    if alignment is not None:
        paragraph_format.alignment = ALIGNMENTS[alignment]
    if line_spacing is not None:
        paragraph_format.line_spacing = line_spacing
    if first_line_indent_cm is None:
        paragraph_format.first_line_indent = None
    else:
        paragraph_format.first_line_indent = Cm(first_line_indent_cm)
    paragraph_format.space_before = Pt(0)
    paragraph_format.space_after = Pt(0)


def apply_run_font(run, *, font_name: str = "Times New Roman", font_size_pt: float | None = None, bold: bool | None = None) -> None:
    run.font.name = font_name
    rpr = run._element.get_or_add_rPr()
    if rpr.rFonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rpr.rFonts.set(qn("w:ascii"), font_name)
    rpr.rFonts.set(qn("w:hAnsi"), font_name)
    rpr.rFonts.set(qn("w:eastAsia"), font_name)
    rpr.rFonts.set(qn("w:cs"), font_name)
    if font_size_pt is not None:
        run.font.size = Pt(font_size_pt)
    if bold is not None:
        run.bold = bold


def add_field(paragraph, instruction: str, placeholder: str = "") -> None:
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")

    paragraph.add_run()._r.append(begin)
    paragraph.add_run()._r.append(instr)
    paragraph.add_run()._r.append(separate)
    if placeholder:
        apply_run_font(paragraph.add_run(placeholder), font_size_pt=12)
    paragraph.add_run()._r.append(end)


def set_outline_level(paragraph, level: int) -> None:
    paragraph_properties = paragraph._p.get_or_add_pPr()
    for child in list(paragraph_properties):
        if child.tag == qn("w:outlineLvl"):
            paragraph_properties.remove(child)
    outline = OxmlElement("w:outlineLvl")
    outline.set(qn("w:val"), str(level))
    paragraph_properties.append(outline)


def add_paragraph(
    document: Document,
    style_name: str,
    text: str = "",
    *,
    alignment: str = "justify",
    first_line_indent_cm: float | None = 1.25,
    bold: bool = False,
    font_size_pt: float | None = None,
    outline_level: int | None = None,
) -> Any:
    paragraph = document.add_paragraph()
    paragraph.style = style_name
    paragraph.alignment = ALIGNMENTS[alignment]
    if first_line_indent_cm is None:
        paragraph.paragraph_format.first_line_indent = None
    else:
        paragraph.paragraph_format.first_line_indent = Cm(first_line_indent_cm)
    if outline_level is not None:
        set_outline_level(paragraph, outline_level)
    if text:
        apply_run_font(paragraph.add_run(text), font_size_pt=font_size_pt, bold=bold)
    return paragraph


def add_blank(document: Document, style_name: str, count: int = 1) -> None:
    for _ in range(count):
        add_paragraph(document, style_name, alignment="left", first_line_indent_cm=None)


def set_default_section(section, spec: dict[str, Any]) -> None:
    margins = spec["global_formatting_rules"]["margins_mm"]
    section.left_margin = Cm(margins["left"] / 10)
    section.right_margin = Cm(margins["right"] / 10)
    section.top_margin = Cm(margins["top"] / 10)
    section.bottom_margin = Cm(margins["bottom"] / 10)
    section.header_distance = Cm(spec["global_formatting_rules"]["header_distance_cm"])
    section.footer_distance = Cm(spec["global_formatting_rules"]["footer_distance_cm"])


def apply_gost_profile(input_path: Path, output_path: Path | None, spec_path: Path | None) -> Path:
    spec = load_spec(spec_path)
    target_path = output_path or input_path
    document = Document(str(input_path))

    style_map = spec["template_style_map"]
    body_rules = spec["global_formatting_rules"]["main_text"]
    body_style_name = get_style_name(document, style_map["body"], "Normal")
    body_style = document.styles[body_style_name]
    apply_paragraph_style(
        body_style,
        font_name=body_rules["font_name"],
        font_size_pt=body_rules["font_size_pt"],
        alignment=body_rules["alignment"],
        line_spacing=body_rules["line_spacing"],
        first_line_indent_cm=body_rules["first_line_indent_cm"],
    )

    chapter_rules = spec["heading_rules"]["chapter_title"]
    chapter_style = get_existing_style(document, style_map["chapter_title"])
    if chapter_style is not None:
        apply_paragraph_style(
            chapter_style,
            font_name=chapter_rules["font_name"],
            font_size_pt=chapter_rules["font_size_pt"],
            bold=chapter_rules["bold"],
            alignment=chapter_rules["alignment"],
            line_spacing=chapter_rules["multi_line_spacing"],
            first_line_indent_cm=None,
        )

    section_rules = spec["heading_rules"]["section_title"]
    section_style = get_existing_style(document, style_map["section_title"])
    if section_style is not None:
        apply_paragraph_style(
            section_style,
            font_name=section_rules["font_name"],
            font_size_pt=section_rules["font_size_pt"],
            bold=section_rules["bold"],
            alignment=section_rules["alignment"],
            line_spacing=section_rules["multi_line_spacing"],
            first_line_indent_cm=None,
        )

    subsection_rules = spec["heading_rules"]["subsection_title"]
    subsection_style = get_existing_style(document, style_map["subsection_title"])
    if subsection_style is not None:
        apply_paragraph_style(
            subsection_style,
            font_name=subsection_rules["font_name"],
            font_size_pt=subsection_rules["font_size_pt"],
            bold=subsection_rules["bold"],
            alignment=subsection_rules["alignment"],
            line_spacing=subsection_rules["multi_line_spacing"],
            first_line_indent_cm=None,
        )

    caption_style = get_existing_style(document, style_map["caption"])
    if caption_style is not None:
        apply_paragraph_style(
            caption_style,
            font_name=spec["table_rules"]["table_font"]["font_name"],
            font_size_pt=12,
            alignment="center",
            line_spacing=1.0,
            first_line_indent_cm=None,
        )

    toc_title_style = get_existing_style(document, style_map["toc_heading"])
    if toc_title_style is not None:
        toc_title_rules = spec["toc_rules"]["title_format"]
        apply_paragraph_style(
            toc_title_style,
            font_name=toc_title_rules["font_name"],
            font_size_pt=toc_title_rules["font_size_pt"],
            bold=toc_title_rules["bold"],
            alignment=toc_title_rules["alignment"],
            line_spacing=toc_title_rules["line_spacing"],
            first_line_indent_cm=None,
        )

    toc_entry_candidates = list(style_map["toc_entry"])
    toc_entry_candidates.extend([f"TOC {index}" for index in range(1, 10)])
    toc_entry_candidates.extend([f"toc {index}" for index in range(1, 10)])
    toc_entry_rules = spec["toc_rules"]["entry_format"]
    applied_toc_styles: set[str] = set()
    for candidate in toc_entry_candidates:
        if candidate in applied_toc_styles:
            continue
        style = get_existing_style(document, [candidate])
        if style is None:
            continue
        apply_paragraph_style(
            style,
            font_name=toc_entry_rules["font_name"],
            font_size_pt=toc_entry_rules["font_size_pt"],
            alignment=toc_entry_rules["alignment"],
            line_spacing=toc_entry_rules["line_spacing"],
            first_line_indent_cm=toc_entry_rules["first_line_indent_cm"],
        )
        applied_toc_styles.add(candidate)

    for section in document.sections:
        set_default_section(section, spec)

    ensure_parent(target_path)
    document.save(target_path)
    return target_path


def append_code_appendix(
    input_path: Path,
    source_path: Path,
    title: str,
    appendix_label: str,
    output_path: Path | None,
    spec_path: Path | None,
    line_numbers: bool,
) -> Path:
    spec = load_spec(spec_path)
    target_path = output_path or input_path
    document = Document(str(input_path))
    style_map = spec["template_style_map"]
    body_style_name = get_style_name(document, style_map["body"], "Normal")

    document.add_page_break()

    appendix_paragraph = document.add_paragraph()
    appendix_paragraph.style = body_style_name
    appendix_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    appendix_paragraph.paragraph_format.first_line_indent = None
    appendix_paragraph.add_run(f"Приложение {appendix_label}").bold = True

    title_paragraph = document.add_paragraph()
    title_paragraph.style = body_style_name
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_paragraph.paragraph_format.first_line_indent = None
    title_paragraph.add_run(title).bold = True

    document.add_paragraph().style = body_style_name

    code_text = source_path.read_text(encoding="utf-8").replace("\t", "    ")
    lines = code_text.splitlines()
    if line_numbers:
        lines = [f"{line_no:04}: {line}" if line else f"{line_no:04}:" for line_no, line in enumerate(lines, start=1)]

    code_paragraph = document.add_paragraph()
    code_paragraph.style = body_style_name
    code_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    code_paragraph.paragraph_format.first_line_indent = None
    code_paragraph.paragraph_format.line_spacing = 1
    run = code_paragraph.add_run("\n".join(lines))
    run.font.name = "Courier New"
    run.font.size = Pt(9)
    rpr = run._element.get_or_add_rPr()
    if rpr.rFonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rpr.rFonts.set(qn("w:ascii"), "Courier New")
    rpr.rFonts.set(qn("w:hAnsi"), "Courier New")
    rpr.rFonts.set(qn("w:eastAsia"), "Courier New")
    rpr.rFonts.set(qn("w:cs"), "Courier New")

    ensure_parent(target_path)
    document.save(target_path)
    return target_path


def summarize_docx(docx_path: Path) -> dict[str, Any]:
    document = Document(str(docx_path))
    non_empty_paragraphs = [paragraph for paragraph in document.paragraphs if paragraph.text.strip()]
    styles_used = sorted({paragraph.style.name for paragraph in non_empty_paragraphs})
    return {
        "document_path": str(docx_path),
        "sections": len(document.sections),
        "paragraphs": len(document.paragraphs),
        "non_empty_paragraphs": len(non_empty_paragraphs),
        "tables": len(document.tables),
        "inline_shapes": len(document.inline_shapes),
        "styles_used": styles_used,
    }


def build_agent_brief_document(template_path: Path, output_path: Path, spec_path: Path | None) -> Path:
    spec = load_spec(spec_path)
    create_docx_from_template(template_path, output_path)
    document = Document(str(output_path))
    clear_document(document)

    style_map = spec["template_style_map"]
    body_style = get_style_name(document, style_map["body"], "Normal")
    chapter_style = body_style
    section_style = body_style

    apply_gost_profile(output_path, None, spec_path)
    document = Document(str(output_path))
    clear_document(document)

    for section in document.sections:
        set_default_section(section, spec)

    add_blank(document, body_style, 2)
    add_paragraph(
        document,
        chapter_style,
        "ПРОФИЛЬ ДЛЯ ИИ-АГЕНТА ПО ПОДГОТОВКЕ ВКР",
        alignment="center",
        first_line_indent_cm=None,
        bold=True,
        font_size_pt=16,
        outline_level=0,
    )
    add_paragraph(
        document,
        body_style,
        "Редактируемый Word-образец на базе пользовательского шаблона и ГОСТ-профиля.",
        alignment="center",
        first_line_indent_cm=None,
        font_size_pt=14,
    )
    add_blank(document, body_style, 3)
    add_paragraph(document, body_style, "Тема работы: [вписать точную тему]", alignment="left", first_line_indent_cm=None, bold=True)
    add_paragraph(document, body_style, "Тип работы: [ВКР / курсовая работа / дипломный проект / пояснительная записка]", alignment="left", first_line_indent_cm=None)
    add_paragraph(document, body_style, "Образовательная организация: [вписать название]", alignment="left", first_line_indent_cm=None)
    add_paragraph(document, body_style, "Кафедра / отделение: [вписать]", alignment="left", first_line_indent_cm=None)
    add_paragraph(document, body_style, "Автор: [ФИО]", alignment="left", first_line_indent_cm=None)
    add_paragraph(document, body_style, "Руководитель: [ФИО]", alignment="left", first_line_indent_cm=None)
    add_paragraph(document, body_style, "Группа: [номер группы]", alignment="left", first_line_indent_cm=None)
    add_paragraph(document, body_style, "Год: [год оформления]", alignment="left", first_line_indent_cm=None)

    document.add_page_break()
    add_paragraph(
        document,
        chapter_style,
        spec["toc_rules"]["page_title"],
        alignment="center",
        first_line_indent_cm=None,
        bold=True,
        font_size_pt=14,
        outline_level=0,
    )
    toc_paragraph = add_paragraph(document, body_style, alignment="left", first_line_indent_cm=None)
    add_field(toc_paragraph, r'TOC \o "1-2" \h \z \u', "Оглавление обновится после обновления полей.")

    document.add_page_break()
    add_paragraph(
        document,
        chapter_style,
        "1. НАЗНАЧЕНИЕ ДОКУМЕНТА",
        alignment="center",
        first_line_indent_cm=None,
        bold=True,
        font_size_pt=14,
        outline_level=0,
    )
    add_paragraph(document, body_style, spec["description"], font_size_pt=14)
    add_paragraph(document, body_style, spec["source_assets"]["mandatory_usage_rule"], font_size_pt=14)
    add_paragraph(document, body_style, spec["automation_tooling"]["agent_usage_rule"], font_size_pt=14)

    add_paragraph(
        document,
        section_style,
        "1.1 Что пользователь должен заполнить перед отправкой ИИ",
        alignment="left",
        first_line_indent_cm=None,
        bold=True,
        font_size_pt=14,
        outline_level=1,
    )
    for field_name, description in spec["required_input_fields"].items():
        add_paragraph(document, body_style, f"{field_name}: {description}", first_line_indent_cm=None)

    add_paragraph(
        document,
        chapter_style,
        "2. ОБЯЗАТЕЛЬНЫЕ ТРЕБОВАНИЯ ОФОРМЛЕНИЯ",
        alignment="center",
        first_line_indent_cm=None,
        bold=True,
        font_size_pt=14,
        outline_level=0,
    )
    formatting = spec["global_formatting_rules"]
    add_paragraph(
        document,
        body_style,
        (
            f"Основной текст оформляется шрифтом {formatting['main_text']['font_name']} "
            f"{formatting['main_text']['font_size_pt']} pt, межстрочный интервал "
            f"{formatting['main_text']['line_spacing']}, абзацный отступ "
            f"{formatting['main_text']['first_line_indent_cm']} см, выравнивание по ширине."
        ),
    )
    add_paragraph(
        document,
        body_style,
        (
            f"Поля страницы: левое {formatting['margins_mm']['left']} мм, правое "
            f"{formatting['margins_mm']['right']} мм, верхнее {formatting['margins_mm']['top']} мм, "
            f"нижнее {formatting['margins_mm']['bottom']} мм."
        ),
    )
    numbering = formatting["page_numbering"]
    add_paragraph(
        document,
        body_style,
        (
            f"Нумерация страниц арабская, положение {numbering['position']}, печатать номер с раздела "
            f"«{numbering['print_from_section']}». Титульный лист и содержание входят в общий счет, "
            f"но номер на них не печатается."
        ),
    )

    add_paragraph(
        document,
        section_style,
        "2.1 Оглавление",
        alignment="left",
        first_line_indent_cm=None,
        bold=True,
        font_size_pt=14,
        outline_level=1,
    )
    toc_rules = spec["toc_rules"]
    add_paragraph(
        document,
        body_style,
        (
            f"Заголовок страницы оглавления и сами записи должны быть оформлены шрифтом "
            f"{toc_rules['entry_format']['font_name']} {toc_rules['entry_format']['font_size_pt']} pt "
            f"с межстрочным интервалом {toc_rules['entry_format']['line_spacing']}."
        ),
    )
    for rule in toc_rules["must_include"]:
        add_paragraph(document, body_style, f"- {rule}", first_line_indent_cm=None)
    for rule in toc_rules["must_not"]:
        add_paragraph(document, body_style, f"- {rule}", first_line_indent_cm=None)

    add_paragraph(
        document,
        chapter_style,
        "3. РЕКОМЕНДУЕМАЯ СТРУКТУРА РАБОТЫ",
        alignment="center",
        first_line_indent_cm=None,
        bold=True,
        font_size_pt=14,
        outline_level=0,
    )
    for item in spec["content_structure_rules"]["minimum_required_sections"]:
        add_paragraph(document, body_style, f"- {item}", first_line_indent_cm=None)
    add_paragraph(
        document,
        section_style,
        "3.1 Рекомендуемая структура ИТ-проекта",
        alignment="left",
        first_line_indent_cm=None,
        bold=True,
        font_size_pt=14,
        outline_level=1,
    )
    for item in spec["content_structure_rules"]["recommended_structure_for_it_project"]:
        add_paragraph(document, body_style, f"- {item}", first_line_indent_cm=None)

    add_paragraph(
        document,
        chapter_style,
        "4. ТАБЛИЦЫ, РИСУНКИ, ПРИЛОЖЕНИЯ И ЛИСТИНГИ",
        alignment="center",
        first_line_indent_cm=None,
        bold=True,
        font_size_pt=14,
        outline_level=0,
    )
    add_paragraph(document, body_style, spec["table_rules"]["reference_rule"])
    add_paragraph(
        document,
        body_style,
        (
            f"Подпись таблицы оформлять как «{spec['table_rules']['caption_layout']['number_line']}», "
            f"слово «Таблица» размещать {spec['table_rules']['caption_layout']['number_alignment']}."
        ),
    )
    add_paragraph(document, body_style, f"Подпись рисунка: {spec['figure_rules']['caption_format']}.")
    add_paragraph(document, body_style, spec["appendix_rules"]["start_rule"])
    add_paragraph(document, body_style, spec["appendix_rules"]["label_rule"])
    add_paragraph(document, body_style, spec["code_listing_rules"]["placement"])

    add_paragraph(
        document,
        chapter_style,
        "5. КОМАНДЫ АВТОМАТИЗАЦИИ WORD",
        alignment="center",
        first_line_indent_cm=None,
        bold=True,
        font_size_pt=14,
        outline_level=0,
    )
    add_paragraph(document, body_style, "Ниже приведены готовые команды для другого ИИ-агента.", first_line_indent_cm=None)
    for command in spec["automation_tooling"]["recommended_commands"]:
        add_paragraph(document, body_style, command, alignment="left", first_line_indent_cm=None)

    add_paragraph(
        document,
        chapter_style,
        "6. ЧЕК-ЛИСТ ПЕРЕД ФИНАЛЬНОЙ СДАЧЕЙ",
        alignment="center",
        first_line_indent_cm=None,
        bold=True,
        font_size_pt=14,
        outline_level=0,
    )
    for item in spec["validation_checklist"]:
        add_paragraph(document, body_style, f"- {item}", first_line_indent_cm=None)

    add_paragraph(
        document,
        chapter_style,
        "7. ЗАПРЕТЫ ДЛЯ ИИ",
        alignment="center",
        first_line_indent_cm=None,
        bold=True,
        font_size_pt=14,
        outline_level=0,
    )
    for item in spec["strict_prohibitions"]:
        add_paragraph(document, body_style, f"- {item}", first_line_indent_cm=None)

    ensure_parent(output_path)
    document.save(output_path)
    refresh_fields(output_path)
    return output_path


def command_inspect_template(args: argparse.Namespace) -> None:
    save_json(read_package_styles(args.template), args.output)


def command_create_from_template(args: argparse.Namespace) -> None:
    create_docx_from_template(args.template, args.output)
    print(str(args.output))


def command_refresh_fields(args: argparse.Namespace) -> None:
    refresh_fields(args.input)
    print(str(args.input))


def command_export_pdf(args: argparse.Namespace) -> None:
    export_pdf(args.input, args.output)
    print(str(args.output))


def command_apply_gost(args: argparse.Namespace) -> None:
    target = apply_gost_profile(args.input, args.output, args.spec)
    print(str(target))


def command_append_code_appendix(args: argparse.Namespace) -> None:
    target = append_code_appendix(
        input_path=args.input,
        source_path=args.source,
        title=args.title,
        appendix_label=args.label,
        output_path=args.output,
        spec_path=args.spec,
        line_numbers=not args.no_line_numbers,
    )
    print(str(target))


def command_summarize_doc(args: argparse.Namespace) -> None:
    save_json(summarize_docx(args.input), args.output)


def command_finalize_doc(args: argparse.Namespace) -> None:
    working_path = apply_gost_profile(args.input, None, args.spec)
    refresh_fields(working_path)
    if args.pdf is not None:
        export_pdf(working_path, args.pdf)
        print(json.dumps({"docx": str(working_path), "pdf": str(args.pdf)}, ensure_ascii=False, indent=2))
        return
    print(str(working_path))


def command_generate_agent_brief(args: argparse.Namespace) -> None:
    target = build_agent_brief_document(args.template, args.output, args.spec)
    if args.pdf is not None:
        export_pdf(target, args.pdf)
        print(json.dumps({"docx": str(target), "pdf": str(args.pdf)}, ensure_ascii=False, indent=2))
        return
    print(str(target))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="CLI-инструмент для ИИ-агентов, работающих с Word-документами по шаблону и ГОСТ-профилю.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    inspect_parser = subparsers.add_parser("inspect-template", help="Извлечь стили из .dotm/.docx в JSON.")
    inspect_parser.add_argument("--template", type=Path, required=True)
    inspect_parser.add_argument("--output", type=Path)
    inspect_parser.set_defaults(func=command_inspect_template)

    create_parser = subparsers.add_parser("create-from-template", help="Создать .docx из .dotm-шаблона.")
    create_parser.add_argument("--template", type=Path, required=True)
    create_parser.add_argument("--output", type=Path, required=True)
    create_parser.set_defaults(func=command_create_from_template)

    refresh_parser = subparsers.add_parser("refresh-fields", help="Обновить поля, включая оглавление.")
    refresh_parser.add_argument("--input", type=Path, required=True)
    refresh_parser.set_defaults(func=command_refresh_fields)

    export_parser = subparsers.add_parser("export-pdf", help="Обновить поля и выгрузить PDF.")
    export_parser.add_argument("--input", type=Path, required=True)
    export_parser.add_argument("--output", type=Path, required=True)
    export_parser.set_defaults(func=command_export_pdf)

    gost_parser = subparsers.add_parser("apply-gost", help="Применить ГОСТ-профиль к стилям и секциям документа.")
    gost_parser.add_argument("--input", type=Path, required=True)
    gost_parser.add_argument("--output", type=Path)
    gost_parser.add_argument("--spec", type=Path, default=DEFAULT_SPEC_PATH)
    gost_parser.set_defaults(func=command_apply_gost)

    appendix_parser = subparsers.add_parser(
        "append-code-appendix",
        help="Добавить приложение с листингом кода в конец документа.",
    )
    appendix_parser.add_argument("--input", type=Path, required=True)
    appendix_parser.add_argument("--source", type=Path, required=True)
    appendix_parser.add_argument("--title", required=True)
    appendix_parser.add_argument("--label", required=True)
    appendix_parser.add_argument("--output", type=Path)
    appendix_parser.add_argument("--spec", type=Path, default=DEFAULT_SPEC_PATH)
    appendix_parser.add_argument("--no-line-numbers", action="store_true")
    appendix_parser.set_defaults(func=command_append_code_appendix)

    summarize_parser = subparsers.add_parser("summarize-doc", help="Выдать краткую JSON-сводку по .docx.")
    summarize_parser.add_argument("--input", type=Path, required=True)
    summarize_parser.add_argument("--output", type=Path)
    summarize_parser.set_defaults(func=command_summarize_doc)

    finalize_parser = subparsers.add_parser(
        "finalize-doc",
        help="Применить ГОСТ-профиль, обновить поля и при необходимости экспортировать PDF.",
    )
    finalize_parser.add_argument("--input", type=Path, required=True)
    finalize_parser.add_argument("--spec", type=Path, default=DEFAULT_SPEC_PATH)
    finalize_parser.add_argument("--pdf", type=Path)
    finalize_parser.set_defaults(func=command_finalize_doc)

    brief_parser = subparsers.add_parser(
        "generate-agent-brief",
        help="Создать Word-образец с инструкциями для ИИ на базе шаблона и ГОСТ-профиля.",
    )
    brief_parser.add_argument("--template", type=Path, required=True)
    brief_parser.add_argument("--output", type=Path, required=True)
    brief_parser.add_argument("--spec", type=Path, default=DEFAULT_SPEC_PATH)
    brief_parser.add_argument("--pdf", type=Path)
    brief_parser.set_defaults(func=command_generate_agent_brief)

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()

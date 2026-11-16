import io
import os
from docx import Document
from pptx import Presentation
from PIL import Image, ImageDraw, ImageFont

TEMPLATE_DIR = "templates"
DOCX_TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "Дилерский_договор_Мишка-2_2.docx")
PPTX_TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "MISHKA_certificate_A4.pptx")
CERT_BG_PATH = os.path.join(TEMPLATE_DIR, "cert_bg.png")
FONT_PATH = os.path.join(TEMPLATE_DIR, "font.ttf")


def _replace_text_in_paragraph(paragraph, replacements: dict):
    """Заменяет текст внутри параграфа Word с сохранением базового форматирования."""
    full_text = paragraph.text
    if not full_text:
        return

    modified = False
    for old_val, new_val in replacements.items():
        if old_val in full_text:
            full_text = full_text.replace(old_val, str(new_val))
            modified = True

    if modified:
        if paragraph.runs:
            first_run_font = paragraph.runs[0].font.name
            first_run_size = paragraph.runs[0].font.size
            first_run_bold = paragraph.runs[0].font.bold
            for run in paragraph.runs:
                run.text = ""
            paragraph.runs[0].text = full_text
            paragraph.runs[0].font.name = first_run_font
            paragraph.runs[0].font.size = first_run_size
            paragraph.runs[0].font.bold = first_run_bold
        else:
            paragraph.text = full_text


def build_dealer_contract(data: dict) -> io.BytesIO:
    """
    Заполняет шаблон Дилерского договора (.docx) и возвращает поток байтов.
    Поддерживает как плейсхолдеры {{...}}, так и замену по исходному тексту шаблона.
    """
    doc = Document(DOCX_TEMPLATE_PATH)

    contract_num = data.get("contract_number", "777")
    contract_date = data.get("contract_date", "«10» ноября 2026")
    dealer_name = data.get("dealer_name", "")
    dealer_sig_gen = data.get("dealer_signatory_gen", "")
    dealer_short = data.get("dealer_short_sign", "")
    dealer_address = data.get("dealer_address", "")
    dealer_inn_kpp = data.get("dealer_inn_kpp", "")
    dealer_rs = data.get("dealer_rs", "")
    dealer_bank = data.get("dealer_bank", "")
    dealer_ks = data.get("dealer_ks", "")
    dealer_bik = data.get("dealer_bik", "")
    dealer_phone = data.get("dealer_phone", "")
    dealer_email = data.get("dealer_email", "____________________")

    replacements = {
        "{{CONTRACT_NUMBER}}": contract_num,
        "№ 777": f"№ {contract_num}",
        "{{CONTRACT_DATE}}": contract_date,
        "«10» ноября 2026": contract_date,
        "{{DEALER_NAME}}": dealer_name,
        "ООО МАЧНЕВ И КОШКИ": dealer_name,
        "{{DEALER_SIGNATORY_GEN}}": dealer_sig_gen,
        "Генерального директора Мачнева А.И.": dealer_sig_gen,
        "{{DEALER_SHORT_SIGN}}": dealer_short,
        "Мачнев А.И.": dealer_short,
        "{{DEALER_ADDRESS}}": dealer_address,
        "111531, г. Москва, ул. шоссе Энтузиастов, д. 100": dealer_address,
        "{{DEALER_INN_KPP}}": dealer_inn_kpp,
        "4953077739 / 4953088974": dealer_inn_kpp,
        "{{DEALER_RS}}": dealer_rs,
        "40040040040040040044": dealer_rs,
        "{{DEALER_BANK}}": dealer_bank,
        "ПАО Альфа-банк": dealer_bank,
        "{{DEALER_KS}}": dealer_ks,
        "50050050050050050055": dealer_ks,
        "{{DEALER_BIK}}": dealer_bik,
        "01020304": dealer_bik,
        "{{DEALER_PHONE}}": dealer_phone,
        "+79060400043": dealer_phone,
        "{{DEALER_EMAIL}}": dealer_email,
    }

    for p in doc.paragraphs:
        _replace_text_in_paragraph(p, replacements)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "ДИЛЕР" in cell.text and "Эл. почта: ____________________" in cell.text:
                    for p in cell.paragraphs:
                        if "Эл. почта: ____________________" in p.text:
                            p.text = p.text.replace(
                                "Эл. почта: ____________________",
                                f"Эл. почта: {dealer_email}"
                            )
                for p in cell.paragraphs:
                    _replace_text_in_paragraph(p, replacements)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output


def build_certificate_pptx(data: dict) -> io.BytesIO | None:
    """
    Заполняет шаблон Сертификата Дилера (.pptx) с сохранением фирменного дизайна.
    """
    if not os.path.exists(PPTX_TEMPLATE_PATH):
        return None

    prs = Presentation(PPTX_TEMPLATE_PATH)

    cert_num = data.get("cert_number", "№ 2708/25")
    if not str(cert_num).startswith("№"):
        cert_num = f"№ {cert_num}"

    cert_company = data.get("cert_company", "")
    cert_inn = data.get("cert_inn", "")
    if cert_inn and not str(cert_inn).upper().startswith("ИНН"):
        cert_inn = f"ИНН {cert_inn}"

    cert_valid = data.get("cert_valid_until", "31.12.2027")
    if not str(cert_valid).lower().startswith("действителен"):
        cert_valid_full = f"Действителен до {cert_valid}"
    else:
        cert_valid_full = cert_valid

    replacements = {
        "{{CERT_NUMBER}}": cert_num,
        "№ 2708/25": cert_num,
        "{{CERT_COMPANY}}": cert_company,
        "ООО фирма «Мосводопром»": cert_company,
        "{{CERT_INN}}": cert_inn,
        "ИНН 7721023832": cert_inn,
        "{{CERT_VALID_UNTIL}}": cert_valid_full,
        "Действителен до 31.12.2027": cert_valid_full,
        "31.12.2027": cert_valid,
    }

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        for old_val, new_val in replacements.items():
                            if old_val in run.text:
                                run.text = run.text.replace(old_val, str(new_val))

                    full_p_text = "".join(r.text for r in paragraph.runs)
                    for old_val, new_val in replacements.items():
                        if old_val in full_p_text:
                            full_p_text = full_p_text.replace(old_val, str(new_val))
                            if paragraph.runs:
                                paragraph.runs[0].text = full_p_text
                                for r in paragraph.runs[1:]:
                                    r.text = ""

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output


def build_certificate_pdf_pillow(data: dict) -> io.BytesIO | None:
    """
    Опциональный генератор PDF-сертификата через Pillow (если в папке templates/ лежит cert_bg.png).
    """
    if not os.path.exists(CERT_BG_PATH):
        return None

    img = Image.open(CERT_BG_PATH).convert("RGB")
    draw = ImageDraw.Draw(img)
    width, height = img.size

    def get_font(size: int):
        if os.path.exists(FONT_PATH):
            return ImageFont.truetype(FONT_PATH, size)
        return ImageFont.load_default()

    cert_num = data.get("cert_number", "№ 2708/25")
    if not str(cert_num).startswith("№"):
        cert_num = f"№ {cert_num}"

    cert_company = data.get("cert_company", "")
    cert_inn = data.get("cert_inn", "")
    if cert_inn and not str(cert_inn).upper().startswith("ИНН"):
        cert_inn = f"ИНН {cert_inn}"
    cert_valid = data.get("cert_valid_until", "31.12.2027")

    center_x = int(width * 0.62)
    draw.text((center_x, int(height * 0.24)), cert_num, fill="#0D2C54", font=get_font(int(height * 0.035)), anchor="mm")
    draw.text((center_x, int(height * 0.43)), cert_company, fill="#081C38", font=get_font(int(height * 0.045)), anchor="mm")
    draw.text((center_x, int(height * 0.50)), cert_inn, fill="#1E88E5", font=get_font(int(height * 0.032)), anchor="mm")
    draw.text((center_x, int(height * 0.67)), f"Действителен до {cert_valid}", fill="#4A5568", font=get_font(int(height * 0.025)), anchor="mm")

    pdf_bytes = io.BytesIO()
    img.save(pdf_bytes, format="PDF", resolution=300.0)
    pdf_bytes.seek(0)
    return pdf_bytes

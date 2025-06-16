import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from datetime import date 


def get_student_data(file_path):
    df = pd.read_excel(file_path, header=7)

    selected_columns = df.iloc[:, [0, 1, 9, 12, 13, 14]]
    selected_columns.columns = ["ID", 'F.I.Sh.', 'Yonalish', "diplom_id", "tartib_raqam", "Kvalifikatsiya"]

    selected_columns = selected_columns.dropna(
        how='all',
        subset=["ID", 'F.I.Sh.', 'Yonalish', 'Kvalifikatsiya', 
                "diplom_id", "tartib_raqam"]
    )

    selected_columns = selected_columns.fillna("")

    return selected_columns.to_dict(orient='records')

def date_to_string(date_obj):
    day = date_obj.day
    month = date_obj.month
    year = date_obj.year
    month_names = [
        "Yanvar", "Fevral", "Mart", "Aprel", "May", "Iyun",
        "Iyul", "Avgust", "Sentabr", "Oktabr", "Noyabr", "Dekabr"
    ]
    month_name = month_names[month - 1]
    return f"{year} yil {day}-{month_name.lower()} dagi"


def create_diplom_kuchirma_hujjat(student_data, sana,  fayl_nomi='bitiruvchi_diplom_kuchirma.docx'):
    doc = Document()

    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    section.top_margin = Inches(0.39)
    section.bottom_margin = Inches(0.39)
    section.left_margin = Inches(0.39)
    section.right_margin = Inches(0.39)

    def add_line(cell, text, size=12, bold=False, center=True):
        p = cell.add_paragraph()
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
        run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
        run._element.rPr.rFonts.set(qn('w:hAnsi'), 'Times New Roman')
        run.font.size = Pt(size)
        run.bold = bold
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER if center else WD_PARAGRAPH_ALIGNMENT.LEFT

        p_format = p.paragraph_format
        p_format.line_spacing = Pt(12)
        p_format.space_before = Pt(0)
        p_format.space_after = Pt(0)

    def fill_cell(cell, item):
        add_line(cell, "O‘ZBEKISTON RESPUBLIKASI", 12, bold=True)
        add_line(cell, "BAKALAVR", 12, bold=True)
        add_line(cell, "DIPLOMI", 12, bold=True)
        qaror_id = item['diplom_id'] if item['diplom_id'] != '' else "B №____________"
        add_line(cell, qaror_id, 14)
        add_line(cell, "K O‘ Ch I R M A", 16, bold=True)
        add_line(cell, "SAMARQAND DAVLAT UNIVERSITETI", 14, bold=True)
        add_line(cell, "Davlat attestasiya komissiyasining", 14, bold=True)
        add_line(cell, f"{sana}", 14, bold=True)
        add_line(cell, "qaroriga  binoan", 14, bold=True)
        fish = item['F.I.Sh.'] if item['F.I.Sh.'] != '' else "____________________________________"
        add_line(cell, f"{fish}ga", 14)
        yunalish = item['Yonalish'] if item['Yonalish'] != '' else "____________________________________"
        add_line(cell, f"{yunalish} yo’nalishi bo’yicha", 14)
        add_line(cell, "B A K A L A V R", 16, bold=True)
        add_line(cell, "DARAJASI", 14, bold=True)
        kval = item['Kvalifikatsiya'] if item['Kvalifikatsiya'] != '' else "____________________________________"
        add_line(cell, f"va {kval} kvalifikatsiyasi berildi", 14)
        tartib_raqam = item['tartib_raqam'] if item['tartib_raqam'] != '' else "________________"
        add_line(cell, f"Ro‘yxatga olish raqami {tartib_raqam}", 14, center=False)
        id = item['ID'] if item['ID'] != '' else "________"
        add_line(cell, f"    Ushbu ko‘chirma faqat {id} sonli yo‘llanma bilan o‘z kuchiga ega.", 14, center=False)
        add_line(cell, "Davlat attestatsiya va taqsimot", 14, center=False)
        add_line(cell, "komissiyalari raisi.", 14, center=False)
        add_line(cell, "Rektor                          R.I.Xalmuradov", 14)

    # Talabalarni juftlab sahifalarga joylashtirish
    for i in range(0, len(student_data), 2):
        table = doc.add_table(rows=1, cols=2)
        row = table.rows[0]
        chap, ong = row.cells

        fill_cell(chap, student_data[i])
        if i + 1 < len(student_data):
            fill_cell(ong, student_data[i + 1])

        doc.add_page_break()

    doc.save(fayl_nomi)
    return fayl_nomi




def main(file_path, sana):
    sana_str = date_to_string(sana)
    student_data = get_student_data(file_path)
    output_file = create_diplom_kuchirma_hujjat(student_data, sana_str)
    return output_file

if __name__ == "__main__":
    file_path = 'битирувчи 2024-2025 кит чун АСЛ (2).xlsx'  # Replace with your actual file path
    sana = date(2023, 10, 1)  
    main(file_path, sana)

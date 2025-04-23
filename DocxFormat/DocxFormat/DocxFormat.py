# -*- coding: utf-8 -*-
import sys # даёт доступ к данным из командной строки которую мы запускаем в c++
from docx import Document # доступ к файлам типа .docx
from docx.shared import Pt # доступ к размеру шрифта
from docx.shared import Mm # доступ к миллиметрам
from docx.enum.text import WD_LINE_SPACING # доступ к межстрочному интервалу
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT # доступ к выравниванию текста
from docx2pdf import convert # доступ к конвертации в pdf

def DocxFormat(input_path, output_path, extra_fields=None, use_pdf=None):
    doc = Document(input_path) # создаём переменную на основе docx документа и его расположения

    if not use_pdf:
        sections = doc.sections # .sections хранит информацию о полях
        for section in sections:
            section.top_margin = Mm(20)
            section.bottom_margin = Mm(20)
            section.left_margin = Mm(30)
            section.right_margin = Mm(10)
        # изменяем поля под нужный формат

        for para in doc.paragraphs: # проходим по абзацам
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE # Междустрочный интервал
            para.paragraph_format.first_line_indent = Mm(12.5) # Абзацный отступ
            para.paragraph_format.space_before = Pt(0)  # Убирать отступ сверху
            para.paragraph_format.space_after = Pt(0)   # Убирать отступ снизу
            para.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY # Выравнивание по ширене
            for run in para.runs: # разбиваем абзацы на блоки с одинаковым форматирование
                run.font.name = 'Times New Roman' 
                run.font.size = Pt(14)
        # font хранит в себе настройки шрифта и через разные методы мы их изменяем

    if extra_fields:
        title = Document()

        def add_center(text="", font_size=12):
            p = title.add_paragraph(text)
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            for run in p.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(font_size)

        def add_right(text="", font_size=14):
            p = title.add_paragraph(text)
            p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            for run in p.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(font_size)

        add_center("Министерство науки и высшего образования Российской Федерации")
        add_center("Федеральное государственное бюджетное ")
        add_center("образовательное учреждение высшего образования")
        add_center("«Новгородский государственный университет имени Ярослава Мудрого»")
        add_center(extra_fields[0])
        add_center(" ", 12)
        add_center(extra_fields[1])
        add_center(" ", 14)
        add_center(" ", 14)
        add_center(" ", 14)
        add_center(" ", 14)
        add_center(" ", 14)
        add_center(" ", 14)
        add_center(extra_fields[2], 16)
        add_center(" ", 14)
        add_center(" ", 14)
        add_center(extra_fields[3] + "по учебной дисциплине «" + extra_fields[4] + 
                   "» по специальности " + extra_fields[5] + " " + extra_fields[6], 14)

        title.add_paragraph()
        title.add_paragraph()

        add_right("Руководитель     ")
        add_right("______________/ " + extra_fields[9])
        add_right("«___»_________________2025 г.")
        add_right(" ", 14)
        add_right(" ", 14)
        add_right("Студент группы " + extra_fields[8])
        add_right("______________/ " + extra_fields[7])
        add_right("«___»_________________2025 г.")

        title.add_paragraph()
        title.add_paragraph()
        title.add_paragraph()
        title.add_paragraph()
        title.add_paragraph()
        title.add_paragraph()
        title.add_paragraph()
        title.add_paragraph()
        title.add_paragraph()
        title.add_paragraph()
        title.add_paragraph()

        for element in doc.element.body[:]:
            title.element.body.append(element)
        doc._body.clear_content()
        for element in title.element.body[:]:
            doc.element.body.append(element)
        
    if not use_pdf:
        doc.save(output_path) # метод из Document для сохранения результата

    if use_pdf:
        pdf_path = output_path.rsplit('.', 1)[0] + '.pdf'
        convert(input_path, pdf_path)

if __name__ == "__main__":
    input_path = sys.argv[1]
    output_path = sys.argv[2]
    use_pdf = sys.argv[3].lower() == "true"

    extra_fields = sys.argv[4:] if len(sys.argv) > 4 else None

    DocxFormat(input_path, output_path, extra_fields, use_pdf)
    # argv это список аргументов переданых скипту при запуске
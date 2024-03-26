import openpyxl
from PIL import Image, ImageDraw, ImageFont

#abrir a planilha
workbook_students = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_students = workbook_students['Sheet1']

for i, line in enumerate(sheet_students.iter_rows(min_row = 2)):
    #cada celula que contem a informação
    course_name = line[0].value
    students_name = line[1].value
    participation_role = line[2].value
    start_date = line[3].value
    finish_date = line[4].value
    hours = line[5].value
    certification_date = line[6].value

    font_name = ImageFont.truetype('/tahomabd.ttf', 90)
    default_font = ImageFont.truetype('./tahoma.ttf', 85)
    date_font = ImageFont.truetype('./tahoma.ttf', 60)
    certification_date_font = ImageFont.truetype('./tahomabd.ttf', 45)

    image = Image.open('./certificado_padrao.jpg')
    draw = ImageDraw.Draw(image)
    draw.text((1025, 825), students_name, fill= 'black', font= font_name)
    draw.text((1069, 950), course_name, fill= 'black', font= default_font)
    draw.text((1432, 1064), participation_role, fill= 'black', font= default_font)
    draw.text((1490, 1183), f'{hours} horas', fill= 'black', font= default_font)
    draw.text((733, 1928), finish_date, fill= 'black', font= date_font)
    draw.text((733, 1775), start_date, fill= 'black', font= date_font)
    draw.text((2223, 1940), certification_date, fill= 'black', font= certification_date_font)

    image.save(f'./{students_name}-certificade.png')
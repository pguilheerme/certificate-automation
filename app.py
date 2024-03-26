import openpyxl
from PIL import Image, ImageDraw, ImageFont

#abrir a planilha
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for i, line in enumerate(sheet_alunos.iter_rows(min_row = 2, max_row= 2)):
    #cada celula que contem a informação
    course_name = line[0].value
    alunos_name = line[1].value
    participation_role = line[2].value
    start_date = line[3].value
    finish_date = line[4].value
    hours = line[5].value
    certification_date = line[6].value

    font_name = ImageFont.truetype('/tahomabd.ttf', 90)
    default_font = ImageFont.truetype('./tahoma.ttf', 85)

    image = Image.open('./certificado_padrao.jpg')
    draw = ImageDraw.Draw(image)
    draw.text((1020, 825), alunos_name, fill= 'black', font= font_name)
    draw.text((1064, 950), course_name, fill= 'black', font= default_font)
    # draw.text((1050, 825), participation_role, fill= 'black', font= default_font)
    # draw.text((1050, 825), start_date, fill= 'black', font= default_font)
    # draw.text((1050, 825), finish_date, fill= 'black', font= default_font)
    # draw.text((1050, 825), hours, fill= 'black', font= default_font)
    # draw.text((1050, 825), certification_date, fill= 'black', font= default_font)

    image.save(f'./{alunos_name}-certificade.png')
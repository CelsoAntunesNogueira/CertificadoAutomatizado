import openpyxl
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime

# Abrindo a Planilha
workbook_alunos = openpyxl.load_workbook('planilha.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participacao = linha[2].value

    # Função para formatar datas corretamente
    def formatar_data(valor_celula):
        if isinstance(valor_celula, datetime):
            return valor_celula.strftime('%d/%m/%Y')
        elif isinstance(valor_celula, str):
            return valor_celula  # Se já for string, mantém
        else:
            return "00/00/0000"  # Valor padrão caso esteja vazio ou inválido

    # Aplicando a formatação correta nas datas
    data_inicio = formatar_data(linha[3].value)
    data_final = formatar_data(linha[4].value)
    carga_horaria = linha[5].value if linha[5].value else "0"
    data_emissao = formatar_data(linha[6].value)

    # Definindo a fonte
    fonte_nome = ImageFont.truetype('./tahomabd.ttf', 90)
    fonte_geral = ImageFont.truetype('./tahoma.ttf', 80)
    fonte_data = ImageFont.truetype('./tahoma.ttf', 55)

    # Modelo do certificado
    image = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)

    # Adicionando os textos ao certificado
    desenhar.text((1020, 827), nome_participante, fill='black', font=fonte_nome)
    desenhar.text((1060, 950), nome_curso, fill='black', font=fonte_geral)
    desenhar.text((1435, 1065), nome_participante, fill='black', font=fonte_geral)
    desenhar.text((1480, 1182), str(carga_horaria), fill='black', font=fonte_geral)

    desenhar.text((750, 1770), data_inicio, fill='blue', font=fonte_data)
    desenhar.text((750, 1930), nome_participante, fill='blue', font=fonte_data)
    desenhar.text((2220, 1930), data_emissao, fill='blue', font=fonte_data)

    # Gerando o certificado com numeração e ordem alfabética
    nome_arquivo = f'certificados/{indice}_{nome_participante}_certificado.png'
    image.save(nome_arquivo)

    print(f'Certificado gerado: {nome_arquivo}')

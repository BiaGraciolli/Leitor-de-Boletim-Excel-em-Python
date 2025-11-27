import pandas as pd
from prettytable import PrettyTable
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

def gerarpdf():
    try:
        pdf=canvas.Canvas(f'{arq.iloc[x,0]}.pdf')
        pdf.setTitle(f'Notas de Projetos - {arq.iloc[x,0]}')
        pdf.setFont("Helvetica-Oblique", 20)
        pdf.drawString(50,790,f'{arq.iloc[x,1]}')
        pdf.setFont("Helvetica-Oblique", 15)
        pdf.drawString(50,770,'Notas de Projetos')
        pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
        pdf.drawString(50,720,'INFORMAÇÕES DO ALUNO:')
        pdf.setFont("Arial", 14)
        pdf.drawString(60,700,f'Nome: {arq.iloc[x,1]}')
        pdf.drawString(60,680,f'Matrícula: {arq.iloc[x,0]}')
        pdf.drawString(60,660,f'Turma: {arq.iloc[x,2]}')
        pdf.setFont("Helvetica-Oblique", 15)
        pdf.drawString(50,600,'NOTAS:')
        pdf.setFont("Arial", 14)
        z=580
        listamaterias = list(arq.columns[3:len(arq.columns)-1])
        for materia in listamaterias:
            valor = arq.loc[x, materia]
            if float(valor)>=6:
                pdf.setFillColor('blue')
            else:
                pdf.setFillColor('red')
            if str(valor)=='' or str(valor).lower()=='nan':
                valor='NÃO INFORMADA'
                pdf.setFillColor('red')
            pdf.drawString(60,z,f'{materia}: {valor}')
            z+=-20
        pdf.setFillColor('black')
        pdf.setFont("Helvetica-Oblique", 15)
        pdf.drawString(50,z-40,'MÉDIA FINAL:')
        pdf.setFont("Arial", 14)
        if str(arq.iloc[x,len(arq.columns)-1])=='' or str(arq.iloc[x,len(arq.columns)-1]).lower()=='nan':
                media='NÃO CALCULADA'
                pdf.setFillColor('red')
        else:
            media=arq.iloc[x,len(arq.columns)-1]
            if media>=6:
                pdf.setFillColor('blue')
            else:
                pdf.setFillColor('red')
        pdf.drawString(60,z-60,f'{media}')         
        pdf.save()
    except Exception as error:
        print(f'---> Erro - {error}\n\n')

while True:
    try:
        diretorio=input('Digite o caminho completo do arquivo Excel\n(Exemplo: E:/notas/planilhanotas.xlsx) ou S para sair do programa:').strip()
        if diretorio.upper()=='S':
            break
        planilha=input('Digite o nome da planilha ou S para sair:').strip()
        if planilha.upper()=='S':
            break
        arq = pd.read_excel(diretorio, sheet_name=planilha)
    except Exception as error:
        print(f'---> Erro - {error}\n\n')
    else:
        print(f'Arquivo encontrado!\n')
        break

if(diretorio!='S' and planilha !='S'):
    for x in range(0, len(arq)):
        linha=x
        listamaterias=list(arq.columns[3:len(arq.columns)-1])
        textomaterias = ''
        for materia in listamaterias:
            valor = arq.loc[x, materia]
            if float(valor)>=6:
                cor="#09669c"
            else:
                cor="#ab1103"
            if str(valor)=='' or str(valor).lower()=='nan':
                valor='NÃO INFORMADA'
                cor="#ab1103"
            textomaterias += f'''
            <div class="classeConfQuadrado">
                <div class="classeMateria"><strong>{materia}</strong></div>
                <div class="classeValor" style="color:{cor};">{valor}</div>
            </div>'''
            if str(arq.iloc[x,len(arq.columns)-1])=="" or str(arq.iloc[x,len(arq.columns)-1]).lower()=='nan':
                media='NÃO CALCULADA'
                cor="#ab1103"
            else:
                media=arq.iloc[x,len(arq.columns)-1]
                if media>=6:
                    cor="#09669c"
                else:
                    cor="#ab1103"
        try:
            gerarpdf()
            site=open(f'{arq.iloc[x,0]}.html','w')
            site.write(f'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="windows-1252">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Notas de Projetos - {arq.iloc[x,0]}</title>
    <style>
        *{{
            margin:0;
            padding:0;
            font-family:'Segoe UI';
        }}
        body{{
            background-color:#f5f7fa;
            padding:20px;
        }}
        header{{
            background:#230e4a;
            color:white;
            padding:30px;
            text-align:center;
        }}
        h1{{
            font-size:26.4pt;
            margin-bottom:10px;
        }}
        h3{{
            color:#230e4a;
            margin-bottom:10px;
            }}
        .classeBloco{{
            width:1000px;
            margin:0px auto;
            background:white;
        }}
        .classeEspacamento{{
            padding:30px;
        }}
        .classeInformacoes{{
            background:#f5f7fa;
            padding:20px;
            border-left-width:4px;
            border-left-style:solid;
            border-left-color:#230e4a;
        }}
        .classeQuadradoNota{{
            font-size:18pt;
            color:#230e4a;
            margin-bottom:20px;
            padding-bottom:10px;
        }}
        .classeConfQuadrado{{
            background:white;
            border-radius:8px;
            padding:20px;
            text-align:center;
            border-bottom-width:2px;
            border-bottom-style:solid;
            border-bottom-color:#eaeaea;
        }}
        .classeMateria{{
            font-size:19px;
            color:#614870;
            margin-bottom:10px;
        }}
        .classeValor{{
            font-size:30px;
            font-weight:bold;
        }}
        .classeMedia{{
            background:{cor};
            color:white;
            padding:25px;
            text-align:center;
            margin-top:20px;
        }}
        .classeNomeMedia{{
            font-size:19.2px;
            margin-bottom:10px;
        }}
        .classeMediaNota{{
            font-size:40px;
            font-weight:bold;
        }}
        .classeInfomacoes{{
            background:#f8f9fa;
            padding:25px;
            margin-top:30px;
        }}
        .classeTextoInformacoes{{
            padding:10px;
            border-bottom-width:1px;
            border-bottom-style:solid;
            border-bottom-color:#eaeaea;
        }}
        .containerBotaoPDF {{
            text-align: center;
            margin-top: 50px;
            margin-bottom: 30px;
        }}
        .botaoPDF {{
            background: #230e4a;
            color: white;
            padding: 18px 60px;
            font-size: 24px;
            font-weight: bold;
        }}
    </style>
</head>
<body>
    <div class="classeBloco">
        <header>
            <h1>{arq.iloc[x,1]}</h1>
            <div><h4>Notas de Projetos</h4></div>
        </header>
        <div class="classeEspacamento">
            <div>
                <div class="classeInformacoes">
                    <h3>Informações</h3>
                    <p><strong>Nome:</strong> {arq.iloc[x,1]}</p>
                    <p><strong>Matrícula:</strong> {arq.iloc[x,0]}</p>
                    <p><strong>Turma:</strong> {arq.iloc[x,2]}</p>
                </div>
            </div>
        <div>
            <h2 class="classeNota"><br/>Notas:</h2>
            <div class="classeQuadradoNota">
                    {textomaterias}
            </div>
                <div class="classeMedia">
                    <div class="classeNomeMedia">Média Final</div>
                    <div class="classeMediaNota">{media}</div>
                </div>
            </div>
            <div class="classeInfomacoes">
                <h2 class="section-title">Informações Adicionais</h2>
                <div class="classeTextoInformacoes">
                    <p>• Notas não inseridas: O aluno não entregou o projeto da disciplina ou o professor não adicionou a nota.<br /></p>
                    <p>• A média AT gerada será lançada no boletim para todas as disciplinas de informática para esse bimestre.<br /></p>
                    <p>• Qualquer dúvida procure o professor da disciplina para maiores esclarecimentos.<br /></p>
                </div>
            </div>
            <div class="containerBotaoPDF">
                <a href="{arq.iloc[x,0]}.pdf" target="_blank">
                    <button class="botaoPDF">Abrir PDF</button>
                </a>
            </div>
        </div>
    </div>
</body>
</html>
''')
            site.close()
        except Exception as error:
            print(f'---> Erro - {error}\n\n')
        print(f'Numero da linha do excel:{linha+1}')
        listacolunas=list(arq.columns)
        grid=PrettyTable(listacolunas)
        linha = list(arq.iloc[x])
        tabela=list()
        for valores in linha:
            if str(valores)=="" or str(valores).lower()=='nan':
                tabela.append('NÃO INFORMADA')
            else:
                tabela.append(valores)
        grid.add_row(tabela)
        print(grid)
        print('Arquivo HTML gerado com sucesso!\n')

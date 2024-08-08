from flask import Flask, request, render_template, send_file
import pandas as pd
from io import BytesIO
import unicodedata
import re

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

def remove_accents_and_special_characters(df1, df2):
    # Normalizar texto para decompor os caracteres acentuados
    nfkd_form = unicodedata.normalize('NFKD', df1, df2)
    # Remover caracteres diacríticos (acentos)
    text_without_accents = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
    # Remover caracteres especiais, mantendo apenas letras e números
    text_normalized = re.sub(r'[^a-zA-Z0-9\s]', '', text_without_accents)
    return text_normalized

@app.route('/transferencia', methods=['POST'])
def transferencia():
    arquivo1 = request.files['arquivo1']
    arquivo2 = request.files['arquivo2']
    
    coluna_comum_arquivo1 = request.form['coluna_comum_arquivo1']
    coluna_comum_arquivo2 = request.form['coluna_comum_arquivo2']
    
    df1 = pd.read_excel(arquivo1, engine='openpyxl')
    df2 = pd.read_excel(arquivo2, engine='openpyxl')
    
    print("DataFrame 1 carregado:")
    print(df1.head())
    print("DataFrame 2 carregado:")
    print(df2.head())    

    df1 = df1.rename(columns={coluna_comum_arquivo1: 'comum'})
    df2 = df2.rename(columns={coluna_comum_arquivo2: 'comum'})

    remove_accents_and_special_characters(df1, df2)

    if 'comum' not in df1.columns or 'comum' not in df2.columns:
        return "Erro: A coluna comum não existe em um dos arquivos."
    
    cidades_comum = df2['comum'].dropna().unique()

    print("Cidades em Comum no Segundo Arquivo:")
    print(cidades_comum)

    df1_filtrado = df1[df1['comum'].isin(cidades_comum)]
    df2_filtrado = df2[df2['comum'].isin(cidades_comum)]

    print("Data Frame 1 Filtrado:")
    print(df1_filtrado)
    print("Data Frame 2 Filtrado:")
    print(df2_filtrado)

    df_resultante = pd.merge(df1_filtrado, df2_filtrado, on='comum', how='outer')

    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_resultante.to_excel(writer, sheet_name='Dados Completos', index=False)
    
    output.seek(0)

    return send_file(output, download_name='dados_completos.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
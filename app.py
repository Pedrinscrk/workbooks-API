from flask import Flask, request, render_template, send_file
import pandas as pd
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
RESULTS_FOLDER = "resultados"

# Configura os diretórios de upload e resultados
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULTS_FOLDER, exist_ok=True)

@app.route("/")
def index():
    """Rota principal que exibe o formulário."""
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload_files():
    """Recebe os arquivos enviados e organiza as planilhas."""
    # Limpa a pasta de uploads
    for f in os.listdir(UPLOAD_FOLDER):
        os.remove(os.path.join(UPLOAD_FOLDER, f))

    # Salva os arquivos enviados
    if "folder" in request.files:
        uploaded_files = request.files.getlist("folder")
        for file in uploaded_files:
            filename = secure_filename(file.filename)
            file.save(os.path.join(UPLOAD_FOLDER, filename))

    # Processa as regras enviadas
    regras = request.form.get("regras")
    processa_planilhas(regras)

    # Envia o arquivo final para download
    resultado_final = os.path.join(RESULTS_FOLDER, "planilha_organizada.xlsx")
    return send_file(resultado_final, as_attachment=True)

def processa_planilhas(regras):
    """Processa as planilhas na pasta de uploads e cria a planilha organizada."""
    # Lê todas as planilhas na pasta de uploads
    arquivos = [os.path.join(UPLOAD_FOLDER, f) for f in os.listdir(UPLOAD_FOLDER) if f.endswith(('.xlsx', '.xls'))]
    df_final = pd.DataFrame()

    for arquivo in arquivos:
        df = pd.read_excel(arquivo, header=None)  # Lê o Excel sem definir cabeçalho

        # Localiza dinamicamente o cabeçalho correto
        for i, row in df.iterrows():
            if 'Código' in row.values:  # Procura a palavra 'Código' como referência
                df.columns = df.iloc[i]  # Define essa linha como cabeçalho
                df = df[i + 1:].reset_index(drop=True)  # Remove as linhas acima do cabeçalho
                break

        # Padroniza os nomes das colunas (remove espaços e converte para minúsculas)
        df.columns = df.columns.str.strip().str.lower()
        print("Colunas ajustadas:", df.columns)  # Mostra os nomes das colunas no terminal

        df_final = pd.concat([df_final, df], ignore_index=True)

    # Ajusta o comportamento com base nas regras
    colunas_desejadas = [col.strip().lower() for col in regras.split(",") if col.strip()]
    colunas_existentes = [col for col in colunas_desejadas if col in df_final.columns]

    if not colunas_existentes:
        raise ValueError(f"Nenhuma coluna válida foi encontrada. Colunas disponíveis: {df_final.columns}")

    df_final = df_final[colunas_existentes]

    # Salva a planilha final na pasta de resultados
    resultado_final = os.path.join(RESULTS_FOLDER, "planilha_organizada.xlsx")
    df_final.to_excel(resultado_final, index=False)

if __name__ == "__main__":
    app.run(debug=True)

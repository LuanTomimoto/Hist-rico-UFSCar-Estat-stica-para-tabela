import pdfplumber
import re
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import requests
from io import BytesIO

# Mapeamento dos nomes das disciplinas para abreviações
DISCIPLINE_MAPPING = {
    "Cálculo 1": "C1",
    "Fundamentos de Probabilidade": "FP",
    "Geometria Analítica": "GA",
    "Introdução à Estatística": "IE",
    "Programação 1": "PA1",
    "Análise Descritiva": "ADED",
    "Cálculo de Várias Variáveis": "CVV",
    "Probabilidade 1": "P1",
    "Álgebra Linear": "AL",
    "Álgebra Linear 1": "AL",
    "Introdução à Inferência Estatística": "IIE",
    "Probabilidade 2": "P2",
    "Séries e Equações Diferenciais": "SED",
    "Cálculo Numérico": "CN",
    "Amostragem": "Am",
    "Inferência Estatística": "INF",
    "Processos Estocásticos": "ProcE",
    "Programação Estatística": "ProgE",
    "Teoria de Matrizes": "TM",
    "Análise de Regressão": "AR",
    "Estatística Bayesiana": "EB",
    "Multivariada 1": "M1",
    "Estatística Multivariada 1": "M1",
    "Planejamento": "Pl",
    "Séries Temporais": "ST",
    "Programação 2": "PA2",
    "Análise de Sobrevivência": "AS",
    "Multivariada 2": "M2",
    "Estatística Multivariada 2": "M2",
    "Não Paramétrica": "NP",
    "Métodos Computacionais Intensivos": "MCI",
    "Modelos Lineares Generalizados": "MLG",
    "Mineração de Dados": "MD",
    "Estatística 1": "Estat1",
    "Estatística 2": "Estat2",
    "Probabilidade A": "ProbA",
    "Probabilidade B": "ProbB",
    "Probabilidade C": "ProbC",
    "Teoria das Organizações": "TeoriaOrg",
    "Computação Básica": "CompBasic",
    "Estatística Computacional A": "EstatCompA",
     "Estatística Computacional B": "EstatCompB",
    "Introdução à Computação": "IntrodComp",
    "Programação Científica": "ProgCie",
    "Inferência Estatística A": "InfA",
}

# Função para extrair dados de um PDF
def extract_data_from_pdf(pdf_path):
    data = []
    current_semester = None
    student_code = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            # Extrai o código do aluno (RA)
            if not student_code:
                ra_match = re.search(r'Aluno:\s*(\d{6})', text)
                if ra_match:
                    student_code = ra_match.group(1)

            # Processa cada linha
            for line in text.split('\n'):
                # Extrai o semestre (ex: 2012/1)
                semester_match = re.match(r'(\d{4}/\d)', line)
                if semester_match:
                    current_semester = semester_match.group(1)
                    continue

                # Extrai o código, nome e turma da disciplina
                discipline_match = re.match(r'(\d+)\s*-\s*([\w\sçãõáéíóúâêîôûàèìòùäëïöü]+)\s*-\s*([A-Z])', line)
                if discipline_match:
                    discipline_code = discipline_match.group(1)
                    discipline_name = discipline_match.group(2).strip()
                    discipline_class = discipline_match.group(3)

                    # Mapeia o nome da disciplina para a abreviação correspondente
                    normalized_name = discipline_name.strip()
                    discipline_abbr = next(
                        (abbr for name, abbr in DISCIPLINE_MAPPING.items() if name.lower() == normalized_name.lower()),
                        normalized_name
                    )

                    # Extrai a avaliação (nota) e a frequência
                    details = re.findall(r'\d+,\d+|\d+', line)
                    avaliacao = float(details[-2].replace(',', '.')) if len(details) >= 2 else None
                    freq = float(details[-1].replace(',', '.')) if len(details) >= 1 else None

                    # Determina o status de aprovação
                    if freq is not None and avaliacao is not None:
                        if freq < 75 and avaliacao < 6:
                            aprovado = "REPROVADO POR NOTA E FREQUÊNCIA"
                        elif freq < 75:
                            aprovado = "REPROVADO POR FREQUÊNCIA"
                        elif avaliacao < 6:
                            aprovado = "REPROVADO POR NOTA"
                        else:
                            aprovado = "APROVADO"
                    else:
                        aprovado = None

                    # Adiciona os dados à lista
                    data.append({
                        "Estudante": student_code,
                        "Disciplina": discipline_abbr,
                        "Aprovado": aprovado,
                        "Nota": avaliacao,
                        "Frequência": freq,
                        "Ano": current_semester,
                        "Situação": ""
                    })

    return data

# Função para salvar os dados em um arquivo Excel
def save_to_excel(data, output_path):
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False)

# Função para baixar um PDF de uma URL do Google Drive
def download_pdf_from_drive(url):
    file_id = url.split('/d/')[1].split('/')[0]  # Extrai o ID do arquivo da URL
    download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
    response = requests.get(download_url)
    return BytesIO(response.content)

# Função principal para processar PDFs
def process_pdfs(pdf_path, output_path):
    if pdf_path.startswith("http"):
        pdf_file = download_pdf_from_drive(pdf_path)
        data = extract_data_from_pdf(pdf_file)
    else:
        if os.path.isdir(pdf_path):
            pdf_files = [os.path.join(pdf_path, f) for f in os.listdir(pdf_path) if f.endswith(".pdf")]
            all_data = []
            for pdf_file in pdf_files:
                data = extract_data_from_pdf(pdf_file)
                all_data.extend(data)
            save_to_excel(all_data, output_path)
        else:
            data = extract_data_from_pdf(pdf_path)
            save_to_excel(data, output_path)
    messagebox.showinfo("Sucesso", f"Dados salvos em {output_path}")

# Interface gráfica
def create_gui():
    def browse_pdf():
        path = filedialog.askdirectory() if folder_var.get() else filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        pdf_path_entry.delete(0, tk.END)
        pdf_path_entry.insert(0, path)

    def browse_output():
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        output_path_entry.delete(0, tk.END)
        output_path_entry.insert(0, path)

    def start_processing():
        pdf_path = pdf_path_entry.get()
        output_path = output_path_entry.get()
        if not pdf_path or not output_path:
            messagebox.showerror("Erro", "Por favor, insira os caminhos corretamente.")
            return
        process_pdfs(pdf_path, output_path)

    root = tk.Tk()
    root.title("Processador de Históricos Escolares")

    # Variável para controle de pasta/arquivo
    folder_var = tk.BooleanVar(value=False)

    # Campos de entrada
    tk.Label(root, text="Caminho do PDF ou Pasta de PDFs:").grid(row=0, column=0, padx=5, pady=5)
    pdf_path_entry = tk.Entry(root, width=50)
    pdf_path_entry.grid(row=0, column=1, padx=5, pady=5)
    tk.Button(root, text="Procurar", command=browse_pdf).grid(row=0, column=2, padx=5, pady=5)

    tk.Label(root, text="Caminho de Saída do Excel:").grid(row=1, column=0, padx=5, pady=5)
    output_path_entry = tk.Entry(root, width=50)
    output_path_entry.grid(row=1, column=1, padx=5, pady=5)
    tk.Button(root, text="Procurar", command=browse_output).grid(row=1, column=2, padx=5, pady=5)

    # Checkbox para selecionar pasta
    tk.Checkbutton(root, text="Selecionar pasta de PDFs", variable=folder_var).grid(row=2, column=1, padx=5, pady=5)

    # Botão para iniciar o processamento
    tk.Button(root, text="Processar", command=start_processing).grid(row=3, column=1, padx=5, pady=10)

    # Adiciona a informação de desenvolvedor no rodapé
    tk.Label(root, text="Desenvolvido por Luan Tomimoto, PET Estatística UFSCAR", font=("Arial", 8), fg="gray").grid(row=4, column=0, columnspan=3, padx=5, pady=5)

    root.mainloop()


# Inicia a interface gráfica
create_gui()

from pdf2docx import Converter
from docx import Document
from deep_translator import GoogleTranslator
from docx2pdf import convert
import os
from tkinter import Tk, filedialog, messagebox, StringVar, Label, Button, Entry, Radiobutton, IntVar


# === Etapa 1: Converter PDF em DOCX preservando layout ===
def pdf_para_docx(pdf_path):
    docx_path = pdf_path.replace(".pdf", ".docx")
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()
    return docx_path


# === Etapa 2: Traduzir o texto dentro do DOCX ===
def traduzir_docx(docx_path, idioma_destino):
    tradutor = GoogleTranslator(source="auto", target=idioma_destino)
    doc = Document(docx_path)

    for p in doc.paragraphs:
        if p.text.strip():
            try:
                p.text = tradutor.translate(p.text)
            except Exception as e:
                print(f"⚠️ Erro traduzindo trecho: {e}")
                continue

    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                if celula.text.strip():
                    try:
                        celula.text = tradutor.translate(celula.text)
                    except Exception as e:
                        print(f"⚠️ Erro traduzindo célula: {e}")
                        continue

    novo_docx = docx_path.replace(".docx", f"_{idioma_destino}.docx")
    doc.save(novo_docx)
    return novo_docx


# === Etapa 3: Converter DOCX traduzido em PDF ===
def docx_para_pdf(docx_path):
    pasta = os.path.dirname(docx_path)
    convert(docx_path, pasta)
    return docx_path.replace(".docx", ".pdf")


# === Função principal ===
def traduzir_pdf_layout_total(pdf_path, idioma_destino):
    docx_temp = pdf_para_docx(pdf_path)
    traduzido_docx = traduzir_docx(docx_temp, idioma_destino)
    saida_pdf = docx_para_pdf(traduzido_docx)
    print(f"✅ PDF final salvo em: {saida_pdf}")
    return saida_pdf


# === Interface gráfica ===
def selecionar_arquivo():
    arquivo = filedialog.askopenfilename(title="Selecione um PDF", filetypes=[("Arquivos PDF", "*.pdf")])
    caminho_var.set(arquivo)


def traduzir():
    caminho = caminho_var.get().strip()
    idioma = idioma_var.get().strip()
    if not caminho or not idioma:
        messagebox.showerror("Erro", "Selecione o PDF e o idioma de destino (ex: en, es, fr, it).")
        return

    try:
        saida = traduzir_pdf_layout_total(caminho, idioma)
        messagebox.showinfo("Concluído", f"Tradução concluída com sucesso!\nArquivo salvo em:\n{saida}")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha durante a tradução:\n{e}")


root = Tk()
root.title("Tradutor de PDFs — SERTA ⚙️ (Layout Perfeito via Word)")
root.geometry("560x300")
root.resizable(False, False)

caminho_var = StringVar()
idioma_var = StringVar()

Label(root, text="Selecione o PDF:", font=("Arial", 11)).pack(pady=10)
Entry(root, textvariable=caminho_var, width=65).pack()
Button(root, text="Escolher PDF", command=selecionar_arquivo).pack(pady=5)

Label(root, text="Idioma destino (ex: en, es, fr, it):", font=("Arial", 11)).pack(pady=10)
Entry(root, textvariable=idioma_var, width=15, font=("Arial", 12)).pack()

Button(root, text="Traduzir", command=traduzir, bg="#4CAF50", fg="white", font=("Arial", 12, "bold")).pack(pady=20)

root.mainloop()

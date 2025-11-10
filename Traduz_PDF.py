from pdf2docx import Converter, parse
from docx import Document
from deep_translator import GoogleTranslator
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os, shutil, tempfile
from pathlib import Path
from tkinter import Tk, filedialog, messagebox, StringVar, Label, Button, Entry
import win32com.client as win32


# === Etapa 1: Converter PDF → DOCX ===
def pdf_para_docx(pdf_path):
    docx_path = pdf_path.replace(".pdf", ".docx")
    try:
        print("➡️ Iniciando conversão PDF → DOCX...")
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()
        print(f"✅ Conversão concluída: {docx_path}")
    except Exception as e:
        print(f"⚠️ Conversão direta falhou ({e}), tentando método alternativo...")
        try:
            parse(pdf_path, docx_path, start=0, end=None)
            print(f"✅ Conversão alternativa concluída: {docx_path}")
        except Exception as e2:
            raise RuntimeError(f"Erro crítico ao converter PDF → DOCX: {e2}")
    return docx_path


# === Etapa 2: Traduzir o DOCX preservando layout e formatação ===
def traduzir_docx(docx_path, idioma_destino):
    tradutor = GoogleTranslator(source="auto", target=idioma_destino)
    doc = Document(docx_path)

    # Estilo base global
    for style in doc.styles:
        if style.type == 1:  # Parágrafos
            style.font.name = "Arial"
            style.font.size = Pt(10)

    # Traduz parágrafos
    for p in doc.paragraphs:
        texto = p.text.strip()
        if texto:
            try:
                traduzido = tradutor.translate(texto)
                p.text = traduzido
                p.style = doc.styles["Normal"]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.0
            except Exception as e:
                print(f"⚠️ Erro traduzindo parágrafo: {e}")
                continue

    # Traduz tabelas preservando proporções
    for tabela in doc.tables:
        tabela.autofit = True
        tabela.allow_autofit = True
        tabela.style = "Table Grid"
        for linha in tabela.rows:
            for celula in linha.cells:
                celula.vertical_alignment = 1
                for par in celula.paragraphs:
                    par.style = doc.styles["Normal"]
                    par.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    par.paragraph_format.space_before = Pt(0)
                    par.paragraph_format.space_after = Pt(2)
                    par.paragraph_format.line_spacing = 1.0
                texto_celula = celula.text.strip()
                if texto_celula:
                    try:
                        celula.text = tradutor.translate(texto_celula)
                    except Exception as e:
                        print(f"⚠️ Erro traduzindo célula: {e}")
                        continue

    novo_docx = docx_path.replace(".docx", f"_{idioma_destino}.docx")
    doc.save(novo_docx)
    print(f"✅ DOCX traduzido salvo: {novo_docx}")
    return novo_docx


# === Etapa 3: Converter DOCX → PDF via Microsoft Word COM (estável) ===
def docx_para_pdf(docx_path):
    original_path = Path(docx_path)
    pasta_destino = original_path.parent
    saida_pdf = pasta_destino / (original_path.stem + ".pdf")

    print("➡️ Convertendo DOCX para PDF via Microsoft Word COM...")

    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(str(original_path))
        doc.SaveAs(str(saida_pdf), FileFormat=17)  # 17 = wdFormatPDF
        doc.Close(False)
        word.Quit()
        print(f"✅ PDF final salvo em: {saida_pdf}")
        return str(saida_pdf)

    except Exception as e:
        raise RuntimeError(f"Erro ao converter DOCX → PDF via Word: {e}")


# === Função principal ===
def traduzir_pdf_layout_total(pdf_path, idioma_destino):
    print(f"🔄 Iniciando tradução de: {pdf_path}")
    docx_temp = pdf_para_docx(pdf_path)
    traduzido_docx = traduzir_docx(docx_temp, idioma_destino)
    saida_pdf = docx_para_pdf(traduzido_docx)
    return traduzido_docx, saida_pdf


# === Interface Tkinter ===
def selecionar_arquivo():
    arquivo = filedialog.askopenfilename(
        title="Selecione um PDF",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    caminho_var.set(arquivo)


def traduzir():
    caminho = caminho_var.get().strip()
    idioma = idioma_var.get().strip()

    if not caminho or not idioma:
        messagebox.showerror("Erro", "Selecione o PDF e informe o idioma destino (ex: en, es, fr, it).")
        return

    try:
        docx_trad, pdf_trad = traduzir_pdf_layout_total(caminho, idioma)
        messagebox.showinfo(
            "Tradução concluída ✅",
            f"Arquivos gerados:\n\n📄 DOCX traduzido:\n{docx_trad}\n\n"
            f"📄 PDF traduzido:\n{pdf_trad}\n\n"
            f"Ambos foram salvos na mesma pasta do original.\n"
            f"O arquivo DOCX foi mantido para revisão manual."
        )
    except Exception as e:
        messagebox.showerror("Erro", f"Falha durante a tradução:\n{e}")


# === GUI ===
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

Button(
    root, text="Traduzir", command=traduzir,
    bg="#4CAF50", fg="white", font=("Arial", 12, "bold")
).pack(pady=20)

root.mainloop()

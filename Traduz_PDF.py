from pdf2docx import Converter
import docx
from docx import Document
from deep_translator import GoogleTranslator
from docx2pdf import convert
from docx.shared import Pt
from docx.oxml.ns import qn
import os
from tkinter import Tk, filedialog, messagebox, StringVar, Label, Button, Entry


# ============================================================
#  ETAPA 1 — Converter PDF → DOCX preservando layout original
# ============================================================
def pdf_para_docx(pdf_path):
    docx_path = pdf_path.replace(".pdf", ".docx")
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()
    return docx_path


# ============================================================
#  ETAPA 2 — Traduzir texto mantendo layout de TABELAS IGUAL
# ============================================================
def traduzir_docx(docx_path, idioma_destino):
    tradutor = GoogleTranslator(source="auto", target=idioma_destino)
    doc = Document(docx_path)

    # =============================
    #  MARGENS DISCRETAS (não alteram layout)
    # =============================
    CM = 360000
    for section in doc.sections:
        section.top_margin = int(0.5 * CM)
        section.bottom_margin = int(0.5 * CM)
        section.left_margin = int(0.8 * CM)
        section.right_margin = int(0.8 * CM)

    # =============================
    #  TRADUÇÃO DE PARÁGRAFOS NORMAIS
    # =============================
    for p in doc.paragraphs:
        texto = p.text.strip()
        if texto:
            try:
                p.text = tradutor.translate(texto)
            except:
                pass

    # =============================
    #  TABELAS — manter layout + evitar sobreposição de linhas
    # =============================
    for tabela in doc.tables:
        for row in tabela.rows:

            # === Permite que a ALTURA DA LINHA cresça automaticamente ===
            trPr = row._tr.get_or_add_trPr()
            autoHeight = docx.oxml.parse_xml(
                r"<w:trHeight w:hRule='auto' "
                r"xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'/>"
            )
            trPr.append(autoHeight)

            for celula in row.cells:

                # === Permitir quebra de linha natural dentro da célula ===
                tcPr = celula._tc.get_or_add_tcPr()
                wrap_xml = docx.oxml.parse_xml(
                    r"<w:noWrap w:val='0' "
                    r"xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'/>"
                )
                tcPr.append(wrap_xml)

                # === Padding leve para afastar texto das linhas (evita sobreposição) ===
                padding = docx.oxml.parse_xml(
                    r"<w:tcMar xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>"
                    r"<w:top w:w='30' w:type='dxa'/>"     # ~0.5 mm
                    r"<w:left w:w='15' w:type='dxa'/>"
                    r"<w:bottom w:w='30' w:type='dxa'/>"
                    r"<w:right w:w='15' w:type='dxa'/>"
                    r"</w:tcMar>"
                )
                tcPr.append(padding)

                # === Formatação fina ===
                for para in celula.paragraphs:
                    para.paragraph_format.line_spacing = 1
                    para.paragraph_format.space_after = 0

                    # Fonte ideal para caber sem estragar layout
                    for run in para.runs:
                        run.font.size = Pt(7.5)

                # === TRADUÇÃO DO TEXTO DA CÉLULA ===
                texto_cel = celula.text.strip()
                if texto_cel:
                    try:
                        celula.text = tradutor.translate(texto_cel)
                    except:
                        pass

    # =============================
    #  SALVAR RESULTADO FINAL
    # =============================
    novo_docx = docx_path.replace(".docx", f"_{idioma_destino}.docx")
    doc.save(novo_docx)

    return novo_docx

# ============================================================
#  ETAPA 3 — Converter DOCX traduzido → PDF
# ============================================================
def docx_para_pdf(docx_path):
    pasta = os.path.dirname(docx_path)
    convert(docx_path, pasta)
    return docx_path.replace(".docx", ".pdf")


# ============================================================
#  FLUXO PRINCIPAL
# ============================================================
def traduzir_pdf_layout_total(pdf_path, idioma_destino):
    docx_temp = pdf_para_docx(pdf_path)
    traduzido_docx = traduzir_docx(docx_temp, idioma_destino)
    saida_pdf = docx_para_pdf(traduzido_docx)
    print(f"✅ PDF final salvo em: {saida_pdf}")
    return saida_pdf


# ============================================================
#  INTERFACE TKINTER
# ============================================================
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
        messagebox.showerror("Erro", "Selecione o PDF e o idioma de destino.")
        return

    try:
        saida = traduzir_pdf_layout_total(caminho, idioma)
        messagebox.showinfo("Concluído", f"Tradução concluída!\nPDF salvo em:\n{saida}")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha durante a tradução:\n{e}")


# ============================================================
#  JANELA
# ============================================================
root = Tk()
root.title("Tradutor de PDFs — SERTA ⚙️ (Layout Preservado)")
root.geometry("560x300")
root.resizable(False, False)

caminho_var = StringVar()
idioma_var = StringVar()

Label(root, text="Selecione o PDF:", font=("Arial", 11)).pack(pady=10)
Entry(root, textvariable=caminho_var, width=65).pack()
Button(root, text="Escolher PDF", command=selecionar_arquivo).pack(pady=5)

Label(root, text="Idioma destino (ex: en, es, fr, it):", font=("Arial", 11)).pack(pady=10)
Entry(root, textvariable=idioma_var, width=15, font=("Arial", 12)).pack()

Button(root, text="Traduzir", command=traduzir,
       bg="#4CAF50", fg="white", font=("Arial", 12, "bold")).pack(pady=20)

root.mainloop()

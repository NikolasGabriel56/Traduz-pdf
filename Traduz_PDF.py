from pdf2docx import Converter
import docx
from docx import Document
from deep_translator import GoogleTranslator
from docx2pdf import convert
from docx.shared import Pt
from docx.oxml.ns import qn
import os
from tkinter import Tk, filedialog, messagebox, StringVar, Label, Button, Entry



def pdf_para_docx(pdf_path):
    docx_path = pdf_path.replace(".pdf", ".docx")
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()
    return docx_path



def traduzir_docx(docx_path, idioma_destino):
    tradutor = GoogleTranslator(source="auto", target=idioma_destino)
    doc = Document(docx_path)

    CM = 360000
    for section in doc.sections:
        section.top_margin = int(0.5 * CM)
        section.bottom_margin = int(0.5 * CM)
        section.left_margin = int(0.8 * CM)
        section.right_margin = int(0.8 * CM)

    # --------------------------
    # PARÁGRAFOS FORA DAS TABELAS
    # --------------------------
    for p in doc.paragraphs:
        if not p.text.strip():
            continue

        try:
            novo_texto = tradutor.translate(p.text)
        except:
            novo_texto = p.text

        # salvar primeiro run
        if p.runs:
            ref = p.runs[0]
            estilo = {
                "font_name": ref.font.name,
                "font_size": ref.font.size,
                "bold": ref.bold,
                "italic": ref.italic,
                "underline": ref.underline
            }
        else:
            estilo = None

        p.clear()
        run = p.add_run(novo_texto)

        if estilo:
            run.font.name = estilo["font_name"]
            run.font.size = estilo["font_size"]
            run.bold = estilo["bold"]
            run.italic = estilo["italic"]
            run.underline = estilo["underline"]

        p.paragraph_format.line_spacing = 1
        p.paragraph_format.space_after = 0

    # --------------------------
    # TABELAS — preservação total
    # --------------------------
    for tabela in doc.tables:
        for row in tabela.rows:

            trPr = row._tr.get_or_add_trPr()
            autoHeight = docx.oxml.parse_xml(
                r"<w:trHeight w:hRule='auto' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'/>"
            )
            trPr.append(autoHeight)

            for celula in row.cells:
                for para in celula.paragraphs:

                    texto_original = para.text.strip()
                    if not texto_original:
                        continue

                    try:
                        texto_trad = tradutor.translate(texto_original)
                    except:
                        texto_trad = texto_original

                    # salvar estilo do 1º run
                    ref_style = None
                    if para.runs:
                        r = para.runs[0]
                        ref_style = {
                            "font_name": r.font.name,
                            "font_size": r.font.size,
                            "bold": r.bold,
                            "italic": r.italic,
                            "underline": r.underline
                        }

                    para.clear()

                    # cria run único
                    rnew = para.add_run(texto_trad)

                    # aplica estilo base
                    if ref_style:
                        rnew.font.name = ref_style["font_name"]
                        rnew.font.size = ref_style["font_size"]
                        rnew.bold = ref_style["bold"]
                        rnew.italic = ref_style["italic"]
                        rnew.underline = ref_style["underline"]

                    # ------------------------------
                    #   CORREÇÃO SUAVE DO NEGRITO
                    # ------------------------------
                    if ":" in texto_original:
                        prefixo, sufixo = texto_trad.split(":", 1)

                        para.clear()

                        # prefixo em negrito
                        bold_run = para.add_run(prefixo + ": ")
                        if ref_style:
                            bold_run.font.name = ref_style["font_name"]
                            bold_run.font.size = ref_style["font_size"]
                        bold_run.bold = True

                        # sufixo normal
                        normal_run = para.add_run(sufixo.strip())
                        if ref_style:
                            normal_run.font.name = ref_style["font_name"]
                            normal_run.font.size = ref_style["font_size"]
                        normal_run.bold = False

                    para.paragraph_format.line_spacing = 1
                    para.paragraph_format.space_after = 0

    novo_docx = docx_path.replace(".docx", f"_{idioma_destino}.docx")
    doc.save(novo_docx)
    return novo_docx

def docx_para_pdf(docx_path):
    pasta = os.path.dirname(docx_path)
    convert(docx_path, pasta)
    return docx_path.replace(".docx", ".pdf")



def traduzir_pdf_layout_total(pdf_path, idioma_destino):

    docx_temp = pdf_para_docx(pdf_path)


    traduzido_docx = traduzir_docx(docx_temp, idioma_destino)


    saida_pdf = docx_para_pdf(traduzido_docx)


    try:
        if os.path.exists(docx_temp):
            os.remove(docx_temp)
    except:
        pass

    try:
        if os.path.exists(traduzido_docx):
            os.remove(traduzido_docx)
    except:
        pass

    print(f"✅ PDF final salvo em: {saida_pdf}")
    return saida_pdf



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

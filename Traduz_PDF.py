import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches
from deep_translator import GoogleTranslator
from pathlib import Path
import win32com.client as win32
import tempfile, os, io, warnings
from tkinter import Tk, filedialog, messagebox, StringVar, Label, Button, Entry
from zipfile import ZipFile
import xml.etree.ElementTree as ET


# === Configurações globais ===
warnings.filterwarnings("ignore", category=RuntimeWarning)


# === Etapa 1: Converter PDF → DOCX via Microsoft Word (layout textual) ===
def pdf_para_docx_via_word(pdf_path):
    pdf_path = Path(pdf_path)
    docx_path = pdf_path.with_suffix(".docx")

    print("➡️ Convertendo PDF → DOCX via Word COM...")
    word = win32.Dispatch("Word.Application")
    word.Visible = False
    try:
        # Abre o PDF com segurança
        doc = word.Documents.Open(str(pdf_path), ReadOnly=True, ConfirmConversions=False)

        # Corrige espaçamento inválido (evita division by zero)
        for paragraph in doc.Paragraphs:
            try:
                pf = paragraph.Format
                if pf.LineSpacingRule == 0:
                    pf.LineSpacingRule = 1
            except Exception:
                continue

        try:
            doc.SaveAs(str(docx_path), FileFormat=16)  # wdFormatDocumentDefault
        except Exception as e:
            if "division by zero" in str(e).lower():
                print("⚠️ Aviso: erro interno de espaçamento ignorado (Word).")
            else:
                raise

        doc.Close(False)
    finally:
        word.Quit()

    print(f"✅ Conversão Word concluída: {docx_path}")
    return docx_path



# === Etapa 2: Extrair imagens e gráficos do PDF via PyMuPDF ===
def extrair_imagens_pdf(pdf_path, pasta_saida):
    pdf_path = Path(pdf_path)
    print("🖼️ Extraindo imagens e gráficos do PDF...")

    doc = fitz.open(pdf_path)
    imagens_extraidas = []

    for page_index in range(len(doc)):
        try:
            page = doc.load_page(page_index)
            image_list = page.get_images(full=True)
            if not image_list:
                continue

            for img_index, img in enumerate(image_list):
                try:
                    xref = img[0]
                    base_image = doc.extract_image(xref)
                    image_data = base_image.get("image", b"")
                    image_ext = base_image.get("ext", "png")

                    if not image_data:
                        continue

                    # Segurança contra imagens inválidas (largura/altura zero)
                    pix = fitz.Pixmap(doc, xref)
                    if pix.width == 0 or pix.height == 0:
                        print(f"⚠️ Imagem inválida (0px) na página {page_index+1}, ignorada.")
                        continue

                    img_nome = f"page{page_index+1}_img{img_index+1}.{image_ext}"
                    img_caminho = pasta_saida / img_nome
                    with open(img_caminho, "wb") as img_f:
                        img_f.write(image_data)
                    imagens_extraidas.append(img_caminho)

                except Exception as e:
                    print(f"⚠️ Erro ao extrair imagem da página {page_index+1}: {e}")
                    continue

        except Exception as e:
            print(f"⚠️ Erro ao processar página {page_index+1}: {e}")
            continue

    print(f"✅ {len(imagens_extraidas)} imagens extraídas.")
    return imagens_extraidas



# === Etapa 3: Inserir imagens no DOCX traduzido ===
def inserir_imagens_docx(docx_path, imagens):
    doc = Document(docx_path)
    for img_path in imagens:
        try:
            doc.add_picture(str(img_path), width=Inches(6))
        except Exception as e:
            print(f"⚠️ Erro ao inserir imagem {img_path.name}: {e}")
            continue
    doc.save(docx_path)
    print("🧩 Imagens adicionadas ao DOCX.")



# === Etapa 4: Traduzir o DOCX mantendo estilo ===
def traduzir_docx(docx_path, idioma_destino):
    import pythoncom, time
    from win32com.client import Dispatch, constants

    pythoncom.CoInitialize()
    tradutor = GoogleTranslator(source="auto", target=idioma_destino)
    docx_path = Path(docx_path)
    traduzido_docx = docx_path.with_name(docx_path.stem + f"_{idioma_destino}.docx")

    print(f"🌍 Iniciando tradução completa via Word COM → {idioma_destino}")

    word = None
    doc = None
    try:
        word = Dispatch("Word.Application")
        word.Visible = False
        time.sleep(1)

        doc = word.Documents.Open(str(docx_path), ReadOnly=False, ConfirmConversions=False)
        total_paragrafos = doc.Paragraphs.Count
        print(f"📄 Documento contém {total_paragrafos} parágrafos diretos.")

        traduzidos = 0

        # === 1️⃣ Tradução de parágrafos diretos ===
        for i, paragrafo in enumerate(doc.Paragraphs, 1):
            texto = paragrafo.Range.Text.strip()
            if texto:
                try:
                    traduzido = tradutor.translate(texto)
                    if traduzido.strip():
                        paragrafo.Range.Text = traduzido
                        traduzidos += 1
                except Exception as e:
                    print(f"⚠️ Parágrafo {i} ignorado: {e}")
            if i % 10 == 0:
                print(f"📝 Progresso: {i}/{total_paragrafos} parágrafos traduzidos")

        # === 2️⃣ Tradução dentro de tabelas ===
        print("📊 Traduzindo conteúdo de tabelas...")
        for t_index, tabela in enumerate(doc.Tables, 1):
            for r in range(1, tabela.Rows.Count + 1):
                for c in range(1, tabela.Columns.Count + 1):
                    try:
                        celula = tabela.Cell(r, c)
                        texto = celula.Range.Text.replace('\r\x07', '').strip()
                        if texto:
                            traduzido = tradutor.translate(texto)
                            if traduzido.strip():
                                celula.Range.Text = traduzido
                                traduzidos += 1
                    except Exception as e:
                        print(f"⚠️ Erro traduzindo célula {r},{c} da tabela {t_index}: {e}")

        # === 3️⃣ Tradução de cabeçalhos e rodapés ===
        print("📄 Traduzindo cabeçalhos e rodapés...")
        for section in doc.Sections:
            for header in section.Headers:
                try:
                    texto = header.Range.Text.strip()
                    if texto:
                        traduzido = tradutor.translate(texto)
                        header.Range.Text = traduzido
                        traduzidos += 1
                except Exception as e:
                    print(f"⚠️ Cabeçalho ignorado: {e}")
            for footer in section.Footers:
                try:
                    texto = footer.Range.Text.strip()
                    if texto:
                        traduzido = tradutor.translate(texto)
                        footer.Range.Text = traduzido
                        traduzidos += 1
                except Exception as e:
                    print(f"⚠️ Rodapé ignorado: {e}")

        # === 4️⃣ Tradução de StoryRanges (caixas de texto, shapes, etc.) ===
        print("🧩 Traduzindo caixas de texto e shapes...")
        try:
            story = doc.StoryRanges(constants.wdMainTextStory)
            while story is not None:
                try:
                    if hasattr(story, "Text") and story.Text.strip():
                        traduzido = tradutor.translate(story.Text)
                        if traduzido.strip():
                            story.Text = traduzido
                            traduzidos += 1
                except Exception as e:
                    print(f"⚠️ Ignorando StoryRange inválido: {e}")
                story = story.NextStoryRange
        except Exception as e:
            print(f"⚠️ Falha ao iterar StoryRanges: {e}")

        # === 5️⃣ Salvamento final ===
        print(f"✅ {traduzidos} blocos de texto traduzidos com sucesso.")
        doc.SaveAs(str(traduzido_docx))
        doc.Close(False)
        print(f"💾 Arquivo traduzido salvo em: {traduzido_docx}")
        return traduzido_docx

    except Exception as e:
        raise RuntimeError(f"Erro durante tradução via Word COM: {e}")

    finally:
        try:
            if doc:
                doc.Close(False)
        except:
            pass
        try:
            if word:
                word.Quit()
        except:
            pass
        pythoncom.CoUninitialize()


# === Etapa 5: DOCX → PDF ===
def docx_para_pdf(docx_path):
    docx_path = Path(docx_path)
    pdf_path = docx_path.with_suffix(".pdf")
    print("➡️ Convertendo DOCX traduzido → PDF via Word COM...")

    # 1️⃣ Corrige dimensões inválidas (0 px)
    try:
        temp_dir = tempfile.mkdtemp()
        with ZipFile(docx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        drawing_files = list(Path(temp_dir).rglob("drawing*.xml"))
        for xml_file in drawing_files:
            try:
                tree = ET.parse(xml_file)
                root = tree.getroot()
                changed = False
                for elem in root.iter():
                    for attr in ("cx", "cy"):
                        val = elem.attrib.get(attr)
                        if val is not None and val.isdigit() and int(val) == 0:
                            elem.attrib[attr] = "200000"  # ~2 cm
                            changed = True
                if changed:
                    tree.write(xml_file, encoding="utf-8", xml_declaration=True)
            except Exception as e:
                print(f"⚠️ Ignorando erro ao ajustar {xml_file.name}: {e}")
                continue

        fixed_docx = docx_path.with_name(docx_path.stem + "_fixed.docx")
        with ZipFile(fixed_docx, 'w') as zip_out:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = Path(root_dir) / file
                    rel_path = file_path.relative_to(temp_dir)
                    zip_out.write(file_path, rel_path)
        docx_path = fixed_docx
        print("🧩 Ajuste automático de imagens concluído.")
    except Exception as e:
        print(f"⚠️ Falha no ajuste automático de imagens: {e}")

    # 2️⃣ Conversão final via Word COM
    word = win32.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(str(docx_path))
        doc.SaveAs(str(pdf_path), FileFormat=17)  # wdFormatPDF
        doc.Close(False)
        print(f"✅ PDF final salvo: {pdf_path}")
    except Exception as e:
        if "division by zero" in str(e).lower():
            print("⚠️ Erro interno ignorado: division by zero no Word.")
        else:
            raise RuntimeError(f"Erro ao converter DOCX → PDF via Word: {e}")
    finally:
        word.Quit()

    return pdf_path



# === Função principal ===
def traduzir_pdf_hibrido(pdf_path, idioma_destino):
    print(f"🔄 Iniciando tradução híbrida para {pdf_path}")
    temp_dir = Path(tempfile.mkdtemp())

    try:
        docx_temp = pdf_para_docx_via_word(pdf_path)
        imagens = extrair_imagens_pdf(pdf_path, temp_dir)
        inserir_imagens_docx(docx_temp, imagens)
        docx_trad = traduzir_docx(docx_temp, idioma_destino)
        saida_pdf = docx_para_pdf(docx_trad)
        return docx_trad, saida_pdf

    except Exception as e:
        if "division by zero" in str(e).lower():
            raise RuntimeError("Erro interno do Word (division by zero). Arquivo pode conter formatação inválida.")
        raise



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
        docx_trad, pdf_trad = traduzir_pdf_hibrido(caminho, idioma)
        messagebox.showinfo(
            "Tradução concluída ✅",
            f"📄 DOCX traduzido: {docx_trad}\n\n📄 PDF final: {pdf_trad}\n\n"
            f"As imagens e gráficos foram reinseridos automaticamente."
        )
    except Exception as e:
        messagebox.showerror("Erro", f"Falha durante a tradução:\n{e}")


# === GUI ===
root = Tk()
root.title("Tradutor Híbrido PDF → DOCX → PDF — SERTA ⚙️")
root.geometry("600x320")
root.resizable(False, False)

caminho_var = StringVar()
idioma_var = StringVar()

Label(root, text="Selecione o PDF:", font=("Arial", 11)).pack(pady=10)
Entry(root, textvariable=caminho_var, width=65).pack()
Button(root, text="Escolher PDF", command=selecionar_arquivo).pack(pady=5)

Label(root, text="Idioma destino (ex: en, es, fr, it):", font=("Arial", 11)).pack(pady=10)
Entry(root, textvariable=idioma_var, width=15, font=("Arial", 12)).pack()

Button(
    root, text="Converter e Traduzir", command=traduzir,
    bg="#4CAF50", fg="white", font=("Arial", 12, "bold")
).pack(pady=20)

root.mainloop()

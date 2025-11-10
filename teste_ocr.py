from PIL import Image
import pytesseract

# opcional: se o comando acima falhar, define o caminho manual
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# cria uma pequena imagem com texto
from PIL import ImageDraw, ImageFont

img = Image.new("RGB", (400, 100), color=(255, 255, 255))
draw = ImageDraw.Draw(img)
draw.text((10, 30), "Teste OCR com Tesseract", fill=(0, 0, 0))
img.save("teste.png")

# aplica OCR na imagem
texto = pytesseract.image_to_string(Image.open("teste.png"), lang="por")
print("Texto detectado:", texto.strip())

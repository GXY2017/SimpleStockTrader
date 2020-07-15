import pytesseract
from PIL import Image

def captcha_recognize(img_path, threshold = 200):
    """
    200/256 转为白色，这个比较较高。
    """
    im = Image.open(img_path).convert("L") #translating a color image to greyscale
    # 1. threshold the image
    table = []
    for i in range(256):
        if i < threshold:
            table.append(0)
        else:
            table.append(1)

    out = im.point(table, '1')
    # out.show()
    # 2. recognize with tesseract
    pytesseract.pytesseract.tesseract_cmd = r"C://Program Files//Tesseract-OCR//tesseract.exe"
    num = pytesseract.image_to_string(out)
    return num
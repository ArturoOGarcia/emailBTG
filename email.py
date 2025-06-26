import os
import time
import mss
import win32com.client
import pygetwindow as gw
from datetime import datetime
from PIL import Image

# === CONFIGURAÇÕES ===
destinatario = "luis.alvaro@btgingredients.com; luiz.mizutani@btgingredients.com; steve.gulley@btgingredients.com; vmvargas@btgingredients.com"
copia = "ari.nunes@btgingredients.com; jason.medcalf@btgingredients.com"
agora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

# === 1. ABRIR EXCEL E TIRAR PRIMEIRO PRINT ===
caminho_excel_1 = r"C:\Users\artur\OneDrive - BTG Ingredients LLC\Daily Flash\20250623_BI_CUST_PL.XLSX"

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb1 = excel.Workbooks.Open(caminho_excel_1)
ws1 = wb1.Worksheets(1)
ws1.Activate()
excel.ActiveWindow.Zoom = 70

time.sleep(2)
for w in gw.getWindowsWithTitle("20250623_BI_CUST_PL"):
    if w.isMinimized:
        w.restore()
    w.activate()
    w.maximize()
    break

time.sleep(1)

nome_arquivo1 = f"print1_{agora}.png"
caminho_arquivo1 = os.path.join(os.getcwd(), nome_arquivo1)

area1 = {
    "left": 520,
    "top": 330,
    "width": 660,
    "height": 550
}

with mss.mss() as sct:
    img1 = sct.grab(area1)
    mss.tools.to_png(img1.rgb, img1.size, output=caminho_arquivo1)

print(f"Primeiro print salvo: {caminho_arquivo1}")

# === 2. ABRIR SEGUNDO EXCEL E TIRAR SEGUNDO PRINT ===
caminho_excel_2 = r"C:\Users\artur\OneDrive - BTG Ingredients LLC\Daily Flash\Faturamento - 23-06.xlsx"

wb2 = excel.Workbooks.Open(caminho_excel_2)
ws2 = wb2.Worksheets(1)
ws2.Activate()
excel.ActiveWindow.Zoom = 100

time.sleep(2)
for w in gw.getWindowsWithTitle("Faturamento - 23-06"):
    if w.isMinimized:
        w.restore()
    w.activate()
    w.maximize()
    break

time.sleep(1)

nome_arquivo2 = f"print2_{agora}.png"
caminho_arquivo2 = os.path.join(os.getcwd(), nome_arquivo2)

area2 = {
    "left": 100,
    "top": 305,
    "width": 870,
    "height": 640
}

with mss.mss() as sct:
    img2 = sct.grab(area2)
    mss.tools.to_png(img2.rgb, img2.size, output=caminho_arquivo2)

print(f"Segundo print salvo: {caminho_arquivo2}")

# === 3. ENVIAR POR OUTLOOK COM IMAGENS INLINE E PLANILHAS ANEXADAS ===
outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)

# mail.To = destinatario
mail.to = "arturo.garcia@asteritrading.com"
mail.Subject = "Daily Flash - BTG Ingredients / Asteri Trading " + datetime.now().strftime("%d/%m/%Y")
# mail.CC = copia
# mail.BCC = "arturo.garcia@asteritrading.com"

mail.HTMLBody = f"""
<p>Hi!</p>
<p>Here is the billing for BTG Ingredients and Asteri Trading up to {datetime.now().strftime('%d/%m/%Y')}.</p>
<p><img src="cid:flash1" width="800"></p>
<p><img src="cid:flash2" width="800"></p>
<p>Best regards,<br>Arturo Garcia</p>
"""

# Anexar prints com content-id para inline
attachment1 = mail.Attachments.Add(caminho_arquivo1)
attachment1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "flash1")

attachment2 = mail.Attachments.Add(caminho_arquivo2)
attachment2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "flash2")

wb1.Close(SaveChanges=False)
wb2.Close(SaveChanges=False)
excel.Quit()

# Anexar os arquivos Excel originais
mail.Attachments.Add(caminho_excel_1)
mail.Attachments.Add(caminho_excel_2)

mail.Display()
print("E-mail criado com sucesso!")
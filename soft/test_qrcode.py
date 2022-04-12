# pip install pillow
# pip install qrcode

import qrcode
import xlsxwriter

list_caseids = ('RS22-SEV-BEZ','RS22-GIN-VOD','RS22-KON-ZIF','RS22-FUQ-ETE','RS22-LER-NEZ','RS22-HIR-PIR','RS22-FEZ-HES','RS22-UTU-ISO',
'RS22-BOJ-BUF','RS22-BAF-VEC','RS22-JIZ-SUK','RS22-GAG-PIG','RS22-FOF-DAM','RS22-FEH-XUD','RS22-SOX-LAD',)

# Link for website
input_data = 'ord' + "513551"
#Creating an instance of qrcode
qr = qrcode.QRCode(
        version=1,
        box_size=5,
        border=5)
qr.add_data(input_data)
qr.make(fit=True)
img = qr.make_image(fill='black', back_color='white')
img.save('qrcode001.png')




# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('images.xlsx')
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
worksheet.set_column('A:A', 30)

# Insert an image.
worksheet.write('A2', 'Insert an image in a cell:')
worksheet.insert_image('B2', 'qrcode001.png')

workbook.close()
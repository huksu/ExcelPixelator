from PIL import Image
import xlsxwriter

workbook = xlsxwriter.Workbook('image.xlsx')
worksheet = workbook.add_worksheet()

pixelwidth = .08
pixelheight = .75
cellwidth = 15
cellheight = 15

imageFile = "image.png"
img = Image.open(imageFile)
rgb_img = img.convert('RGB')

width, height = rgb_img.size
cols = int(width / cellwidth)
rows = int(height / cellheight)

worksheet.set_column(0, cols-1, cellwidth*pixelwidth)
for j in range(0,rows):
    worksheet.set_row(j, cellheight*pixelheight)
    for i in range(0,cols):
        sumR = 0
        sumG = 0
        sumB = 0
        for m in range(0,cellwidth):
            for n in range(0,cellheight):
                pixel = rgb_img.getpixel((i*cellwidth+m, j*cellheight+n))
                sumR = sumR + pixel[0]
                sumG = sumG + pixel[1]
                sumB = sumB + pixel[2]
        avgR = sumR / (cellwidth*cellheight)
        avgG = sumG / (cellwidth*cellheight)
        avgB = sumB / (cellwidth*cellheight)
        hexcolor = '#%02x%02x%02x' % (avgR,avgG,avgB)
        format = workbook.add_format()
        format.set_bg_color(hexcolor)
        worksheet.write(j, i, '', format)
        
workbook.close()



#import pillow
import numpy
import imageio
#import myqr
import qrcode
# 存在问题：颜色格式，图片大小，小块像素，二维码信息
qr = qrcode.QRCode(
    version=2,
    error_correction=qrcode.constants.ERROR_CORRECT_L,
    box_size=10,
    border=4,
)
#img_data = 'https://github.com'
img_data = 'https://www.bilibili.com'
'''img_data = "Hello fuking you \n\
           jjjjjjjjjjjjjjj \n\
           kkkkkkkkkkkkkkk \n\
           qqqqqqqqqqqqqqqqqqqqqqq \n\
           ooooooooooooooooooooooooooooo \n\
           ppppppppppppppppppppppppppppppp \n\
           qqqqqqqqqqqqqqqqqqqqqqqqqqqqqqqqqqqqqqq \n\
           lllllllllllllllllllllllllllllllllllllllllllllll \n\
           mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm"   '''
qr.add_data(img_data)
qr.make(fit=True)

img = qr.make_image(fill_color="black", back_color="white")
img2 = qr.make_image(fill_color="white", back_color="black")
#print(img)
fir = r"C:\Users\xyue\Desktop\1.png"
img.save(fir)


#import cv2
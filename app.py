import os, io, datetime, base64, smtplib
from flask import Flask, request, jsonify, send_from_directory
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from PIL import Image, ImageOps
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

app = Flask(__name__, static_folder='.', static_url_path='')

# ======= 【请在这里修改你的邮箱配置】 =======
SMTP_SERVER = 'smtp.exmail.qq.com'  # 如果不是企业微信邮箱，请百度对应邮箱的SMTP地址
SMTP_PORT = 465
SENDER_EMAIL = tylerqin@gtiggs.com 
SENDER_PASSWORD = vfm3x9Rd74PMCF56
# ==========================================

def process_img(b64):
    if ',' in b64: b64 = b64.split(',')[1]
    img = Image.open(io.BytesIO(base64.b64decode(b64)))
    return ImageOps.exif_transpose(img)

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/upload', methods=['POST'])
def upload():
    try:
        data = request.json
        text, imgs = data.get('text'), data.get('images')
        recipients = data.get('recipients').split(',')
        
        # 制作PPT
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 加文字
        tb = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(11.3), Inches(1.5))
        p = tb.text_frame.paragraphs[0]
        p.text = text
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(32)
        
        # 加图片 (左中右排列)
        for i, img_b64 in enumerate(imgs):
            img = process_img(img_b64)
            output = io.BytesIO()
            img.save(output, format='JPEG')
            output.seek(0)
            slide.shapes.add_picture(output, Inches(1.66 + i*4), Inches(2.5), width=Inches(3.5))

        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        ppt_io.seek(0)

        # 发邮件
        msg = MIMEMultipart()
        msg['Subject'] = f"试身照片: {text}"
        msg['From'] = SENDER_EMAIL
        msg['To'] = ",".join(recipients)
        msg.attach(MIMEText(f"Hi All,\n附件是[{text}]的试身照片。\n请在回答前，先问我问题...", 'plain'))
        
        part = MIMEBase('application', "octet-stream")
        part.set_payload(ppt_io.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{text}.pptx"')
        msg.attach(part)
        
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as s:
            s.login(SENDER_EMAIL, SENDER_PASSWORD)
            s.sendmail(SENDER_EMAIL, recipients, msg.as_string())
            
        return jsonify({"status": "success"})
    except Exception as e:
        return jsonify({"status": "error", "msg": str(e)}), 500

if __name__ == '__main__':
    app.run()
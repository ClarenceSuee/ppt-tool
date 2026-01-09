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

# ======= 【配置区域：请修改这里】 =======
# 1. 设置一个简单的密码，防止外人乱用
TEAM_PASSWORD = "666" 

# 2. 邮箱设置 (以腾讯企业邮为例)
SMTP_SERVER = 'smtp.exmail.qq.com'
SMTP_PORT = 465
SENDER_EMAIL = tylerqin@gtiggs.com 
SENDER_PASSWORD = vfm3x9Rd74PMCF56
# ====================================

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
        
        # 1. 验证密码
        if data.get('pass') != TEAM_PASSWORD:
            return jsonify({"status": "error", "msg": "密码错误"}), 403

        text = data.get('text')
        sender = data.get('sender', '匿名')
        imgs = data.get('images')
        recipients = data.get('recipients').split(',')
        
        # 2. 制作PPT
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 文本框
        tb = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(11.3), Inches(1.5))
        p = tb.text_frame.paragraphs[0]
        p.text = text
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(32)
        
        # 图片 (居中排列)
        for i, img_b64 in enumerate(imgs):
            img = process_img(img_b64)
            output = io.BytesIO()
            img.save(output, format='JPEG')
            output.seek(0)
            # 计算坐标：起始位置 + 偏移量
            slide.shapes.add_picture(output, Inches(1.66 + i*4), Inches(2.5), width=Inches(3.5))

        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        ppt_io.seek(0)

        # 3. 发邮件
        filename = f"{text}_{datetime.datetime.now().strftime('%Y%m%d')}.pptx"
        msg = MIMEMultipart()
        msg['Subject'] = f"试身照片: {text} (发送人: {sender})"
        msg['From'] = SENDER_EMAIL
        msg['To'] = ",".join(recipients)
        
        body = f"""Hi All,
附件是[{text}]的试身照片, 请查收, 谢谢。"""
        
        msg.attach(MIMEText(body, 'plain'))
        
        part = MIMEBase('application', "octet-stream")
        part.set_payload(ppt_io.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
        msg.attach(part)
        
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as s:
            s.login(SENDER_EMAIL, SENDER_PASSWORD)
            s.sendmail(SENDER_EMAIL, recipients, msg.as_string())
            
        return jsonify({"status": "success"})
    except Exception as e:
        print(e)
        return jsonify({"status": "error", "msg": str(e)}), 500

if __name__ == '__main__':
    app.run()

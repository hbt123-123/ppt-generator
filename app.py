import os
import uuid
import json
import io
import base64
import logging
from datetime import datetime
from flask import Flask, render_template, request, send_file, jsonify
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image
import tempfile

# 先创建必要的目录
os.makedirs('logs', exist_ok=True)
os.makedirs('uploads', exist_ok=True)

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/app.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)
logger.info("应用启动中...")

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB限制
app.config['SECRET_KEY'] = str(uuid.uuid4())  # 添加安全密钥

# 验证目录可写性
try:
    logger.info(f"上传目录检查成功: {app.config['UPLOAD_FOLDER']}")
    logger.info("日志目录检查成功")
    
    # 验证上传目录可写性
    test_file = os.path.join(app.config['UPLOAD_FOLDER'], '.test_writability')
    with open(test_file, 'w') as f:
        f.write('test')
    os.remove(test_file)
    logger.info("目录可写性测试通过")
    
except Exception as e:
    logger.error(f"目录权限检查失败: {str(e)}")
    raise

# 主题映射 - 仅保留6个有效主题
THEMES = {
    "spring": "themes/spring.pptx",
    "summer": "themes/summer.pptx",
    "autumn": "themes/autumn.pptx",
    "winter": "themes/winter.pptx",
    "cat": "themes/cat.pptx",
    "red": "themes/red.pptx"
}

# 版式映射
SLIDE_LAYOUTS = {
    "title_subtitle": 0,
    "title_only": 5,
    "section_header": 2,
    "content": 1,
    "two_content": 3,
    "comparison": 4,
    "content_caption": 7,
    "picture_caption": 8,
    "video": 6
}


def compress_image(image_data, max_size=1920, quality=85):
    """压缩图片并转换为JPEG格式"""
    try:
        # 检查image_data格式并提取Base64部分
        if ',' not in image_data:
            logger.error("图片数据格式不正确")
            return None
            
        # 将Base64数据转换为图像
        base64_part = image_data.split(',')[1]
        img_bytes = base64.b64decode(base64_part)
        
        # 创建临时文件
        temp_path = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.jpg') as temp_file:
                temp_path = temp_file.name

            # 打开图像并压缩
            img = Image.open(io.BytesIO(img_bytes))
            logger.info(f"原始图片尺寸: {img.size}, 模式: {img.mode}")
            
            # 转换模式为RGB
            if img.mode in ('RGBA', 'P'):
                img = img.convert('RGB')

            # 调整大小
            width, height = img.size
            if width > max_size or height > max_size:
                ratio = min(max_size / width, max_size / height)
                new_size = (int(width * ratio), int(height * ratio))
                img = img.resize(new_size, Image.Resampling.LANCZOS)
                logger.info(f"调整后图片尺寸: {new_size}")

            # 保存压缩后的图像
            img.save(temp_path, 'JPEG', quality=quality)

            # 读取压缩后的数据
            with open(temp_path, 'rb') as f:
                compressed_data = f.read()
                
            logger.info(f"图片压缩完成，大小: {len(compressed_data)} bytes")
            return compressed_data
            
        finally:
            # 确保临时文件被删除
            if temp_path and os.path.exists(temp_path):
                os.unlink(temp_path)

    except Exception as e:
        logger.error(f"图片压缩失败: {str(e)}")
        # 尝试返回原始数据
        try:
            return base64.b64decode(image_data.split(',')[1])
        except:
            return None


@app.route('/')
def index():
    """首页 - 展示所有主题模板"""
    theme_previews = {
        theme: f"/static/theme-preview/{theme}.jpg"
        for theme in THEMES.keys()
    }
    now_date = datetime.now().strftime("%Y年%m月%d日")
    return render_template('index.html', themes=THEMES.keys(),
                           theme_previews=theme_previews, now_date=now_date)


@app.route('/generate', methods=['POST'])
def generate_ppt():
    """生成PPT文件"""
    try:
        data = request.get_json()
        logger.info(f"收到生成请求: {data.keys()}")

        title = data.get('title', '毕业论文答辩')
        subtitle = data.get('subtitle', '')
        presenter = data.get('presenter', '张三')
        date = data.get('date', datetime.now().strftime("%Y年%m月%d日"))
        team = data.get('team', '资源与环境学院')
        theme = data.get('theme', 'spring')
        slides_data = data.get('slides', [])
        thank_you_text = data.get('thank_you_text', '感谢观看')

        font_size = data.get('font_size', 24)
        line_spacing = data.get('line_spacing', 1.15)
        auto_numbering = data.get('auto_numbering', True)

        template_path = THEMES.get(theme, THEMES['spring'])
        logger.info(f"使用主题模板: {template_path}")
        prs = Presentation(template_path)

        # 添加封面页
        cover_layout = SLIDE_LAYOUTS["title_subtitle"] if data.get('use_subtitle', True) else SLIDE_LAYOUTS[
            "title_only"]
        slide = prs.slides.add_slide(prs.slide_layouts[cover_layout])

        # 设置封面标题
        if slide.shapes.title:
            slide.shapes.title.text = title
            slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(44)

        # 设置副标题
        if data.get('use_subtitle', True) and len(slide.placeholders) > 1:
            subtitle_shape = slide.placeholders[1]
            subtitle_text = ""
            if subtitle: subtitle_text += subtitle + "\n"
            if presenter: subtitle_text += f"主讲人: {presenter}\n"
            if date: subtitle_text += f"日期: {date}\n"
            if team: subtitle_text += f"团队: {team}"

            subtitle_shape.text = subtitle_text
            for paragraph in subtitle_shape.text_frame.paragraphs:
                paragraph.font.size = Pt(24)

        # 添加目录页
        if data.get('include_toc', True) and any(slide.get('include_in_toc', True) for slide in slides_data):
            toc_slide = prs.slides.add_slide(prs.slide_layouts[SLIDE_LAYOUTS["content"]])
            if toc_slide.shapes.title:
                toc_slide.shapes.title.text = "目录"
                toc_slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(36)

            if len(toc_slide.placeholders) > 1:
                toc_content = toc_slide.placeholders[1].text_frame
                toc_content.clear()

                for i, slide_data in enumerate(slides_data):
                    if slide_data.get('include_in_toc', True) and slide_data.get('type') != 'section_header':
                        p = toc_content.add_paragraph()
                        p.text = slide_data.get('title', f"章节 {i + 1}")
                        p.level = 0
                        p.font.size = Pt(font_size)
                        run = p.add_run()
                        run.text = f" ...... {i + 1}"
                        run.font.size = Pt(font_size - 2)
                        run.font.color.rgb = RGBColor(150, 150, 150)

        # 处理内容幻灯片
        for i, slide_data in enumerate(slides_data):
            slide_type = slide_data.get('type', 'content')
            layout_index = SLIDE_LAYOUTS.get(slide_type, SLIDE_LAYOUTS["content"])

            # 添加节标题
            if slide_type == 'section_header':
                slide = prs.slides.add_slide(prs.slide_layouts[layout_index])
                if slide.shapes.title:
                    slide.shapes.title.text = slide_data.get('title', f"第 {i + 1} 部分")
                    slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(40)

                # 添加节副标题
                if len(slide.placeholders) > 1:
                    subtitle_shape = slide.placeholders[1]
                    subtitle_shape.text = slide_data.get('subtitle', '')
                    for paragraph in subtitle_shape.text_frame.paragraphs:
                        paragraph.font.size = Pt(28)
                continue

            # 添加其他类型幻灯片
            slide = prs.slides.add_slide(prs.slide_layouts[layout_index])

            # 设置标题
            if slide.shapes.title:
                slide.shapes.title.text = slide_data.get('title', f"幻灯片 {i + 1}")
                slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(36)

            # 根据幻灯片类型设置内容
            if slide_type == 'content' and len(slide.placeholders) > 1:
                # 标题和内容版式
                content_shape = slide.placeholders[1]
                tf = content_shape.text_frame
                tf.clear()
                for idx, content in enumerate(slide_data.get('content', [])):
                    if content.strip():
                        p = tf.add_paragraph()
                        if auto_numbering:
                            p.text = f"{idx + 1}. {content}"
                        else:
                            p.text = f"• {content}"
                        p.level = 0
                        p.font.size = Pt(font_size)

            elif slide_type == 'two_content' and len(slide.placeholders) > 2:
                # 两栏内容版式
                # 左侧内容
                left_content_shape = slide.placeholders[1]
                tf_left = left_content_shape.text_frame
                tf_left.clear()
                for idx, content in enumerate(slide_data.get('left_content', [])):
                    if content.strip():
                        p = tf_left.add_paragraph()
                        if auto_numbering:
                            p.text = f"{idx + 1}. {content}"
                        else:
                            p.text = f"• {content}"
                        p.level = 0
                        p.font.size = Pt(font_size)

                # 右侧内容
                right_content_shape = slide.placeholders[2]
                tf_right = right_content_shape.text_frame
                tf_right.clear()
                for idx, content in enumerate(slide_data.get('right_content', [])):
                    if content.strip():
                        p = tf_right.add_paragraph()
                        if auto_numbering:
                            p.text = f"{idx + 1}. {content}"
                        else:
                            p.text = f"• {content}"
                        p.level = 0
                        p.font.size = Pt(font_size)

            elif slide_type == 'comparison' and len(slide.placeholders) > 4:
                # 比较版式
                # 左侧标题
                if slide.placeholders[1]:
                    slide.placeholders[1].text = slide_data.get('left_subtitle', '左侧标题')

                # 右侧标题
                if slide.placeholders[2]:
                    slide.placeholders[2].text = slide_data.get('right_subtitle', '右侧标题')

                # 左侧内容
                left_content_shape = slide.placeholders[3]
                tf_left = left_content_shape.text_frame
                tf_left.clear()
                for idx, content in enumerate(slide_data.get('left_content', [])):
                    if content.strip():
                        p = tf_left.add_paragraph()
                        p.text = f"• {content}"
                        p.level = 0
                        p.font.size = Pt(font_size)

                # 右侧内容
                right_content_shape = slide.placeholders[4]
                tf_right = right_content_shape.text_frame
                tf_right.clear()
                for idx, content in enumerate(slide_data.get('right_content', [])):
                    if content.strip():
                        p = tf_right.add_paragraph()
                        p.text = f"• {content}"
                        p.level = 0
                        p.font.size = Pt(font_size)

            elif slide_type == 'content_caption' and len(slide.placeholders) > 1:
                # 内容与标题版式
                content_shape = slide.placeholders[1]
                tf = content_shape.text_frame
                tf.clear()
                for idx, content in enumerate(slide_data.get('content', [])):
                    if content.strip():
                        p = tf.add_paragraph()
                        if auto_numbering:
                            p.text = f"{idx + 1}. {content}"
                        else:
                            p.text = f"• {content}"
                        p.level = 0
                        p.font.size = Pt(font_size)

            # 图片处理
            if slide_data.get('image'):
                try:
                    compressed_img = compress_image(slide_data['image'])
                    
                    # 检查压缩后的图片数据是否有效
                    if compressed_img:
                        img_stream = io.BytesIO(compressed_img)

                        # 图片与标题版式 - 居中大图
                        if slide_type == 'picture_caption':
                            left = (prs.slide_width - Inches(8)) / 2
                            top = Inches(2)
                            slide.shapes.add_picture(img_stream, left, top, width=Inches(8))
                            logger.info(f"添加居中大图成功")

                        # 内容与标题版式 - 右侧图片
                        elif slide_type == 'content_caption':
                            left = Inches(6)
                            top = Inches(1.5)
                            height = Inches(4.5)
                            slide.shapes.add_picture(img_stream, left, top, height=height)
                            logger.info(f"添加右侧内容图片成功")

                        # 标准内容版式 - 右侧图片
                        elif slide_type in ['content', 'two_content']:
                            left = Inches(6)
                            top = Inches(1.5)
                            height = Inches(4.5)
                            slide.shapes.add_picture(img_stream, left, top, height=height)
                            logger.info(f"添加内容区右侧图片成功")
                    else:
                        logger.warning("图片数据无效，跳过添加图片")

                except Exception as e:
                    logger.error(f"添加图片失败: {str(e)}")

            # 视频处理
            if slide_data.get('video_url') and slide_type == 'video':
                # 居中视频区域
                left = (prs.slide_width - Inches(6)) / 2
                top = (prs.slide_height - Inches(4)) / 2
                width = Inches(6)
                height = Inches(4)

                # 添加视频容器
                shape = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
                )
                shape.text = slide_data.get('video_text', '视频内容\n点击播放')
                shape.fill.solid()
                shape.fill.fore_color.rgb = RGBColor(240, 240, 240)
                shape.line.color.rgb = RGBColor(200, 200, 200)

                # 添加播放图标
                play_left = left + (width - Inches(1.5)) / 2
                play_top = top + (height - Inches(1.5)) / 2
                play_shape = slide.shapes.add_shape(
                    MSO_SHAPE.TRIANGLE, play_left, play_top, Inches(1.5), Inches(1.5)
                )
                play_shape.rotation = 0
                play_shape.fill.solid()
                play_shape.fill.fore_color.rgb = RGBColor(50, 120, 200)
                play_shape.line.color.rgb = RGBColor(50, 120, 200)

                # 添加超链接
                play_shape.click_action.hyperlink.address = slide_data['video_url']

        # 添加感谢页
        end_slide = prs.slides.add_slide(prs.slide_layouts[0])  # 使用空白版式
        left = top = Inches(1)
        width = prs.slide_width - Inches(2)
        height = prs.slide_height - Inches(2)
        textbox = end_slide.shapes.add_textbox(left, top, width, height)
        tf = textbox.text_frame
        tf.clear()

        # 使用用户自定义的感谢文本
        thank_you_lines = thank_you_text.split('\n')
        for i, line in enumerate(thank_you_lines):
            p = tf.add_paragraph()
            p.text = line
            if i == 0:  # 第一行作为主标题
                p.font.size = Pt(60)
                p.font.bold = True
            else:
                p.font.size = Pt(36)
            p.alignment = PP_ALIGN.CENTER
            if i > 0:
                p.space_before = Pt(30)

        # 保存PPT到内存
        ppt_stream = io.BytesIO()
        prs.save(ppt_stream)
        ppt_stream.seek(0)

        # 返回文件
        filename = title.replace(" ", "_") + ".pptx"
        logger.info(f"PPT生成成功: {filename}")
        return send_file(
            ppt_stream,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )

    except Exception as e:
        logger.exception("生成PPT时发生严重错误")
        return jsonify({
            "error": "生成PPT时发生错误",
            "message": str(e)
        }), 500


if __name__ == '__main__':
    app.run(debug=True, port=5000)
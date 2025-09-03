from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

def create_name_badge(names):
    # 创建PDF文件
    pdf_filename = "name_badges.pdf"
    c = canvas.Canvas(pdf_filename, pagesize=A4)

    # 注册字体 - 使用系统自带的字体，确保中文显示正常
    if os.name == "nt":
        # 尝试注册微软雅黑字体（Windows系统）
        pdfmetrics.registerFont(TTFont('SimSong', 'simsong.ttf'))
        font_name = 'SimSong'
    elif os.name == "posix":
        # 尝试注册苹果系统字体（macOS系统）的宋体
        pdfmetrics.registerFont(TTFont('Songti-Light', '/System/Library/Fonts/Supplemental/Songti.ttc', subfontIndex=1))
        font_name = 'Songti-Light'
    else:
        font_name = 'Helvetica'  # 默认字体

    # 获取页面尺寸
    width, height = A4

    for name in names:
        # 处理两个字的名字，中间加空格
        if len(name) == 2:
            name = name[0] + " " + name[1]

        # 设置字体大小
        font_size = 180 # min(300, 1000 // max(1, len(name)))
        spacing_ratio = 0.28  # 控制上下部分间距的比例

        # 四个字
        if len(name) >= 4:
            font_size = font_size * 3 / len(name) + 5

        # 下半部分 - 正常方向的名字
        c.setFont(font_name, font_size)
        c.drawCentredString(width/2, height * spacing_ratio, name)

        # 上半部分 - 翻转的名字
        c.saveState()
        c.translate(width/2, height * (1 - spacing_ratio))  # 移到上半部分中心
        c.rotate(180)                    # 旋转180度
        c.drawCentredString(0, 0, name)   # 在新的原点绘制
        c.restoreState()

        # 添加新页面
        c.showPage()

    # 保存PDF
    c.save()
    print(f"名牌已创建: {os.path.abspath(pdf_filename)}")
    return pdf_filename

if __name__ == "__main__":
    # 从name.txt读取名字
    if os.path.exists("name.txt"):
        with open("name.txt", "r", encoding="utf-8") as f:
            names = [line.strip() for line in f if line.strip()]
    else:
        names = input("请输入名字（多个名字用空格分隔）: ").split()

    create_name_badge(names)

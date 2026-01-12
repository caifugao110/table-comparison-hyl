from PIL import Image, ImageDraw, ImageFont
import os

def generate_compare_excel_ico():
    # 定义图标尺寸（仅128x128）
    size = (128, 128)
    # 配色（沿用代码主题色）
    main_color = (102, 126, 234)   # 主蓝
    white = (255, 255, 255)        # 白色
    accent = (255, 107, 53)        # 提示色
    
    # 加载字体（兼容系统，无则用默认）
    try:
        font = ImageFont.truetype("arial.ttf", 28)  # 调整字体大小适配128尺寸
    except:
        font = ImageFont.load_default()

    # 创建图像
    img = Image.new("RGB", size, main_color)
    draw = ImageDraw.Draw(img)
    
    # 绘制Excel表格格子
    cell_w = size[0] // 4
    # 左格子（基准文件）
    draw.rectangle([cell_w, cell_w, cell_w*2, cell_w*3], fill=white, width=2)
    draw.text((cell_w+8, cell_w+12), "A", font=font, fill=main_color)
    # 右格子（对比文件）
    draw.rectangle([cell_w*2.5, cell_w, cell_w*3.5, cell_w*3], fill=white, width=2)
    draw.text((cell_w*2.8, cell_w+12), "B", font=font, fill=main_color)
    
    # 绘制对比箭头
    arrow_x = cell_w*2 + 4
    arrow_y = cell_w*2
    draw.line([arrow_x, arrow_y, cell_w*2.5, arrow_y], fill=white, width=4)
    draw.polygon([cell_w*2.5, arrow_y-4, cell_w*2.5, arrow_y+4, cell_w*2.8, arrow_y], fill=white)
    
    # 保存路径
    ico_path = os.path.join(r"d:\mygit\table-comparison-hyl\ico", "compare_excel.ico")
    img.save(
        ico_path, format="ICO",
        sizes=[(size[0], size[1])]
    )
    
    print(f"✅ 图标生成完成：{os.path.abspath(ico_path)}")

if __name__ == "__main__":
    # 需安装Pillow：pip install pillow
    generate_compare_excel_ico()
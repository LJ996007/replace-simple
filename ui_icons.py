"""按显示尺寸生成抗锯齿图标，不依赖字符字体或外部图片文件。"""
from PIL import Image, ImageDraw


def checkbox_icon(size, color, state='unchecked'):
    """细方框与内收的圆端勾；高分辨率绘制后缩小，避免字体替换。"""
    scale = size * 4 / 20
    image = Image.new('RGBA', (size * 4, size * 4))
    draw = ImageDraw.Draw(image)
    box = tuple(round(v * scale) for v in (2.5, 2.5, 17.5, 17.5))
    draw.rectangle(box, outline=color, width=max(1, round(scale)))
    points = {'checked': [(6, 10), (9, 13), (14, 7)],
              'mixed': [(6, 10), (14, 10)]}.get(state)
    if points:
        points = [(round(x * scale), round(y * scale)) for x, y in points]
        width = max(1, round(1.7 * scale))
        draw.line(points, fill=color, width=width, joint='curve')
        radius = width / 2
        for x, y in points:
            draw.ellipse((x-radius, y-radius, x+radius, y+radius), fill=color)
    return image.resize((size, size), Image.Resampling.LANCZOS)


def zoom_icon(size, color, zoom_in=False):
    scale = size * 4 / 24
    image = Image.new('RGBA', (size * 4, size * 4))
    draw = ImageDraw.Draw(image)
    width = max(1, round(1.8 * scale))
    def line(points):
        coords = [(round(x * scale), round(y * scale)) for x, y in points]
        draw.line(coords, fill=color, width=width)
        radius = width / 2
        for x, y in (coords[0], coords[-1]):
            draw.ellipse((x-radius,y-radius,x+radius,y+radius),fill=color)
    line([(15,15),(20,20)])
    draw.ellipse(tuple(round(v * scale) for v in (3,3,17,17)),outline=color,width=width)
    line([(7,10),(13,10)])
    if zoom_in:
        line([(10,7),(10,13)])
    return image.resize((size,size),Image.Resampling.LANCZOS)

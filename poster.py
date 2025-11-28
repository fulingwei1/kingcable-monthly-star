from io import BytesIO
from typing import Dict, Any

from PIL import Image, ImageDraw, ImageFont


def _load_font(size: int) -> ImageFont.ImageFont:
    """
    尝试加载几种常见中文字体，失败则退回 PIL 内置字体（不会再抛异常）。
    """
    font_paths = [
        "fonts/SourceHanSansSC-Regular.otf",
        "fonts/SourceHanSansSC-Regular.ttf",
        "fonts/SimHei.ttf",
        "/System/Library/Fonts/STHeiti Medium.ttc",
        "/System/Library/Fonts/STHeiti.ttc",
    ]
    for path in font_paths:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            continue
    # 兜底：内置小字体，保证不崩
    return ImageFont.load_default()


def make_initial_avatar(name: str, size: int = 260) -> Image.Image:
    """
    当没上传头像时，用姓名生成一个夸张一点的圆形首字母头像。
    """
    text = (name or "").strip()
    ch = text[0] if text else "?"

    # 用名字的 hash 决定背景色，保证同一人颜色稳定
    h = abs(hash(name)) if name else 0
    bg_color = (80 + h % 150, 80 + (h // 10) % 150, 80 + (h // 100) % 150)

    img = Image.new("RGBA", (size, size), bg_color)
    draw = ImageDraw.Draw(img)

    font = _load_font(int(size * 0.6))
    w, h_text = draw.textsize(ch, font=font)
    x = (size - w) / 2
    y = (size - h_text) / 2 - 10

    draw.text((x, y), ch, font=font, fill=(255, 255, 255, 255))

    # 裁成圆形
    mask = Image.new("L", (size, size), 0)
    mdraw = ImageDraw.Draw(mask)
    mdraw.ellipse((0, 0, size, size), fill=255)
    circle = Image.new("RGBA", (size, size))
    circle.paste(img, (0, 0), mask)

    return circle


def _draw_wrapped_text(
    draw: ImageDraw.ImageDraw,
    text: str,
    font: ImageFont.ImageFont,
    xy,
    max_width: int,
    fill=(0, 0, 0, 255),
    line_spacing: int = 4,
):
    """
    简单按像素宽度换行绘制多行中文文本。
    """
    x, y = xy
    lines = []

    for para in (text or "").splitlines():
        para = para.rstrip()
        if not para:
            lines.append("")
            continue
        buf = ""
        for ch in para:
            w, _ = draw.textsize(buf + ch, font=font)
            if w <= max_width:
                buf += ch
            else:
                lines.append(buf)
                buf = ch
        if buf:
            lines.append(buf)

    line_height = getattr(font, "size", 16) + line_spacing

    for line in lines:
        draw.text((x, y), line, font=font, fill=fill)
        y += line_height


def generate_poster(
    template_img: Image.Image,
    avatar_img: Image.Image,
    star: Dict[str, Any],
    month_label: str,
) -> Image.Image:
    """
    把模板图、头像、姓名/部门/奖项/评语等信息合成一张海报。
    模板坐标默认写死，你要微调就自己改常量。
    """
    base = template_img.convert("RGBA")

    # 一些布局常量（按你模板大概 1080x1920 竖版来设计的）
    AVATAR_SIZE = 260
    AVATAR_POS = (80, 260)  # 左上角
    NAME_POS = (380, 260)
    DEPT_POS = (380, 320)
    AWARD_POS = (380, 380)
    MONTH_POS = (380, 440)
    COMMENT_POS = (120, 540)
    COMMENT_WIDTH = base.size[0] - 240

    # 处理头像为圆形
    avatar = avatar_img.convert("RGBA").resize((AVATAR_SIZE, AVATAR_SIZE))
    mask = Image.new("L", (AVATAR_SIZE, AVATAR_SIZE), 0)
    mdraw = ImageDraw.Draw(mask)
    mdraw.ellipse((0, 0, AVATAR_SIZE, AVATAR_SIZE), fill=255)
    circle_avatar = Image.new("RGBA", (AVATAR_SIZE, AVATAR_SIZE))
    circle_avatar.paste(avatar, (0, 0), mask)

    base.paste(circle_avatar, AVATAR_POS, circle_avatar)

    draw = ImageDraw.Draw(base)

    font_title = _load_font(46)
    font_sub = _load_font(32)
    font_comment = _load_font(30)

    name = star.get("name", "")
    dept = f"{star.get('dept1', '')} {star.get('dept2', '')}".strip()
    award = star.get("award", "")
    comment = star.get("comment", "") or star.get("raw", "")

    draw.text(NAME_POS, f"{name}", font=font_title, fill=(0, 0, 0, 255))
    if dept:
        draw.text(DEPT_POS, dept, font=font_sub, fill=(0, 0, 0, 255))
    if award:
        draw.text(AWARD_POS, award, font=font_sub, fill=(0, 0, 0, 255))
    if month_label:
        draw.text(MONTH_POS, str(month_label), font=font_sub, fill=(0, 0, 0, 255))

    _draw_wrapped_text(
        draw,
        comment,
        font_comment,
        COMMENT_POS,
        max_width=COMMENT_WIDTH,
        fill=(0, 0, 0, 255),
    )

    return base




from PIL import Image, ImageDraw, ImageFont
import io

# 字体：先尝试用仓库里字体文件，没有就用默认
def get_font(size: int):
    try:
        # 如果将来你在仓库里放了自己的字体，比如 fonts/NotoSansSC-Regular.otf
        # 把下面改成那个路径即可
        return ImageFont.truetype("NotoSansSC-Regular.otf", size)
    except Exception:
        return ImageFont.load_default()

AVATAR_SIZE = 200
AVATAR_POS = (100, 200)
NAME_POS = (350, 210)
AWARD_POS = (350, 270)
COMMENT_POS = (100, 350)
COMMENT_WIDTH = 800

def _text_wh(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont):
    """
    用 textbbox 计算文字宽高，兼容新版本 Pillow（已经不推荐 textsize 了）
    """
    if not text:
        return 0, 0
    bbox = draw.textbbox((0, 0), text, font=font)
    w = bbox[2] - bbox[0]
    h = bbox[3] - bbox[1]
    return w, h

def make_initial_avatar(name, size=256):
    """
    没有上传头像时，用姓名首字画一个圆形头像。
    """
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse((0, 0, size, size), fill=(80, 130, 255))

    font = get_font(int(size * 0.45))
    initial = (name or "★")[0]
    w, h = _text_wh(draw, initial, font)
    draw.text(((size - w) / 2, (size - h) / 2), initial, font=font, fill="white")
    return img

def circle(img, size):
    img = img.convert("RGBA").resize((size, size))
    mask = Image.new("L", (size, size), 0)
    d = ImageDraw.Draw(mask)
    d.ellipse((0, 0, size, size), fill=255)
    img.putalpha(mask)
    return img

def wrap(draw, text, font, width):
    """
    用 textbbox 做简单的中文自动换行。
    """
    out, line = [], ""
    for ch in text:
        w, _ = _text_wh(draw, line + ch, font)
        if w <= width:
            line += ch
        else:
            if line:
                out.append(line)
            line = ch
    if line:
        out.append(line)
    return out

def generate_poster(template, star, avatar_img):
    tpl = template.copy().convert("RGBA")
    draw = ImageDraw.Draw(tpl)

    avatar = circle(avatar_img, AVATAR_SIZE)
    tpl.paste(avatar, AVATAR_POS, mask=avatar)

    name_font = get_font(40)
    award_font = get_font(32)
    comment_font = get_font(26)

    name_text = star.get("name", "")
    award_text = star.get("award", "")
    comment_text = star.get("comment") or ""

    draw.text(NAME_POS, name_text, font=name_font, fill="black")
    draw.text(AWARD_POS, award_text, font=award_font, fill="black")

    comment_lines = wrap(draw, comment_text, comment_font, COMMENT_WIDTH)
    x, y = COMMENT_POS
    line_height = 36
    for line in comment_lines:
        draw.text((x, y), line, font=comment_font, fill="black")
        y += line_height

    buf = io.BytesIO()
    tpl.save(buf, format="PNG")
    buf.seek(0)
    return buf, tpl


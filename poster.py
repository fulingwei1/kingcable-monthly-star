from PIL import Image, ImageDraw, ImageFont
import io

# 不再依赖外部字体文件，统一走兜底逻辑
def get_font(size: int):
    try:
        # 如果以后你在仓库里放了合法的字体文件，比如 "NotoSansSC-Regular.otf"
        # 可以把下面这一行改成对应文件名
        return ImageFont.truetype("NotoSansSC-Regular.otf", size)
    except Exception:
        # 找不到就用默认字体，至少保证不崩
        return ImageFont.load_default()

AVATAR_SIZE = 200
AVATAR_POS = (100, 200)
NAME_POS = (350, 210)
AWARD_POS = (350, 270)
COMMENT_POS = (100, 350)
COMMENT_WIDTH = 800

def make_initial_avatar(name, size=256):
    img = Image.new("RGBA", (size, size), (0,0,0,0))
    draw = ImageDraw.Draw(img)
    draw.ellipse((0,0,size,size), fill=(80,130,255))
    font = get_font(int(size * 0.45))
    initial = (name or "★")[0]
    w,h = draw.textsize(initial, font=font)
    draw.text(((size-w)/2,(size-h)/2), initial, font=font, fill="white")
    return img

def circle(img, size):
    img = img.convert("RGBA").resize((size,size))
    mask = Image.new("L", (size, size), 0)
    d = ImageDraw.Draw(mask)
    d.ellipse((0,0,size,size), fill=255)
    img.putalpha(mask)
    return img

def wrap(draw, text, font, width):
    out, line = [], ""
    for ch in text:
        w,_ = draw.textsize(line+ch, font=font)
        if w <= width:
            line += ch
        else:
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

    draw.text(NAME_POS, star["name"], font=name_font, fill="black")
    draw.text(AWARD_POS, star["award"], font=award_font, fill="black")

    comment_text = star.get("comment") or ""
    comment_lines = wrap(draw, comment_text, comment_font, COMMENT_WIDTH)
    x,y = COMMENT_POS
    for line in comment_lines:
        draw.text((x,y), line, font=comment_font, fill="black")
        y += 36

    buf = io.BytesIO()
    tpl.save(buf, format="PNG")
    buf.seek(0)
    return buf, tpl

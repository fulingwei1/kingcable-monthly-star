from PIL import Image, ImageDraw, ImageFont
import io

FONT = "assets/msyh.ttc"   # 你必须放一个中文字体到 assets/

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
    font = ImageFont.truetype(FONT, int(size * 0.45))
    initial = name[0]
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

    # 头像
    avatar = circle(avatar_img, AVATAR_SIZE)
    tpl.paste(avatar, AVATAR_POS, mask=avatar)

    name_font = ImageFont.truetype(FONT, 40)
    award_font = ImageFont.truetype(FONT, 32)
    comment_font = ImageFont.truetype(FONT, 26)

    draw.text(NAME_POS, star["name"], font=name_font, fill="black")
    draw.text(AWARD_POS, star["award"], font=award_font, fill="black")

    comment_lines = wrap(draw, star["comment"], comment_font, COMMENT_WIDTH)
    x,y = COMMENT_POS
    for line in comment_lines:
        draw.text((x,y), line, font=comment_font, fill="black")
        y += 36

    buf = io.BytesIO()
    tpl.save(buf, format="PNG")
    buf.seek(0)
    return buf, tpl

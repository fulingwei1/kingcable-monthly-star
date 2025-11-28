import openpyxl

def load_ws(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    return wb[wb.sheetnames[0]]

def get_month_columns(ws):
    month_map = {}
    row = 2
    col = 6  # F 列开始
    while True:
        val = ws.cell(row=row, column=col).value
        if val is None:
            break
        key = str(val).strip()
        if key:
            month_map[key] = col
        col += 1
    return month_map

def parse_name_award(text):
    t = text.strip()
    first = t.splitlines()[0]

    # 去掉“推荐”
    for p in ["推荐：", "推荐:", "推荐 " ,"推荐"]:
        if first.startswith(p):
            first = first[len(p):]
            break

    # 拆姓名 + 奖项
    for sep in ["：", ":", "-", "－"]:
        if sep in first:
            name, award = first.split(sep, 1)
            return name.strip(), award.strip()

    # 兜底
    if len(first) >= 3:
        return first[:2], first[2:]
    return first, ""

def parse_comment(text):
    t = text.strip()
    if "评语" in t:
        idx = t.find("评语")
        s = t[idx+2:].lstrip("：:").strip()
        return s

    lines = t.splitlines()
    if len(lines) > 1:
        return "\n".join(lines[1:]).strip()
    return t

def extract(ws, col):
    rows = []
    r = 4
    while True:
        seq = ws.cell(r, 1).value
        if seq is None:
            break

        raw = ws.cell(r, col).value
        if raw is None:
            r += 1
            continue
        if not isinstance(raw, str):
            raw = str(raw)

        text = raw.strip()
        if not text or text == "本次暂无":
            r += 1
            continue
        
        dept1 = ws.cell(r, 2).value or ""
        dept2 = ws.cell(r, 3).value or ""

        name, award = parse_name_award(text)
        comment = parse_comment(text)

        rows.append({
            "dept1": str(dept1),
            "dept2": str(dept2),
            "name": name,
            "award": award,
            "comment": comment,
            "raw": text
        })
        r += 1

    return rows

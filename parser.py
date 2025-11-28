import openpyxl
from typing import Dict, List, Any


def load_ws(file):
    """
    从 Streamlit 上传的 xlsx 文件里读取第一个工作表。
    file 可以是路径，也可以是 UploadedFile 对象。
    """
    wb = openpyxl.load_workbook(file, data_only=True)
    return wb[wb.sheetnames[0]]


def get_month_columns(ws) -> Dict[str, int]:
    """
    读取第 1 行，从第 6 列开始，所有包含“月”的单元格，构造成：
    { '2024年11月份': 6, '2024年12月份': 7, '25年3月份': 8, ... }
    """
    header_row = 1
    month_map: Dict[str, int] = {}

    for col in range(6, 50):  # F 列开始向右扫
        val = ws.cell(header_row, col).value
        if not isinstance(val, str):
            continue
        text = val.strip()
        if not text:
            continue
        if "月" in text:
            month_map[text] = col

    return month_map


def parse_name_award(text: str):
    """
    从一整段“推荐 + 奖项”文本中尽量拆出：
    - name: 姓名
    - award: 奖项名称（XXX之星）

    兼容格式：
    - '推荐：张三-敬业之星'
    - '推荐张三：敬业之星'
    - '推荐张三-敬业之星\n评语：……'
    - '张三敬业之星'
    """
    t = text.strip()
    first_line = t.splitlines()[0]

    # 去掉开头的“推荐”
    for prefix in ["推荐：", "推荐:", "推荐 ", "推荐"]:
        if first_line.startswith(prefix):
            first_line = first_line[len(prefix):].strip()
            break

    # 常见分隔符
    for sep in ["：", ":", "-", "－"]:
        if sep in first_line:
            name, award = first_line.split(sep, 1)
            return name.strip("【】 "), award.strip()

    # 兜底：取前 2 字作姓名，后面当奖项
    first_line = first_line.strip("【】 ")
    if len(first_line) >= 3:
        return first_line[:2], first_line[2:]
    return first_line, ""


def parse_comment(text: str) -> str:
    """
    抽取评语：
    - 如果包含“评语”，取“评语”后面的内容
    - 否则，如果有多行，取第 2 行开始
    - 再否则，就返回整段文本
    """
    t = text.strip()
    if "评语" in t:
        idx = t.find("评语")
        sub = t[idx + len("评语") :]
        sub = sub.lstrip("：:").strip()
        return sub

    lines = t.splitlines()
    if len(lines) > 1:
        return "\n".join(l.strip() for l in lines[1:] if l.strip())
    return t


def extract(ws, col: int) -> List[Dict[str, Any]]:
    """
    从指定月份列（col）里抽取所有“每月之星”记录。

    约定：
    - 行 3 开始是数据（A 列序号为 1,2,3…）
    - A 列为空视为数据结束
    - 对应月份列为空 or “本次暂无” → 跳过
    """
    results: List[Dict[str, Any]] = []
    row = 3

    while True:
        seq = ws.cell(row, 1).value
        if seq is None:
            break  # 序号为空，认为到尾部了

        raw = ws.cell(row, col).value
        if raw is None:
            row += 1
            continue

        if isinstance(raw, str):
            text = raw.strip()
        else:
            text = str(raw).strip()

        if (not text) or text == "本次暂无":
            row += 1
            continue

        dept1 = ws.cell(row, 2).value or ""
        dept2 = ws.cell(row, 3).value or ""

        name, award = parse_name_award(text)
        comment = parse_comment(text)

        results.append(
            {
                "row": row,
                "dept1": str(dept1),
                "dept2": str(dept2),
                "name": name,
                "award": award,
                "comment": comment,
                "raw": text,
            }
        )

        row += 1

    return results


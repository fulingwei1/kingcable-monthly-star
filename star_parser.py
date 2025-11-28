import openpyxl
import re
from typing import Dict, List, Any


def load_ws(file):
    """
    从上传的 Excel 文件中读取第一个工作表。
    file 可以是路径，也可以是 Streamlit 的 UploadedFile。
    """
    wb = openpyxl.load_workbook(file, data_only=True)
    return wb[wb.sheetnames[0]]


def get_month_columns(ws) -> Dict[str, int]:
    """
    读取第 1 行，从第 6 列开始，所有包含“月”的单元格，构造成：
    { '2024年11月份': 6, '2024年12月份': 7, '25年10月份': 8, ... }
    """
    header_row = 1
    month_map: Dict[str, int] = {}

    for col in range(6, 60):  # F 列开始往右扫一大段，够用了
        val = ws.cell(header_row, col).value
        if not isinstance(val, str):
            continue
        text = val.strip()
        if text and "月" in text:
            month_map[text] = col

    return month_map


def split_cell_into_people(text: str) -> List[str]:
    """
    将一个单元格内容拆分成多个“个人推荐片段”。

    支持三种人头格式：
    A: 推荐：张三-敬业之星
    B: 张三-敬业之星 / 张三:敬业之星
    C: 张三【敬业之星】

    不再依赖换行，而是在整段字符串中用正则查找所有“人头”的起始位置，
    然后按这些起点把文本切成若干 segment。
    """
    if not isinstance(text, str):
        text = str(text or "")
    text = text.strip()
    if not text:
        return []

    # 人头正则：可选“推荐”，然后 2~4 个汉字名字，后面跟“之星”的奖项
    header_pattern = re.compile(
        r'(推荐[:： ]*)?'                 # 可选的“推荐”
        r'([\u4e00-\u9fff]{2,4})'        # 姓名：2~4 个汉字
        r'\s*'
        r'(?:'
        r'【[^】\n]{0,15}之星】'          # 朱文杰【精准接线之星】
        r'|[-－:：][^，。；\n]{0,15}之星' # 朱文杰-精准接线之星 / ：敬业之星
        r')'
    )

    matches = list(header_pattern.finditer(text))
    if not matches:
        # 完全识别不出人头，就当成一个整体，后面交给 parse_name_award 兜底
        return [text]

    segments: List[str] = []
    for i, m in enumerate(matches):
        start = m.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        seg = text[start:end].strip()
        if seg:
            segments.append(seg)

    return segments


def parse_name_award(text: str):
    """
    从一个“个人片段”里解析姓名和奖项，返回 (name, award)。

    支持三类格式：
    A: 推荐：张三-敬业之星
    B: 张三-敬业之星
    C: 张三【敬业之星】
    """
    t = (text or "").strip()
    if not t:
        return "", ""

    first = t.splitlines()[0].strip()

    # 格式 C：张三【敬业之星】
    m = re.match(r'^([\u4e00-\u9fff]{2,4})【(.+?)】', first)
    if m:
        return m.group(1).strip(), m.group(2).strip()

    # 格式 A：去掉“推荐”
    for prefix in ["推荐：", "推荐:", "推荐 ", "推荐"]:
        if first.startswith(prefix):
            first = first[len(prefix):].strip()
            break

    # 格式 A & B：张三-敬业之星 / 张三：敬业之星
    for sep in ["：", ":", "-", "－"]:
        if sep in first:
            name, award = first.split(sep, 1)
            return name.strip(), award.strip()

    # 兜底：不符合以上格式，就按前 2 字做姓名，剩下做奖项
    first = first.stri

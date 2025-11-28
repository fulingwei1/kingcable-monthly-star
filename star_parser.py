import openpyxl
import re
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


def split_cell_into_people(text: str) -> List[str]:
    """
    将一个单元格内容拆分成多个“个人推荐片段”。

    支持三种人头格式：
    A: 推荐：张三-敬业之星
    B: 张三-敬业之星 / 张三:敬业之星
    C: 张三【敬业之星】

    不再依赖换行，直接在整段 string 上用正则找所有“人头”起始位置，
    然后按这些起点把文本切成若干 segment。
    """
    if not isinstance(text, str):
        text = str(text or "")
    text = text.strip()
    if not text:
        return []

    # 人头正则：可选“推荐”，然后 2~4 个汉字的人名，后面跟“之星”类的奖项
    # 例如：
    #   推荐：朱文杰-精准接线之星
    #   朱文杰-精准接线之星
    #   朱文杰【精准接线之星】
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
        # 完全识别不出人头，就当成一个整体，后面再由 parse_name_award 自己想办法
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
    从人头行中解析姓名和奖项，支持三类格式：
    A: 推荐：张三-敬业之星
    B: 张三-敬业之星
    C: 张三【敬业之星】
    """
    t = (text or "").strip()
    if not t:
        return "", ""

    first = t.splitlines

def extract(ws, col: int) -> List[Dict[str, Any]]:
    """
    从指定月份列（col）里抽取所有“每月之星”记录。

    约定：
    - 行 3 开始是数据（A 列序号为 1,2,3…）
    - A 列为空视为数据结束
    - 对应月份列为空 or “本次暂无” → 跳过
    - 如果一个格子里推荐了多个人，拆成多条记录。
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

        text = str(raw).strip()
        if (not text) or text == "本次暂无":
            row += 1
            continue

        dept1 = ws.cell(row, 2).value or ""
        dept2 = ws.cell(row, 3).value or ""

        # 这里拆多人
        segments = split_cell_into_people(text)
        if not segments:
            row += 1
            continue

        for seg in segments:
            name, award = parse_name_award(seg)
            comment = parse_comment(seg)

            if not name or name in ("推荐", "评语"):
                continue

            results.append(
                {
                    "row": row,
                    "dept1": str(dept1),
                    "dept2": str(dept2),
                    "name": name,
                    "award": award,
                    "comment": comment,
                    "raw": seg,  # 每条记录只保留属于自己那一段的原始文本
                }
            )

        row += 1

    return results


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

    # F 列开始往右扫一大段，够用
    for col in range(6, 80):
        val = ws.cell(header_row, col).value
        if not isinstance(val, str):
            continue
        text = val.strip()
        if text and "月" in text:
            month_map[text] = col

    return month_map


def _split_cell_into_people(text: str) -> List[str]:
    """
    将一个单元格内容拆分成多个“个人推荐片段”。

    支持几类人头格式（可以混用）：
    - 推荐：张三-敬业之星
    - 推荐：张三-核心技术骨干
    - 张三-敬业之星 / 张三:敬业之星 / 张三：敬业之星
    - 张三【敬业之星】 / 张三【核心技术骨干】

    逻辑：在整段文本中找到所有“人头”的起始位置，再按照这些起点切分。
    """
    if not isinstance(text, str):
        text = str(text or "")
    text = text.strip()
    if not text:
        return []

    header_pattern = re.compile(
        r'(推荐[:： ]*)?'                 # 可选的“推荐”
        r'([\u4e00-\u9fff]{2,4})'        # 姓名：2~4 个汉字
        r'\s*'
        r'(?:'
        r'【[^】\n]{0,30}】'              # 张三【敬业之星】 / 张三【核心技术骨干】
        r'|[-－:：][^，。；\n]{0,30}'      # 张三-敬业之星 / 张三-核心技术骨干 / 张三：敬业之星
        r')'
    )

    matches = list(header_pattern.finditer(te_

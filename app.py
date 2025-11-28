import streamlit as st
from PIL import Image
from io import BytesIO

from star_parser import load_ws, get_month_columns, extract
from poster import generate_poster, make_initial_avatar


st.set_page_config(page_title="金凯博自动化每月之星海报生成器", layout="wide")
st.title("金凯博自动化 · 每月之星海报生成系统")


# ---------------- 1. 上传文件 ----------------
st.subheader("① 上传 Excel & 海报模板")

excel_file = st.file_uploader(
    "上传《金凯博自动化-每月之星》Excel 文件", type=["xlsx"], key="excel_file"
)
template_file = st.file_uploader(
    "上传海报模板 PNG/JPG（透明区域留给头像和文字）",
    type=["png", "jpg", "jpeg"],
    key="poster_file",
)

st.divider()

# ---------------- 2. 选择月份 → 动态解析当月之星 ----------------
st.subheader("② 选择月份并查看对应的每月之星")

if excel_file is None:
    st.info("先上传 Excel 文件。")
else:
    ws = load_ws(excel_file)
    month_map = get_month_columns(ws)

    if not month_map:
        st.error("在第 1 行没有识别到任何“*月*”的列，请检查表头。")
    else:
        month_names = list(month_map.keys())
        default_index = max(len(month_names) - 1, 0)

        selected_month = st.selectbox(
            "选择月份", month_names, index=default_index, key="month_select"
        )

        stars = []
        if selected_month:
            col = month_map[selected_month]
            stars = extract(ws, col)

        st.session_state["stars"] = stars
        st.session_state["selected_month"] = selected_month

        if not stars:
            st.warning(f"这一列没有识别到任何每月之星。")
        else:
            st.success(f"已识别出 {len(stars)} 位每月之星：")

            for i, star in enumerate(stars):
                header = f"{star['dept1']}/{star['dept2']} · {star['name']}（{star['award']}）"
                with st.expander(header, expanded=False):
                    st.write("**部门：**", f"{star['dept1']} / {star['dept2']}")
                    st.write("**原始文本：**")
                    st.code(star["raw"], language="text")

                    default_comment = (star.get("comment") or "").strip() or star["raw"]
                    new_comment = st.text_area(
                        "评语（可修改）",
                        value=default_comment,
                        key=f"comment_{i}",
                    )
                    star["comment"] = new_comment

                    avatar_file = st.file_uploader(
                        "上传头像（可选）",
                        type=["png", "jpg", "jpeg"],
                        key=f"avatar_{i}",
                    )
                    if avatar_file is not None:
                        star["avatar_bytes"] = avatar_file.read()
                        st.image(star["avatar_bytes"], width=80)
                    else:
                        star["avatar_bytes"] = None

st.divider()

# ---------------- 3. 生成海报 ----------------
st.subheader("③ 生成单人海报预览")

stars = st.session_state.get("stars") or []
selected_month = st.session_state.get("selected_month")

if not stars:
    st.info("先在上面选好月份，并确认已经识别出当月的每月之星。")
elif template_file is None:
    st.info("请先上传海报模板图片。")
else:
    template_img = Image.open(template_file).convert("RGBA")
    names = [f"{s['name']}（{s['award']}）" for s in stars]
    idx = st.selectbox(
        "选择要生成海报的员工", range(len(stars)), format_func=lambda i: names[i]
    )

    star = stars[idx]

    if star.get("avatar_bytes"):
        avatar_img = Image.open(BytesIO(star["avatar_bytes"])).convert("RGBA")
    else:
        avatar_img = make_initial_avatar(star["name"])

    poster_img = generate_poster(template_img, avatar_img, star, selected_month)

    st.image(poster_img, caption="海报预览", use_column_width=True)

    buf = BytesIO()
    poster_img.save(buf, format="PNG")
    buf.seek(0)
    st.download_button(
        "下载 PNG 海报",
        data=buf,
        file_name=f"{selected_month}_{star['name']}_每月之星.png",
        mime="image/png",
    )


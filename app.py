import streamlit as st
from PIL import Image

from parser import load_ws, get_month_columns, extract
from poster import generate_poster, make_initial_avatar

st.set_page_config(layout="wide", page_title="金凯博自动化 每月之星系统")

st.title("金凯博自动化 · 每月之星海报生成系统")

xlsx = st.file_uploader("上传《金凯博自动化-每月之星》Excel 文件", type=["xlsx"])
template = st.file_uploader("上传海报模板 PNG", type=["png"])

if xlsx and template:
    ws = load_ws(xlsx)
    template_img = Image.open(template)
    months = get_month_columns(ws)

    month = st.selectbox("选择月份", list(months.keys()))
    col = months[month]

    stars = extract(ws, col)

    st.subheader("解析结果")
    st.write(f"共找到 {len(stars)} 位每月之星")

    avatars = {}
    results = []

    for i, star in enumerate(stars):
        with st.expander(f"{star['dept1']} / {star['dept2']} - {star['name']}（{star['award']}）"):
            st.text_area("评语（可修改）", key=f"comment_{i}", value=star["comment"], height=120)
            file = st.file_uploader("上传头像（可选）", type=["png","jpg"], key=f"avatar_{i}")
            if file:
                avatars[i] = Image.open(file)
                st.image(avatars[i], width=120)

    if st.button("生成海报"):
        for i, star in enumerate(stars):
            star["comment"] = st.session_state[f"comment_{i}"]

            avatar = avatars.get(i)
            if avatar is None:
                avatar = make_initial_avatar(star["name"], 256)

            buf, img = generate_poster(template_img, star, avatar)
            st.image(img, caption=f"{star['name']} - {star['award']}")
            st.download_button(f"下载 {star['name']}.png", buf, file_name=f"{star['name']}.png")

import streamlit as st
from PIL import Image

from parser import load_ws, get_month_columns, extract
from poster import generate_poster, make_initial_avatar   # 你已有的头像/海报模块


st.set_page_config(layout="wide", page_title="金凯博自动化 每月之星系统")

st.title("金凯博自动化 · 每月之星海报生成系统")

# 1. 上传 Excel 与模板
xlsx = st.file_uploader("上传《金凯博自动化-每月之星》Excel 文件", type=["xlsx"])
template_file = st.file_uploader("上传海报模板 PNG/JPG", type=["png", "jpg", "jpeg"])

if not xlsx:
    st.info("请先上传 Excel 文件。")
    st.stop()

ws = load_ws(xlsx)
month_map = get_month_columns(ws)

if not month_map:
    st.error("在第 1 行没有识别到任何包含“月”的表头，请检查表格。")
    st.stop()

# 2. 选择月份
st.subheader("① 选择月份并查看对应的每月之星")

month = st.selectbox("选择月份", list(month_map.keys()))
col = month_map[month]

stars = extract(ws, col)

# 调试信息，可以先打开看结果是否合理，如果觉得丑可以关掉
with st.expander("调试信息（你确认无误后可以删除这块）", expanded=False):
    st.write("月份 → 列号映射：", month_map)
    st.write(f"当前选择：{month} → Excel 列号 {col}")
    st.write("本月抽取到的员工：", [f"{s['name']}({s['dept1']}/{s['dept2']})" for s in stars])

if not stars:
    st.warning(f"{month} 没有任何获奖记录（要么为空，要么写了“本次暂无”）。")
    st.stop()

# 3. 展示当前月份的每月之星列表
st.subheader(f"② {month} · 每月之星列表")

avatars = {}
for i, star in enumerate(stars):
    with st.expander(f"{star['dept1']} / {star['dept2']} - {star['name']}（{star['award']}）", expanded=False):
        st.markdown(f"**部门：** {star['dept1']} / {star['dept2']}")
        st.markdown(f"**原始文本：** {star['raw']}")
        star["comment"] = st.text_area(
            "评语（可修改）",
            key=f"comment_{i}",
            value=star["comment"],
            height=120
        )
        avatar_file = st.file_uploader(
            "上传头像（可选）",
            type=["png", "jpg", "jpeg"],
            key=f"avatar_{i}"
        )
        if avatar_file:
            img = Image.open(avatar_file)
            avatars[i] = img
            st.image(img, width=100)

# 4. 生成海报（这里沿用你已有的 poster.generate_poster 逻辑）
st.subheader("③ 生成海报预览与下载")

if not template_file:
    st.info("如需生成海报，请上传一张海报模板图片。")
else:
    template_img = Image.open(template_file)

    if st.button("生成当前月份所有获奖者的海报"):
        for i, star in enumerate(stars):
            # 从 text_area 里拿最终评语
            star["comment"] = st.session_state.get(f"comment_{i}", star["comment"])

            avatar = avatars.get(i)
            if avatar is None:
                avatar = make_initial_avatar(star["name"], 256)

            buf, img = generate_poster(template_img, star, avatar)
            st.image(img, caption=f"{star['name']} - {star['award']}")
            st.download_button(
                label=f"下载 {star['name']}.png",
                data=buf,
                file_name=f"{month}_{star['name']}.png",
                mime="image/png"
            )


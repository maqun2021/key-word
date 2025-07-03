import streamlit as st
import pandas as pd
import jieba
from wordcloud import WordCloud
import matplotlib.pyplot as plt
from io import BytesIO

st.set_page_config(page_title="关键词云分析", layout="wide")
st.title("📊 关键词云分析工具（Excel数据增强版）")

# 1. 文件上传
data_file = st.file_uploader("请上传Excel文件（.xlsx）", type=["xlsx"])

if data_file:
    df = pd.read_excel(data_file)
    st.write("数据预览：", df.head())
    columns = df.columns.tolist()
    
    # 2. 选择关键词分析列
    key_col = st.selectbox("请选择要进行关键词云分析的列：", columns)
    
    # 3. 选择辅助筛选维度
    filter_cols = st.multiselect("可选：选择其他列作为筛选维度（如负责人、日期等）", [col for col in columns if col != key_col])
    filters = {}
    for col in filter_cols:
        options = df[col].dropna().unique().tolist()
        selected = st.multiselect(f"筛选 {col} 的值：", options)
        if selected:
            filters[col] = selected
    
    # 4. 数据筛选
    filtered_df = df.copy()
    for col, vals in filters.items():
        filtered_df = filtered_df[filtered_df[col].isin(vals)]
    
    # 5. 关键词提取
    text_data = filtered_df[key_col].dropna().astype(str).tolist()
    full_text = " ".join(text_data)
    words = jieba.lcut(full_text)
    # 去除常见无意义词
    stopwords = set(["的", "了", "和", "是", "在", "就", "都", "而", "及", "与", "着", "或", "一个", "没有", "我们", "你们", "他们", "以及", "其", "并", "但", "被", "为", "更", "到", "这", "与", "对", "中", "上", "下", "等", "与", "及", "也", "还", "要", "等", "等", "等"])
    words = [w for w in words if w.strip() and w not in stopwords]
    word_str = " ".join(words)
    
    # 6. 生成词云
    if word_str:
        wc = WordCloud(font_path="msyh.ttc", width=800, height=400, background_color="white").generate(word_str)
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.imshow(wc, interpolation="bilinear")
        ax.axis("off")
        st.pyplot(fig)
        
        # 7. 导出图片
        buf = BytesIO()
        fig.savefig(buf, format="png")
        st.download_button("下载词云图片", data=buf.getvalue(), file_name="wordcloud.png", mime="image/png")
    else:
        st.info("没有足够的关键词生成词云，请检查数据或筛选条件。")
else:
    st.info("请先上传Excel文件。") 
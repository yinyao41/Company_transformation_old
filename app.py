# -*- coding: utf-8 -*-
import streamlit as st
import requests
from docx import Document
from io import BytesIO
from openai import OpenAI
import os

# =============================================================================
# 配置区
# =============================================================================
GITHUB_USERNAME = "yinyao41"
GITHUB_REPO = "Company_transformation"
BRANCH = "master"

# 只使用 data 目录下的一个文件（已移除提示词文件）
TEMPLATE_FILES = [
    "data/转型升级提示词内容 - 1.docx",
]

# =============================================================================
# 精简系统提示词（只参考两个文件模板格式）
# =============================================================================
SYSTEM_PROMPT = """你是一位专业的公司转型升级咨询专家。
请严格按照「转型升级提示词内容 - 1.docx」中的六套方案决策矩阵、输出格式（一页纸决策单 + 详细实施报告）和所有规则要求，
为用户提供的公司生成一份完整、简洁、可执行的转型升级方案报告。

"""

# =============================================================================
# 阿里通义千问客户端
# =============================================================================
DASHSCOPE_API_KEY = st.secrets.get("DASHSCOPE_API_KEY", os.getenv("DASHSCOPE_API_KEY"))
if not DASHSCOPE_API_KEY:
    st.error("缺少 DASHSCOPE_API_KEY！请在 Streamlit Secrets 中添加")
    st.stop()

client = OpenAI(
    api_key=DASHSCOPE_API_KEY,
    base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
)

MODEL_NAME = "qwen-max"

# =============================================================================
# 加载两个模板文件（静默处理 + 强力截断）
# =============================================================================
@st.cache_data(show_spinner="正在加载一个模板文件...")
def load_templates():
    templates = []
    for rel_path in TEMPLATE_FILES:
        raw_url = f"https://raw.githubusercontent.com/{GITHUB_USERNAME}/{GITHUB_REPO}/{BRANCH}/{rel_path}"
        try:
            r = requests.get(raw_url, timeout=12)
            r.raise_for_status()
            doc = Document(BytesIO(r.content))
            text = "\n".join(p.text.strip() for p in doc.paragraphs if p.text.strip())
            if text:
                name = rel_path.split("/")[-1].replace(".docx", "")
                templates.append(f"【{name}】\n{text}\n{'─' * 60}\n")
        except:
            continue
    full_text = "".join(templates)
    if len(full_text) > 18000:
        full_text = full_text[:18000] + "\n\n【模板已自动截断】"
    return full_text

TEMPLATES_TEXT = load_templates()

# =============================================================================
# Streamlit 界面
# =============================================================================
st.set_page_config(page_title="公司转型升级方案生成器", layout="wide")
st.title("🏭 公司转型升级方案生成器")

with st.form(key="company_info_form"):
    company_name = st.text_input("公司名称*", placeholder="例如：山东固丰体育产业有限公司")
    industry = st.text_input("所属行业*", placeholder="例如：体育产业")
    current_status = st.text_area("公司当前情况描述*", placeholder="描述公司规模、问题、优势等...", height=150)
    additional_info = st.file_uploader("上传公司相关文件（可选）", type=["pdf", "docx", "txt"])
    
    submit_button = st.form_submit_button(label="生成转型升级方案")

if submit_button:
    if not company_name or not industry or not current_status:
        st.error("请填写带*的必填项！")
    else:
        extra_text = ""
        if additional_info:
            try:
                if additional_info.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                    doc = Document(BytesIO(additional_info.read()))
                    extra_text = "\n".join(p.text.strip() for p in doc.paragraphs if p.text.strip())
                else:
                    extra_text = additional_info.read().decode("utf-8")
            except:
                st.warning("补充文件解析失败，将使用文字描述")

        user_context = f"""
公司名称：{company_name}
所属行业：{industry}
当前情况：{current_status}
补充材料：{extra_text}
"""

        with st.spinner("正在调用 AI 生成方案..."):
            try:
                response = client.chat.completions.create(
                    model=MODEL_NAME,
                    messages=[
                        {"role": "system", "content": SYSTEM_PROMPT + "\n\n以下是一个模板文件内容：\n" + TEMPLATES_TEXT},
                        {"role": "user", "content": f"基于模板，为以下公司生成完整转型升级方案：\n{user_context}"}
                    ],
                    temperature=0.3,
                    max_tokens=2800,
                    stream=False
                )
                scheme = response.choices[0].message.content
                st.success("✅ 生成完成！")
                st.markdown(scheme)

                st.download_button(
                    label="📥 下载方案（Markdown）",
                    data=scheme,
                    file_name=f"{company_name}_转型升级方案.md",
                    mime="text/markdown"
                )
            except Exception as e:
                error_str = str(e).lower()
                if "context length" in error_str or "maximum" in error_str:
                    st.error("提示词过长，请简化公司描述后重试")
                else:
                    st.error(f"AI 调用失败：{str(e)}")

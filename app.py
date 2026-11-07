import streamlit as st
import requests
from docx import Document
from io import BytesIO
from openai import OpenAI
import os

# 用于解析 PDF —— 需要 requirements.txt 中包含 pypdf
try:
    from pypdf import PdfReader
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

# =============================================================================
# 配置区
# =============================================================================
DASHSCOPE_API_KEY = st.secrets.get("DASHSCOPE_API_KEY", os.getenv("DASHSCOPE_API_KEY"))
if not DASHSCOPE_API_KEY:
    st.error("缺少 DASHSCOPE_API_KEY！请在 Streamlit Secrets 中添加")
    st.stop()

client = OpenAI(
    api_key=DASHSCOPE_API_KEY,
    base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
)

MODEL_NAME = "qwen-plus"  # 推荐使用 plus 更快，必要时改回 qwen-max

# =============================================================================
# 长度控制（避免触发 "Range of input length should be [1, 1000000]" 报错）
# =============================================================================
# 原报错的根因：旧代码对非 docx 文件（例如 PDF）直接用 utf-8 解码二进制内容，
# 一个几 MB 的 PDF 解码后会产生几十万到上百万字符的乱码文本，导致总输入长度超限。
# 这里改为：① 用正确的库解析每种文件类型的文本 ② 对每个文件、以及总输入做长度截断。
MAX_CHARS_PER_FILE = 30000       # 每个补充文件最多保留的字符数
MAX_TOTAL_EXTRA_CHARS = 120000   # 所有补充文件合计最多保留的字符数
MAX_TOTAL_PROMPT_CHARS = 900000  # 系统提示词 + 用户输入 总长度上限（留出安全余量，API 上限是 1,000,000）

# =============================================================================
# 八套转型升级方向文档路径（请把新版 docx 放入 data 目录，并与下面文件名保持一致）
# =============================================================================
SCHEME_DOC_NAME = "附件2-4-1 八套转型升级方向.202605.docx"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCHEME_DOC_CANDIDATES = [
    os.path.join("data", SCHEME_DOC_NAME),
    os.path.join(BASE_DIR, "data", SCHEME_DOC_NAME),
]

FULL_SYSTEM_PROMPT = """【角色设定】
你是一位资深的产业战略咨询专家，专注于传统企业转型升级与科创产业融合领域。

【重要免责声明】
请在输出的最前面明确添加以下提示语句：
**重要提示：本方向仅供参考，不构成任何正式的投资、经营或法律建议。实际操作请咨询专业律师、财务顾问及相关行业专家。**

【八套转型升级方向库】
以下是完整的八套方向内容（已从文档中读取）：
"""

# 自动加载文档内容并拼接到 Prompt
try:
    scheme_doc_path = next((p for p in SCHEME_DOC_CANDIDATES if os.path.exists(p)), None)
    if scheme_doc_path:
        doc = Document(scheme_doc_path)
        doc_text = "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
        FULL_SYSTEM_PROMPT += doc_text
        st.success(" ")
    else:
        FULL_SYSTEM_PROMPT += "（方向文档加载失败，请检查 data 目录）"
except Exception as e:
    FULL_SYSTEM_PROMPT += f"（方向文档加载失败: {str(e)}）"

FULL_SYSTEM_PROMPT += """
【输出要求】
请严格按照以下格式为用户公司生成**八套**专属转型升级方向，八套方向依次为：
方向一：并购科创企业；方向二：产研合作；方向三：数智化改造；方向四：渠道合作；
方向五：精益管理；方向六：全球化出海布局；方向七：投资科创企业；方向八：投资科创基金。

0. 先输出"公司画像（简要提炼）"，包含：所属行业、核心业务、当前痛点/瓶颈、可用资源。
1. 八套方向必须逐一输出，不得遗漏、合并或新增，每个方向的标题格式为：
   "方向X：方向名称，实现[贴合该公司的战略目标]"。
2. 结构要求（每个方向必须包含以下4个固定模块）：
   - 方向核心定位（1-2句话）
   - [公司名]专属落地方向（3-4条具体路径，结合公司实际，具体到业务板块、产品品类、客户群体）
   - 转型核心优势（4条）
   - 核心实施要点（4条）
3. 必须深度定制，结合公司行业、业务、资源、痛点，禁止泛泛而谈。
4. 语言专业，使用"第二曲线"、"生态闭环"、"耐心资本"、"精益红利"等产业术语。
"""


# =============================================================================
# 文件解析函数
# =============================================================================
def extract_text_from_docx(file_bytes: bytes) -> str:
    doc = Document(BytesIO(file_bytes))
    return "\n".join(p.text.strip() for p in doc.paragraphs if p.text.strip())


def extract_text_from_pdf(file_bytes: bytes) -> str:
    if not PDF_SUPPORT:
        raise RuntimeError("未安装 pypdf，无法解析 PDF。请在 requirements.txt 中添加 pypdf 后重新部署。")
    reader = PdfReader(BytesIO(file_bytes))
    texts = []
    for page in reader.pages:
        page_text = page.extract_text() or ""
        if page_text.strip():
            texts.append(page_text.strip())
    return "\n".join(texts)


def extract_text_from_txt(file_bytes: bytes) -> str:
    return file_bytes.decode("utf-8", errors="ignore")


def parse_uploaded_file(uploaded_file) -> str:
    """根据文件类型解析文本，解析失败时抛出异常，由调用方捕获并提示用户。"""
    filename = uploaded_file.name.lower()
    file_bytes = uploaded_file.read()

    if filename.endswith(".docx"):
        return extract_text_from_docx(file_bytes)
    elif filename.endswith(".pdf"):
        return extract_text_from_pdf(file_bytes)
    elif filename.endswith(".txt"):
        return extract_text_from_txt(file_bytes)
    else:
        raise ValueError(f"不支持的文件类型：{uploaded_file.name}")


def truncate(text: str, max_chars: int) -> tuple[str, bool]:
    """截断文本，返回 (截断后的文本, 是否发生了截断)"""
    if len(text) > max_chars:
        return text[:max_chars], True
    return text, False


# =============================================================================
# Streamlit 界面
# =============================================================================
st.set_page_config(page_title="转型升级方向", layout="wide")
st.title("🏭 企业转型升级方向生成器")
st.caption("✅")

DEFAULT_CURRENT_STATUS = """宁波市博华机械配件有限公司成立于2009年，位于浙江省宁波市奉化区，是一家以铝合金铸造及机械精加工为核心业务的传统制造企业。公司具备从模具设计、铝合金重力铸造、翻砂铸造，到CNC加工、数控车削、热处理及检测等较为完整的生产能力，可根据客户图纸提供毛坯件及成品零部件。

公司产品主要包括汽车及摩托车零部件、减震器支架、气动元器件、火车零件、电机及风机配件、石化设备零件、空压机零件和医疗器械零件等，下游应用领域较为分散。公开资料显示，公司产品除供应国内市场外，还出口法国、英国、美国等海外市场，并与安徽中鼎、沃尔沃等客户存在配套合作。

从生产规模看，公司现有员工约90余人，拥有数控加工中心、数控车床、铸造设备、喷砂设备、热处理炉等生产设备，并配置光谱分析仪、三坐标测量仪、X光检测仪等质量检测设备，具备一定的规模化生产及品质管控能力。公开资料披露，公司年产能超过2,000吨，近年销售规模处于亿元左右，其中招聘平台披露其2024年营业收入约5,486万元、年销售额超过8,500万元。

整体来看，博华机械属于典型的"铸造+机械加工"一体化中小型制造企业，业务技术路线成熟，传统制造属性较强，主要竞争力体现在客户定制能力、铸造与机加工一体化能力以及多年积累的生产工艺和客户资源。"""

with st.form(key="company_info_form"):
    company_name = st.text_input("公司名称*", value="宁波市博华机械配件有限公司", placeholder="例如：山东固丰体育产业有限公司")
    industry = st.text_input("所属行业*", value="机械加工", placeholder="例如：体育产业")
    current_status = st.text_area(
        "公司当前情况描述*",
        value=DEFAULT_CURRENT_STATUS,
        placeholder="描述公司规模、问题、优势、核心业务、渠道、品牌等...",
        height=200,
    )
    # 支持多个补充文件
    additional_files = st.file_uploader(
        "额外上传补充文件（可选，支持多选，pdf / docx / txt）",
        type=["pdf", "docx", "txt"],
        accept_multiple_files=True,
    )
    submit_button = st.form_submit_button(label="生成八套转型升级方向")

if submit_button:
    if not company_name or not industry or not current_status:
        st.error("请填写带*的必填项！")
    else:
        extra_text_blocks = []
        parse_warnings = []
        truncate_warnings = []

        if additional_files:
            for f in additional_files:
                try:
                    text = parse_uploaded_file(f)
                    if not text.strip():
                        parse_warnings.append(f"「{f.name}」未提取到有效文本内容（可能是扫描件/图片型PDF），已跳过。")
                        continue
                    text, was_truncated = truncate(text, MAX_CHARS_PER_FILE)
                    if was_truncated:
                        truncate_warnings.append(f"「{f.name}」内容较长，已截取前 {MAX_CHARS_PER_FILE} 字符。")
                    extra_text_blocks.append(f"【补充文件：{f.name}】\n{text}")
                except Exception as e:
                    parse_warnings.append(f"「{f.name}」解析失败：{str(e)}")

        # 合并所有补充文件文本，并对总量再做一次截断
        extra_text = "\n\n".join(extra_text_blocks)
        extra_text, extra_total_truncated = truncate(extra_text, MAX_TOTAL_EXTRA_CHARS)
        if extra_total_truncated:
            truncate_warnings.append(f"补充文件总内容超出 {MAX_TOTAL_EXTRA_CHARS} 字符上限，已整体截断。")

        # 展示解析情况
        if parse_warnings:
            for w in parse_warnings:
                st.warning(f"⚠️ {w}")
        if truncate_warnings:
            for w in truncate_warnings:
                st.info(f"ℹ️ {w}")
        if additional_files:
            success_count = len(extra_text_blocks)
            st.caption(f"共上传 {len(additional_files)} 个文件，成功解析 {success_count} 个。")

        user_context = f"""
公司名称：{company_name}
所属行业：{industry}
当前情况：{current_status}
补充材料：{extra_text if extra_text else "（无）"}
"""

        # 对系统提示词 + 用户输入的总长度做最终兜底检查，避免触发 API 的
        # "Range of input length should be [1, 1000000]" 报错
        total_len = len(FULL_SYSTEM_PROMPT) + len(user_context)
        if total_len > MAX_TOTAL_PROMPT_CHARS:
            overflow = total_len - MAX_TOTAL_PROMPT_CHARS
            # 优先压缩补充材料部分，而不是公司描述或方向库
            new_extra_len = max(0, len(extra_text) - overflow)
            extra_text = extra_text[:new_extra_len]
            user_context = f"""
公司名称：{company_name}
所属行业：{industry}
当前情况：{current_status}
补充材料：{extra_text if extra_text else "（因总长度超限已省略）"}
"""
            st.info("ℹ️ 输入总长度接近模型上限，已自动精简补充材料部分以确保正常生成。")

        with st.spinner("正在生成八套专属方向...（约 30-60 秒）"):
            try:
                response = client.chat.completions.create(
                    model=MODEL_NAME,
                    messages=[
                        {"role": "system", "content": FULL_SYSTEM_PROMPT},
                        {"role": "user", "content": f"请为以下公司生成完整转型升级八套专属方向：\n{user_context}"},
                    ],
                    temperature=0.3,
                    max_tokens=8000,
                    stream=False,
                )
                scheme = response.choices[0].message.content

                st.success("✅ 生成完成！")
                st.warning("**重要提示：本方向仅供参考，不构成任何正式的投资、经营或法律建议。**")
                st.markdown(scheme)

                st.download_button(
                    label="📥 下载方案（Markdown）",
                    data=scheme,
                    file_name=f"{company_name}_转型升级八套方向.md",
                    mime="text/markdown",
                )
            except Exception as e:
                st.error(f"AI 调用失败：{str(e)}")

import streamlit as st
from docx import Document
from io import BytesIO
from openai import OpenAI
import os
import re

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

MODEL_NAME = "qwen-plus"          # 推荐 plus；需要更强推理可改 qwen-max
MAX_OUTPUT_TOKENS = 8000          # 八套方向 + 公司画像需要足够输出空间

# =============================================================================
# 长度控制（防止触发 input length 超限）
# =============================================================================
MAX_CHARS_PER_FILE = 30000
MAX_TOTAL_EXTRA_CHARS = 120000
MAX_TOTAL_PROMPT_CHARS = 900000

# =============================================================================
# 八套转型升级方向文档路径
# =============================================================================
SCHEME_DOC_NAME = "附件2-4-1 八套转型升级方向.202605-2.docx"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCHEME_DOC_CANDIDATES = [
    os.path.join("data", SCHEME_DOC_NAME),
    os.path.join(BASE_DIR, "data", SCHEME_DOC_NAME),
    os.path.join(BASE_DIR, SCHEME_DOC_NAME),
]

# =============================================================================
# 系统提示词（严格按文档要求构建）
# =============================================================================
FULL_SYSTEM_PROMPT = """【角色设定】
你是一位资深的产业战略咨询专家，专注于传统企业转型升级与科创产业融合领域。你擅长将通用的转型升级方法论与企业实际情况深度结合，输出具有高度针对性、可落地性的战略方向。请根据以下**八套**转型升级方向（每案含3个实战案例），针对用户输入的公司，逐一分析每套方向的适用性，并给出具体建议。

【重要免责声明】
请在输出的最前面明确添加以下提示语句：
**重要提示：本方向仅供参考，不构成任何正式的投资、经营或法律建议。实际操作请咨询专业律师、财务顾问及相关行业专家。**

【八套转型升级方向库】
以下是完整的八套方向内容（已从文档中读取）：
"""

# 自动加载文档内容
try:
    scheme_doc_path = next((p for p in SCHEME_DOC_CANDIDATES if os.path.exists(p)), None)
    if scheme_doc_path:
        doc = Document(scheme_doc_path)
        doc_text = "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
        FULL_SYSTEM_PROMPT += doc_text
    else:
        FULL_SYSTEM_PROMPT += "（方向文档加载失败，请检查 data 目录下是否存在对应 docx 文件）"
except Exception as e:
    FULL_SYSTEM_PROMPT += f"（方向文档加载失败: {str(e)}）"

FULL_SYSTEM_PROMPT += """

【输出要求】（必须逐条严格遵守，不得增减、合并或调换模块）

一、整体结构（按顺序输出，不得增删、合并或调换）
1. 最开头原样输出上面的免责声明。
2. 随后输出“## 公司画像（简要提炼）”，严格用以下4个要点概括：
   - **所属行业：**
   - **核心业务：**
   - **当前痛点/瓶颈：**
   - **可用资源：**
   内容必须严格依据用户提供的信息提炼，信息不足处标注“（待补充）”，不得编造事实或数据。
3. 然后依次输出八套专属方向，顺序与名称固定如下：
   方向一：并购科创企业
   方向二：产研合作
   方向三：数智化改造
   方向四：渠道合作
   方向五：精益管理
   方向六：全球化出海布局
   方向七：投资科创企业
   方向八：投资科创基金
4. 文档中“请根据以上8套方向，一一对应分析并给出……”是对你的指令，不要原样输出。

二、每套方向的固定格式（严格包含4个模块，名称不得更改）
### 方向X：方向名称，实现[结合该公司实际的战略目标，一句话]
**方向核心定位**
（1-2句话，说明该方向对该公司的战略意义，必须贴合企业实际情况）
**[公司简称]专属落地方向**
1. ……
2. ……
3. ……（共3-4条，必须具体到业务板块、产品品类、客户群体）
**转型核心优势**
1. ……
2. ……
3. ……
4. ……（必须4条）
**核心实施要点**
1. ……
2. ……
3. ……
4. ……（必须4条）
每个模块内的条目序号都从1开始，不要接着上一个模块继续编号。

三、内容深度与风格要求
1. 必须深度融合该公司的产业特点、核心资源、渠道网络、区域地位等实际情况；每条落地方向都要具体到业务板块、产品品类、客户群体，禁止泛泛而谈。
2. 方向库文档中带花括号{}的内容以及“恒力机械”等示例企业，仅作为格式和深度的参考，禁止照抄，输出中不得出现花括号，必须基于用户公司重新撰写。
3. 涉及金额、比例等量化目标时，应基于用户提供的数据合理测算，并使用“约”“预计”“目标”等措辞，不得写成既成事实，不得编造用户未提供的财务数据。
4. 方向七（投资科创企业）：必须体现“资方直投科创企业、采用专项基金形式”，并引导选用同祺担任基金管理人（负责项目筛选、投资决策和全周期投后管理）。必须明确区分“产业相关型”和“产业不相关型”，并分别给出落地路径。
5. 方向八（投资科创基金）：必须体现“资方以LP身份参与盲池基金”，并引导选用同祺担任基金管理人（GP）；可结合基金规模、期限、管理费、超额收益分配等要素。
6. 语言专业严谨，使用“第二曲线”“生态闭环”“耐心资本”“精益红利”等产业术语，突出企业的品牌势能、区域龙头地位等差异化特征（若用户未提供则据实表述，不要虚构）。
7. 篇幅控制：每条落地方向、优势、要点尽量控制在80字以内，确保公司画像与八套方向全部完整输出，不得因篇幅而省略任何方向或模块。
"""

# =============================================================================
# 输出格式校验
# =============================================================================
EXPECTED_DIRECTIONS = [
    ("一", "并购科创企业"),
    ("二", "产研合作"),
    ("三", "数智化改造"),
    ("四", "渠道合作"),
    ("五", "精益管理"),
    ("六", "全球化出海布局"),
    ("七", "投资科创企业"),
    ("八", "投资科创基金"),
]

MODULE_RULES = [
    ("方向核心定位", None, None),
    ("专属落地方向", 3, 4),
    ("转型核心优势", 4, 4),
    ("核心实施要点", 4, 4),
]

HEADING_RE = re.compile(r"(?m)^[#\s\*]*方向([一二三四五六七八九十])[：:]")
ITEM_RE = re.compile(r"(?m)^\s*(?:[-*•]\s*)?\d+[\.、．)）]\s*\S")


def validate_output(text: str, truncated: bool = False) -> list:
    """返回格式问题列表；列表为空表示通过校验。"""
    issues = []
    if truncated:
        issues.append("输出因长度限制被截断，部分内容缺失")
    if "重要提示" not in text[:500]:
        issues.append("缺少开头的“重要提示”免责声明")

    matches = list(HEADING_RE.finditer(text))
    first_heading_pos = matches[0].start() if matches else len(text)

    # 公司画像
    profile_pos = text.find("公司画像")
    if profile_pos < 0:
        issues.append("缺少“公司画像（简要提炼）”")
    else:
        profile_part = text[profile_pos:first_heading_pos]
        for key in ("所属行业", "核心业务", "当前痛点", "可用资源"):
            if key not in profile_part:
                issues.append(f"公司画像缺少“{key}”")

    # 八套方向
    sections = {}
    for i, m in enumerate(matches):
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        sections.setdefault(m.group(1), text[m.end():end])

    for num, name in EXPECTED_DIRECTIONS:
        label = f"方向{num}（{name}）"
        body = sections.get(num)
        if body is None:
            issues.append(f"缺少{label}")
            continue
        if name not in body[: len(name) + 12]:
            issues.append(f"方向{num}的名称应为“{name}”")

        positions, missing = [], False
        for marker, _, _ in MODULE_RULES:
            pos = body.find(marker)
            if pos < 0:
                issues.append(f"{label}缺少“{marker}”模块")
                missing = True
            positions.append(pos)
        if missing:
            continue
        if positions != sorted(positions):
            issues.append(f"{label}的模块顺序不符合要求")
            continue

        for idx, (marker, low, high) in enumerate(MODULE_RULES):
            if low is None:
                continue
            start = positions[idx]
            end = positions[idx + 1] if idx + 1 < len(positions) else len(body)
            count = len(ITEM_RE.findall(body[start:end]))
            if not (low <= count <= high):
                need = f"{low}条" if low == high else f"{low}-{high}条"
                issues.append(f"{label}的“{marker}”应为{need}，实际{count}条")

        if num == "七":
            for kw in ("专项基金", "同祺"):
                if kw not in body:
                    issues.append(f"{label}未体现“{kw}”")
        if num == "八":
            for kw in ("盲池基金", "LP", "同祺"):
                if kw not in body:
                    issues.append(f"{label}未体现“{kw}”")

    if re.search(r"[{｛][^}｝]{0,300}[}｝]", text):
        issues.append("输出中仍含有未替换的模板占位符 { }")
    return issues


def call_model(messages):
    """调用模型，返回 (文本, 是否因长度被截断)。"""
    response = client.chat.completions.create(
        model=MODEL_NAME,
        messages=messages,
        temperature=0.2,
        max_tokens=MAX_OUTPUT_TOKENS,
        stream=False,
    )
    choice = response.choices[0]
    return (choice.message.content or ""), (choice.finish_reason == "length")


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
    if len(text) > max_chars:
        return text[:max_chars], True
    return text, False


# =============================================================================
# Streamlit 界面
# =============================================================================
st.set_page_config(page_title="转型升级方向", layout="wide")
st.title("🏭 企业转型升级方向生成器")
st.caption("✅ 已按《八套转型升级方向》文档要求对齐结构、深度与语言风格")

DEFAULT_CURRENT_STATUS = """宁波市博华机械配件有限公司成立于2009年，位于浙江省宁波市奉化区，是一家以铝合金铸造及机械精加工为核心业务的传统制造企业。公司具备从模具设计、铝合金重力铸造、翻砂铸造，到CNC加工、数控车削、热处理及检测等较为完整的生产能力，可根据客户图纸提供毛坯件及成品零部件。

公司产品主要包括汽车及摩托车零部件、减震器支架、气动元器件、火车零件、电机及风机配件、石化设备零件、空压机零件和医疗器械零件等，下游应用领域较为分散。公开资料显示，公司产品除供应国内市场外，还出口法国、英国、美国等海外市场，并与安徽中鼎、沃尔沃等客户存在配套合作。

从生产规模看，公司现有员工约90余人，拥有数控加工中心、数控车床、铸造设备、喷砂设备、热处理炉等生产设备，并配置光谱分析仪、三坐标测量仪、X光检测仪等质量检测设备，具备一定的规模化生产及品质管控能力。公开资料披露，公司年产能超过2,000吨，近年销售规模处于亿元左右，其中招聘平台披露其2024年营业收入约5,486万元、年销售额超过8,500万元。

整体来看，博华机械属于典型的"铸造+机械加工"一体化中小型制造企业，业务技术路线成熟，传统制造属性较强，主要竞争力体现在客户定制能力、铸造与机加工一体化能力以及多年积累的生产工艺和客户资源。"""

with st.form(key="company_info_form"):
    company_name = st.text_input(
        "公司名称*",
        value="宁波市博华机械配件有限公司",
        placeholder="例如：山东固丰体育产业有限公司"
    )
    industry = st.text_input(
        "所属行业*",
        value="机械加工",
        placeholder="例如：体育产业"
    )
    current_status = st.text_area(
        "公司当前情况描述*",
        value=DEFAULT_CURRENT_STATUS,
        placeholder="描述公司规模、问题、优势、核心业务、渠道、品牌等...",
        height=220,
    )
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

        extra_text = "\n\n".join(extra_text_blocks)
        extra_text, extra_total_truncated = truncate(extra_text, MAX_TOTAL_EXTRA_CHARS)
        if extra_total_truncated:
            truncate_warnings.append(f"补充文件总内容超出 {MAX_TOTAL_EXTRA_CHARS} 字符上限，已整体截断。")

        if parse_warnings:
            for w in parse_warnings:
                st.warning(f"⚠️ {w}")
        if truncate_warnings:
            for w in truncate_warnings:
                st.info(f"ℹ️ {w}")
        if additional_files:
            st.caption(f"共上传 {len(additional_files)} 个文件，成功解析 {len(extra_text_blocks)} 个。")

        user_context = f"""
公司名称：{company_name}
所属行业：{industry}
当前情况：{current_status}
补充材料：{extra_text if extra_text else "（无）"}
"""

        total_len = len(FULL_SYSTEM_PROMPT) + len(user_context)
        if total_len > MAX_TOTAL_PROMPT_CHARS:
            overflow = total_len - MAX_TOTAL_PROMPT_CHARS
            new_extra_len = max(0, len(extra_text) - overflow)
            extra_text = extra_text[:new_extra_len]
            user_context = f"""
公司名称：{company_name}
所属行业：{industry}
当前情况：{current_status}
补充材料：{extra_text if extra_text else "（因总长度超限已省略）"}
"""
            st.info("ℹ️ 输入总长度接近模型上限，已自动精简补充材料部分以确保正常生成。")

        base_messages = [
            {"role": "system", "content": FULL_SYSTEM_PROMPT},
            {
                "role": "user",
                "content": (
                    "请严格按照【输出要求】，为以下公司输出“免责声明 + 公司画像 + 八套专属转型升级方向”：\n"
                    f"{user_context}"
                ),
            },
        ]

        with st.spinner("正在生成八套专属方向...（约 30-60 秒）"):
            try:
                scheme, was_cut = call_model(base_messages)
                issues = validate_output(scheme, was_cut)

                # 格式不合格时完整重写一次
                if issues:
                    fix_prompt = (
                        "上一版输出未完全符合格式要求，存在以下问题：\n"
                        + "\n".join(f"- {i}" for i in issues)
                        + "\n请严格按照【输出要求】重新输出完整内容（免责声明、公司画像及八套方向），"
                        "不得省略、不得只补充片段；每条控制在80字以内，确保全部输出完整。"
                    )
                    retry_messages = base_messages + [
                        {"role": "assistant", "content": scheme},
                        {"role": "user", "content": fix_prompt},
                    ]
                    scheme2, was_cut2 = call_model(retry_messages)
                    issues2 = validate_output(scheme2, was_cut2)
                    if len(issues2) <= len(issues):
                        scheme, issues = scheme2, issues2

                st.session_state["scheme"] = scheme
                st.session_state["scheme_issues"] = issues
                st.session_state["scheme_company"] = company_name
            except Exception as e:
                st.error(f"AI 调用失败：{str(e)}")

# 结果展示（放在表单外，下载后刷新不丢失）
if st.session_state.get("scheme"):
    st.success("✅ 生成完成！")
    st.warning("**重要提示：本方向仅供参考，不构成任何正式的投资、经营或法律建议。**")

    result_issues = st.session_state.get("scheme_issues", [])
    if result_issues:
        st.warning(
            "⚠️ 输出格式自检发现以下问题，建议重新生成一次：\n\n"
            + "\n".join(f"- {i}" for i in result_issues)
        )
    else:
        st.caption("✅ 已通过格式自检：公司画像 + 八套方向，每套含4个固定模块。")

    st.markdown(st.session_state["scheme"])

    st.download_button(
        label="📥 下载方案（Markdown）",
        data=st.session_state["scheme"],
        file_name=f"{st.session_state.get('scheme_company', '公司')}_转型升级八套方向.md",
        mime="text/markdown",
    )

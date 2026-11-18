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
st.caption("✅ 已按最新规则对齐：方向七=纯财务投资、通过专业投资机构渠道寻找项目并直接投资企业；方向八=科创基金投资")

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
                    "请严格按照【输出要求】，为以下公司输出“免责声明 + 公司画像 + 八套专属转型升级方向”。特别注意方向七必须为纯财务投资，通过专业投资机构渠道寻找项目、由企业直接投资目标企业股权，不得写成战略投资或产业协同：\n"
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
                        "\n特别注意："
                        "\n- 方向七必须是纯财务投资：通过多家专业VC/PE投资机构等渠道寻找项目，但由企业直接投资目标企业股权；投资机构只是项目来源/专业服务渠道，不是基金投资路径；"
                        "\n- 方向七不得出现战略投资、产业协同、供应链绑定、优先采购、联合研发、先投后并、并购储备等战略投资逻辑；不得出现基金、专项基金、LP、GP、基金管理人等内容；"
                        "\n- 方向七必须体现估值、成长性、收入质量、现金流、回报倍数、IRR、退出路径和风险控制；不得写委托宁波股权交易中心或指定任何单一交易平台/投资机构进行标的初筛；"
                        "\n- 方向八写基金投资，但严禁出现任何特定机构名称；"
                        "\n- 全文禁止出现“同祺”以及“获取情报网络”等不当表述。"
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

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
# 系统提示词（已按最新修订版《八套转型升级方向》完整对齐）
# 核心修订：
# 1. 方向七 = 纯直投企业，严禁涉及基金、专项基金、LP/GP、管理人
# 2. 方向八 = 投资科创基金，保持中立，禁止指定/引导任何特定机构
# 3. 全文禁止出现“同祺”及“获取情报网络”等不当表述
# =============================================================================
FULL_SYSTEM_PROMPT = """【角色设定】
你是一位资深的产业战略咨询专家，专注于传统企业转型升级与科创产业融合领域。你擅长将通用的转型升级方法论与企业实际情况深度结合，输出具有针对性、可落地性的战略方向。
请根据以下八套转型升级方向，针对用户提供的公司信息，逐一分析每套方向的适用性并提出具体建议。请严格依据用户提供的信息，不得虚构企业事实、数据、客户关系、行业地位或资源条件。

【重要免责声明】
请在输出最前面原样添加以下提示语句：
重要提示：本方向仅供参考，不构成任何正式的投资、经营或法律建议。实际操作请咨询专业律师、财务顾问及相关行业专家。

【八套转型升级方向库】
方向库中的案例、示例企业仅用于理解方向和参考表达方式，不得照抄；不得将示例信息误写成用户公司的实际情况。输出中不得出现花括号。

方向一：并购科创企业
方向概述：核心在于大型传统企业通过资本运作，直接投资或并购具有创新技术、新兴市场或高成长潜力的科创企业，快速获取前沿技术和创新能力，拓展新业务领域，实现产业升级和多元化发展。
案例分析：
1. 浙江盾安人工环境：作为制冷自控元件及中央空调主机生产商，增资收购大创汽车（专注新能源汽车热管理产品），成功切入新能源汽车产业链，实现战略性业务拓展。
2. 美的集团：斥资292亿元并购德国库卡集团（全球工业机器人巨头），掌握顶级自动化技术，构建“消费电器+暖通空调+机器人自动化”三驾马车格局，从家电巨头转型科技集团。
3. 潍柴动力：收购加拿大巴拉德动力系统公司（氢燃料电池领军企业）19.9%股权成为第一大股东，布局新能源商用车赛道，完成从“烧油”到“烧氢”的技术跨越。
转型优势：快速获取核心技术、拓展新兴市场、优化产业结构、提升品牌价值。
实施要点：明确战略方向、审慎尽职调查、强化整合管理能力、夯实资本运作能力。

方向二：产研合作
方向概述：以科创企业的先进技术和高标准要求反向赋能传统生产企业，通过工艺指导、标准共建、人才交流，推动传统企业技术升级，生产更高附加值产品，提升盈利能力。
案例分析：
1. 上海盈普三维：3D打印设备企业（SLS市占率超60%），与60多家传统零部件、模具、钣金企业合作，带动后者生产工艺和定制化能力升级，匹配3D打印产业高精度需求。
2. 华为云×山东魏桥：引入华为云AI视觉检测系统和工业互联网平台，重构纺纱、织布工艺，断头检测准确率提升至95%以上，生产效率提升30%，从“汗水纺织”转型“智慧纺织”。
3. 大疆创新×深圳模具厂：输出微米级供应链管理标准和CNC自动化改造要求，推动配套模具厂工艺跃升至军工级，成功进入苹果、特斯拉供应链体系。
转型优势：技术溢出效应、提升产品竞争力、优化成本结构、实现利润增长。
实施要点：建立紧密合作机制、明确技术标准导入、加强人才双向交流、绑定长期战略伙伴关系。

方向三：数智化改造
方向概述：围绕生产、销售、供应链、财务等核心环节引入ERP、MES、AI工具等数字化系统，实现从“经验驱动”到“数据驱动”的升级，在不改变主业的前提下提效降本，风险可控、见效较快。
案例分析：
1. 某制造企业：通过引入ERP、生产管理系统、经营数据看板，统一管理订单/库存/采购数据，搭配AI辅助销售跟进，实现库存下降、交付效率提升、管理成本降低，升级为数字化管理企业。
2. 波司登：构建全域数据中台，打通3000+门店和线上商城数据，通过智能补货系统动态调整生产计划，库存周转天数大幅下降，旺季售罄率提升至90%以上，重回国产羽绒服第一梯队。
3. 海螺水泥：推进“智能工厂”建设，引入无人驾驶矿车、AI质控系统，单条生产线减员30%以上，能耗下降5%，年节约成本数亿元，确立行业智能制造标杆。
转型优势：提升经营效率、增强管理透明度、优化成本结构、降低转型风险、形成长期竞争壁垒。
实施要点：明确转型核心目标、从关键痛点切入试点、匹配适配服务商、推动高层牵头组织协同、重视数据治理、分阶段迭代推广、建立长期运营机制。

方向四：渠道合作
方向概述：引导传统经销商从销售低附加值传统产品，转向代理高科技、高毛利的科创企业产品，优化产品结构，提升盈利空间和市场竞争力。
案例分析：
1. 合肥启灏医疗：耳鼻喉科器械企业，对接传统医疗器械经销商，后者摆脱耗材价格战泥潭，获得显著利润增长。
2. 某省级医药代理商：砍掉80%低毛利普药代理权，独家代理癌症早筛基因检测产品，单客利润提升50倍，仅靠少量三甲医院客户维持过亿营收。
3. 传统建材商：与华为全屋智能、欧瑞博合作，利用原有装修工长资源转型智能家居解决方向提供商，客单价从几千元提升至十几万元，实现从“卖砖头”到“卖生活方式”的升级。
转型优势：提升盈利能力、优化产品组合、增强市场地位、降低经营风险。
实施要点：精准选品匹配市场需求、绑定稳定代理合作关系、强化团队专业销售能力、整合现有渠道资源推广。

方向五：精益管理
方向概述：围绕客户价值识别库存积压、等待时间长、返工率高等低效环节，通过流程梳理、标准化作业、现场改善实现从“粗放管理”到“精细化运营”的升级，投入可控、见效较快、风险较低，适用于制造业、供应链等各类传统企业。
案例分析：
1. 某制造企业：导入5S管理、看板管理、质量追溯机制，实现库存资金占用下降、交付稳定性提升、返工率降低，经营从“老板盯人”转向“流程驱动、持续改善”。
2. 丰田汽车：创立“丰田生产方式（TPS）”，通过准时化（JIT）和自动化（Jidoka）消除七大浪费，以极低库存实现极高生产效率，成为全球精益管理教科书。
3. 丹纳赫集团：将精益管理升级为“丹纳赫业务系统（DBS）”，作为投后管理工具植入被并购企业，成功改造数十家传统低效企业，股价40年增长超700倍。
转型优势：降本增效、改善质量与交付能力、提升组织管理能力、投入产出比高。
实施要点：从核心痛点切入优先改善、建立标准化管理体系、形成持续改善机制、培育全员参与文化。

方向六：全球化出海布局
方向概述：依托“一带一路”、RCEP机遇，通过海外建厂、海外销售、海外融资、海外身份配置等方式，构建“国内研发与核心部件制造+海外本地化生产销售+境外资金与身份支持”协同体系，规避贸易壁垒，降低生产成本，完成从出口型厂商向跨国经营企业的转型。
案例分析：
1. 玲珑轮胎：在泰国建设生产基地，依托本地橡胶资源与东盟零关税政策压缩成本，国内总部负责研发与核心部件生产，海外基地量产供货全球市场，双产能布局分散风险，实现产业出海升级。
2. 海尔智家：在美国、意大利、印度等地建设海外生产基地，构建“中国研发+本地生产+全球销售”体系，规避欧美贸易壁垒，降低物流成本，海外业务收入占比持续提升，转型为跨国经营企业。
3. 福耀玻璃：在美国俄亥俄州建设全球最大汽车玻璃单体工厂，总投资超6亿美元，国内总部负责玻璃配方研发与核心设备制造，美国工厂本地化生产与销售，供应通用、福特等车企，规避美国反倾销税，转型为跨国制造企业。
转型优势：规避贸易壁垒、优化生产成本、深耕海外市场、分散经营风险、提升资金配置效率、完善企业家身份配置、享受出海政策。
实施要点：规划出海区域与产能、灵活选择出海建厂模式、做好出海属地合规尽调、设计合规资金出海路径、搭建海外主体与账户体系、分步搭建海外本地供应链、搭建出海本土化管理体系、规划企业主与核心团队身份出海、分阶段落地长效出海运营。

方向七：投资科创企业
方向概述：鼓励资方直接投资具有创新技术、高成长潜力的科创企业（含Pre-IPO阶段的明星标的）。根据标的与资方主业的关系，分为“产业相关”和“产业不相关”两种情形，投资目的、标的筛选与考核口径各不相同，须先判定再设计方案。
情形一：产业相关型——投资与自身主业处于同一或相邻产业链（上下游、配套、渠道）的科创企业。目的是了解上下游技术与市场、建立配套关系（优先采购/供应、联合研发、渠道绑定），或以参股方式为后续并购（见方向一）做准备；收益来源为产业协同价值叠加股权增值，考核重点是配套关系落地、优先供应权及并购转换条款。
情形二：产业不相关型——投资与自身主业无关的优质科创赛道，属纯财务投资，不要求产业协同、不介入经营，追求高收益，实现盈利模式由主业利润向股权收益转型；重点选择高成长、高确定性（如Pre-IPO阶段）标的，考核重点是回报倍数与IRR、估值合理性、风险敞口控制及退出路径。
案例分析：
1. 合肥启灏医疗（产业相关·下游渠道配套）：代理启灏医疗（耳鼻喉科器械）产品的传统经销商，联合出资定向投资启灏医疗，经销商从渠道商升级为产业投资人，在渠道利润之外获得股权增值收益，同时将主要经销商转化为股东，稳固市场渠道、降低营销成本。
2. 智星空间（产业相关·产业链生态布局）：商业卫星产业链核心环节的传统产业资本，战略投资专注商业SAR卫星制造与运营的智星空间，以小资金撬动商业航天大产业，形成“研发—制造—运营—应用”的产业生态，避免独立布局所需的高昂研发成本。
3. 元气森林（股权增值收益参考）：某传统消费企业集团联合头部创投机构定向投资元气森林（无糖饮料科创品牌），除销售高毛利产品外，分享企业估值增长的股权增值收益，转型为“销售+投资”双轮驱动。
转型优势：产业相关型——了解上下游技术与市场动向、锁定配套与优先供应权、借股东身份绑定渠道，并可“先投后并”，为后续并购积累信息、降低整合风险；产业不相关型——不依赖主业协同，直接分享高成长科创赛道的股权增值，以高收益驱动盈利模式向股权收益转型。
实施要点：先判定投资属性并采用不同考核口径（相关型看协同价值与并购转换条款，不相关型看回报倍数、IRR与退出路径）；全面评估标的投资价值，避免估值泡沫；明确股东权利义务与退出机制；不相关型严控投资规模并不干预标的经营，相关型平衡主业与投资双重角色。

方向八：投资科创基金
方向概述：以民营企业产业转型升级和资本增值为目标，通过设立或参与科创专项基金、母子基金、产业基金及并购基金等方式，引入专业投资机构，实现产业资源整合、科技项目投资和资本收益提升。其中，鼓励资方以LP身份参与“盲池基金”（出资时不锁定具体标的，由管理人在约定领域内择优投资多个科创项目）。
核心模式：
模式一：民营企业专项基金——以民营企业为核心发起方，联合专业投资机构设立科创专项基金，围绕企业发展战略定向投资科技型企业，实现产业协同与股权增值。
模式二：母基金与子基金——通过设立母基金、参与专业子基金，布局不同科技产业赛道，借助专业基金管理人的投资能力，分散投资风险，扩大投资覆盖面。
模式三：产业投资基金——围绕企业现有产业链上下游，投资具有技术优势、市场潜力和协同价值的科创企业，为业务拓展、供应链整合及未来并购储备资源。
模式四：并购投资基金——通过设立或参与并购基金，投资具有整合价值的科技型企业，推动产业并购、业务整合及资本退出。
投资策略：
策略一：产业协同型投资——围绕企业主营业务及上下游产业链，投资科创企业，服务于技术升级、供应链整合、业务拓展及未来并购。
策略二：财务收益型投资——不局限于现有产业领域，通过专业投资机构筛选高成长性科技企业，以股权投资、资本增值和市场化退出为主要目标。
案例分析：
1. 某转型升级基金（专业机构管理·盲池基金）：基金规模1亿元，存续期3+2年（投资期3年、退出期2年），聚焦智能制造与医疗健康领域的中后期项目，管理费2%/年，管理人提取20%超额收益，资方以LP身份分散参与多个科创项目。
2. 水木资本×BABC市场化母基金（母基金）：总规模20亿元，采取“基金投资+直接投资+上市公司并购投资”复合策略，筛选并投资优秀基金，兼顾投资广度与产业深度，为被投企业提供退出通道，实现资本与产业深度协同。
3. 信宸资本×嘉信立恒（并购投资基金）：避开拥挤的住宅物业赛道，收购7家区域性设施管理企业整合为“嘉信立恒”，营收从3.6亿元增至超30亿元，通过上市公司并购实现资本回流。
转型优势：发挥民营企业的市场化决策和产业资源优势，引入专业投资机构的项目筛选、投资管理及退出能力，实现产业转型、科技赋能与资本收益的多重价值。
实施要点：明确民营企业基金发起人及出资结构；优先联合具有成熟投资经验的专业市场化基金管理机构；设计专项基金及母子基金架构；建立投资决策、风险控制、投后管理和退出机制，兼顾产业协同与财务回报。

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

三、方向七和方向八的特别要求（必须严格遵守）

方向七：投资科创企业
- 方向七仅讨论资方直接投资科创企业，体现“直投企业”的投资路径。
- 可根据企业情况，分别分析与主业相关的科创企业投资机会，以及与主业关联较弱或不相关的科创企业投资机会。
- 【严禁】涉及基金设立、专项基金、子基金、母基金、LP、GP或基金管理人，不得将直投企业与基金投资混为一谈。
- 应结合公司的产业基础、技术需求、投资能力和风险承受能力提出建议；相关信息未提供时，标注“（待补充）”或给出审慎的原则性建议。

方向八：投资科创基金
- 方向八专门讨论通过投资科创基金参与科创投资，可分析专项基金、子基金、母基金等方式，也可说明以LP身份参与基金的可能路径。
- 基金投资建议应围绕规范投资、技术协同、产业升级及 ex 长期价值展开。
- 可以结合基金规模、存续期限、管理费、收益分配、投资范围、风险控制和投后协同等要素提出建议；若用户未提供相关信息，不得虚构具体数值或已达成的安排。
- 【严禁】使用“获取情报网络”“获取商业情报”“套取信息”或其他暗示非法获取、刺探、交换商业秘密或非公开信息的表述。不得将基金投资描述为获取竞争对手信息的手段。
- 【严禁】指定、推荐或引导选择任何特定投资机构或基金管理人。确有机构选择需求时，仅提出尽职调查和遴选标准（如管理团队经验、投资策略、合规风控、项目来源、投后服务和历史业绩），并提示由用户结合实际情况自行判断或另行沟通。

四、内容深度与风格要求
1. 必须深度融合该公司的产业特点、核心资源、渠道网络、区域地位等实际情况；每条落地方向都要具体到业务板块、产品品类、客户群体，禁止泛泛而谈。
2. 方向库文档中的示例企业仅作为格式和深度的参考，禁止照抄，输出中不得出现花括号，必须基于用户公司重新撰写。
3. 涉及金额、比例等量化目标时，应基于用户提供的数据合理测算，并使用“约”“预计”“目标”等措辞，不得写成既成事实，不得编造用户未提供的财务数据。
4. 语言专业严谨，可使用“第二曲线”“生态协同”“长期资本”“精益红利”等产业术语，但须与公司实际情况相符，不要为使用术语而生硬堆砌。
5. 涉及品牌势能、区域龙头地位、技术优势、客户资源等判断，必须有用户提供的信息支撑；如无依据，不得作此类描述。
6. 语言专业、严谨、清晰，避免夸大收益、保证投资回报或作出确定性承诺。
7. 每条落地方向、优势和实施要点尽量控制在80字以内；确保公司画像与八套方向全部完整输出，不得因篇幅而省略任何方向或模块。
8. 最终输出中不得出现任何特定投资机构或基金管理人的名称，不得引导选择任何特定投资机构或基金管理人。
"""

# =============================================================================
# 输出格式校验（已同步最新规则）
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

# 禁止出现的词汇（方向七严禁基金相关；全文严禁特定机构与情报类表述）
FORBIDDEN_IN_DIR7 = ["专项基金", "子基金", "母基金", "盲池基金", "基金管理人", "担任管理人", "LP身份", "以LP"]
FORBIDDEN_GLOBAL = ["同祺", "获取情报网络", "获取商业情报", "套取信息"]


def validate_output(text: str, truncated: bool = False) -> list:
    """返回格式问题列表；列表为空表示通过校验。"""
    issues = []
    if truncated:
        issues.append("输出因长度限制被截断，部分内容缺失")
    if "重要提示" not in text[:500]:
        issues.append("缺少开头的“重要提示”免责声明")

    # 全局禁止词检查
    for kw in FORBIDDEN_GLOBAL:
        if kw in text:
            issues.append(f"输出中出现禁止表述「{kw}」")

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

        # 方向七专项检查：必须是直投企业，严禁基金相关表述
        if num == "七":
            for kw in FORBIDDEN_IN_DIR7:
                if kw in body:
                    issues.append(f"{label}出现禁止的基金相关表述「{kw}」（方向七仅允许直投企业）")
            # 鼓励出现“直投”或“直接投资”等表述（软提示，不强制）
            if "直投" not in body and "直接投资" not in body and "参股" not in body and "股权" not in body:
                issues.append(f"{label}建议明确体现“直投企业/参股/股权投资”路径")

        # 方向八专项检查：应涉及基金，但不得出现特定机构
        if num == "八":
            fund_keywords = ["基金", "LP", "盲池", "专项", "母基金", "子基金"]
            if not any(kw in body for kw in fund_keywords):
                issues.append(f"{label}未体现基金投资相关内容")

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
st.caption("✅ 已按最新修订版《八套转型升级方向》对齐：方向七=直投企业，方向八=中性基金投资，无特定机构引导")

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
                        "\n特别注意："
                        "\n- 方向七只能写直投企业，严禁出现基金、专项基金、LP、管理人等表述；"
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
        st.caption("✅ 已通过格式自检：公司画像 + 八套方向，每套含4个固定模块；方向七为直投企业，方向八为中性基金投资。")

    st.markdown(st.session_state["scheme"])

    st.download_button(
        label="📥 下载方案（Markdown）",
        data=st.session_state["scheme"],
        file_name=f"{st.session_state.get('scheme_company', '公司')}_转型升级八套方向.md",
        mime="text/markdown",
    )

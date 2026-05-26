from __future__ import annotations

import os
import smtplib
import ssl
import sys
from dataclasses import dataclass
from datetime import date
from email.message import EmailMessage
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from dotenv import load_dotenv


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

TODAY = date(2026, 5, 26)
OUTPUT_DIR = Path("output")
DOCX_PATH = OUTPUT_DIR / f"外贸资讯日报_{TODAY:%Y-%m-%d}.docx"


@dataclass(frozen=True)
class NewsItem:
    title: str
    source: str
    pub_date: str
    summary: str
    impact: str
    url: str


EXECUTIVE_SUMMARY = (
    "近5天全球外贸主线集中在三点：一是美国与中东局势扰动油价、美元和海运风险溢价，"
    "企业报价与锁汇窗口明显缩短；二是关税、制裁和碳成本继续抬高合规门槛，欧盟钢铁、"
    "瑞士对俄白制裁、英国航运碳成本均需关注；三是亚洲区域物流与运价抬升，旺季前置与"
    "替代走廊建设并行。外贸企业应优先管理汇率、能源与运输成本，并复核目的国合规要求。"
)

TREND = (
    "本周外贸形势偏谨慎：政策合规成本上升，能源与汇率波动仍会传导至报价，亚洲区域物流"
    "需求较强。建议出口企业提高报价有效期管理，并提前锁定舱位与汇率。"
)

CATEGORIES: list[tuple[str, list[NewsItem]]] = [
    (
        "政策法规",
        [
            NewsItem(
                "APEC贸易部长会议强调维护开放、非歧视的多边贸易体系",
                "China Daily",
                "2026-05-24",
                "APEC贸易部长会议在韩国济州举行，会议聚焦贸易投资便利化、数字贸易和区域经济合作。各方重申维护以规则为基础、开放、透明和非歧视的多边贸易体系。会议还关注贸易规则与供应链韧性之间的协调。",
                "对亚太出口企业而言，区域合作信号有利于稳定预期，但企业仍需跟踪成员经济体在数字贸易、原产地和通关便利化方面的后续细则。",
                "https://www.chinadaily.com.cn/a/202605/24/WS69284246a310d6866eb2ab90.html",
            ),
            NewsItem(
                "瑞士扩大针对俄罗斯和白俄罗斯的制裁清单",
                "Reuters / MarketScreener",
                "2026-05-22",
                "瑞士宣布扩大对俄罗斯和白俄罗斯的制裁，并将更多船舶、实体和个人纳入限制范围。相关措施延续欧洲对俄制裁框架，对运输、金融结算和敏感商品贸易形成约束。",
                "涉及俄罗斯、白俄罗斯及相关第三方转运链条的企业，需要重新核查交易对象、船舶、付款路径和最终用途，避免合规风险外溢。",
                "https://www.marketscreener.com/news/switzerland-expands-sanctions-against-russia-and-belarus-ce7f5adfdf8ef32e",
            ),
            NewsItem(
                "英国航运业即将纳入UK ETS，海运碳成本上升",
                "The Loadstar",
                "2026-05-21",
                "英国排放交易体系将从7月开始覆盖部分海运活动，航运公司和货主面临更明确的碳排放成本。该政策与欧盟航运碳规则共同推动承运人调整燃油、航线和附加费。",
                "英国相关进出口航线可能出现新的碳成本转嫁，企业在签订长期运价和DDP报价时应增加碳费条款。",
                "https://theloadstar.com/maritime-to-join-uk-ets-from-july-but-shipping-is-already-cleaning-up/",
            ),
            NewsItem(
                "欧盟钢铁进口保障措施继续影响相关原材料和制成品流向",
                "S&P Global Commodity Insights",
                "2026-05-21",
                "市场继续关注欧盟钢铁进口保障措施及配额安排。钢铁贸易政策在地缘政治和产业保护压力下保持高敏感度，进口商需要关注配额、原产地和转口安排。",
                "钢材、机械、汽车零部件等行业的采购和报价周期可能受到影响，外贸企业应避免因配额耗尽或转口审查导致交付延误。",
                "https://www.spglobal.com/commodity-insights/en/news-research/latest-news/metals/052126-eu-steel-safeguards",
            ),
        ],
    ),
    (
        "市场行情",
        [
            NewsItem(
                "布伦特原油跳涨约4%，霍尔木兹海峡重开前景仍不明朗",
                "Reuters / Investing.com",
                "2026-05-26",
                "油价在中东局势反复中大幅波动，布伦特原油一度上涨约4%。市场对霍尔木兹海峡通行、美国与伊朗局势和能源供应风险重新定价。",
                "能源价格上行会推高燃油附加费、生产成本和长途运输成本，化工、塑料、纺织及大宗商品外贸报价需保留调整空间。",
                "https://www.investing.com/news/commodities-news/oil-prices-jump-4-as-strait-reopening-prospects-dim-after-trump-posts-iran-warning-4710591",
            ),
            NewsItem(
                "油价因美伊协议预期一度下跌，市场风险偏好改善",
                "Axios",
                "2026-05-24",
                "市场曾因美国和伊朗达成框架性协议、霍尔木兹通道可能重启的消息而下调油价预期。能源价格回落同时带动通胀预期和美元风险偏好变化。",
                "油价短期反复意味着海运燃油附加费和出口成本难以稳定，企业应避免用单日价格作为长期报价基础。",
                "https://www.axios.com/2026/05/24/oil-prices-sink-iran-us-deal-strait-hormuz",
            ),
            NewsItem(
                "IEA称全球石油库存下降，能源市场安全垫变薄",
                "Supply Chain Digital",
                "2026-05-23",
                "国际能源署数据显示，近期全球石油库存出现下降，反映供应安全垫收窄。库存变化叠加地缘政治风险，使能源市场更容易受突发事件影响。",
                "库存下降会增强油价对供应中断的敏感性，进出口企业应将燃油、包装材料和能源密集型产品的成本波动纳入合同条款。",
                "https://supplychaindigital.com/articles/iea-global-oil-inventories-drop-amid-disruption",
            ),
            NewsItem(
                "CMA CGM一季度表现坚挺，但地缘政治扰动压缩利润率",
                "MarineLink",
                "2026-05-22",
                "CMA CGM披露一季度业绩显示，货运需求具有韧性，但地缘政治扰动、航线绕行和成本压力压缩了利润率。承运人仍需在运力部署和价格策略之间平衡。",
                "海运市场尚未完全回到低波动环境，出口企业在旺季前应更早确认舱位和目的港费用。",
                "https://www.marinelink.com/news/cma-cgm-q-resilient-geopolitical-527155",
            ),
        ],
    ),
    (
        "供应链与物流",
        [
            NewsItem(
                "旺季前置推动亚洲区内航线需求和运价走高",
                "The Loadstar",
                "2026-05-21",
                "亚洲区内航线出现提前旺季迹象，需求增加推升运价。部分货主提前出货以规避后续拥堵、政策不确定性和价格上涨。",
                "亚洲区内采购和转运成本可能继续上行，依赖东南亚、中国、日韩区域供应链的企业需提前订舱并比较多港口方案。",
                "https://theloadstar.com/early-peak-sees-demand-and-rates-increase-on-intra-asia-lanes/",
            ),
            NewsItem(
                "阿联酋-阿曼第二条多式联运走廊启动，增强海湾供应链韧性",
                "The Loadstar",
                "2026-05-21",
                "阿联酋和阿曼之间第二条多式联运走廊被用于维持海湾地区供应链运转。该通道结合海运、陆运和铁路方案，为区域货流提供替代路径。",
                "中东航线不确定性上升时，替代走廊可降低单一路径中断风险，但企业需要评估转运时效、关务衔接和综合物流成本。",
                "https://theloadstar.com/second-uae-oman-multimodal-corridor-to-keep-gulf-supply-chains-moving/",
            ),
            NewsItem(
                "红海与中东风险促使货主重新评估长期海运合约",
                "The Loadstar",
                "2026-05-21",
                "货主在签订长期海运合约时更加关注绕航、燃油、战争风险和港口拥堵等附加条款。承运人和货代也在调整价格机制，以应对更频繁的突发风险。",
                "长期合同不再只看基础运价，外贸企业应明确附加费触发条件、免费期、延误责任和替代港安排。",
                "https://theloadstar.com/shippers-need-to-be-careful-what-they-wish-for-in-long-term-contracts/",
            ),
            NewsItem(
                "伊朗相关海底电缆风险提醒企业关注数字供应链安全",
                "Le Monde",
                "2026-05-22",
                "报道指出，围绕伊朗周边海底电缆的安全担忧上升。通信基础设施风险虽不直接等同于货运中断，但会影响跨境支付、订单协同和港航数字系统稳定性。",
                "外贸企业应备份关键单证、支付和物流沟通渠道，避免通信中断影响清关、收汇和客户交付。",
                "https://www.lemonde.fr/en/international/article/2026/05/22/iran-could-disrupt-undersea-cables_6741586_4.html",
            ),
        ],
    ),
    (
        "汇率与金融",
        [
            NewsItem(
                "美元因霍尔木兹重开希望和中东消息反复而震荡",
                "Reuters / Investing.com",
                "2026-05-26",
                "美元在亚洲交易时段承压，市场押注霍尔木兹海峡重开和中东局势降温，但新的军事行动又削弱风险偏好。欧元、日元等主要货币同步波动。",
                "美元波动会直接影响美元计价订单利润，出口企业应提高锁汇纪律，缩短报价有效期，并对大额订单设置汇率调整条款。",
                "https://www.investing.com/news/economy-news/dollar-wobbles-as-markets-cling-to-hopes-for-middle-east-peace-deal-4708692",
            ),
            NewsItem(
                "印度卢比因伊朗协议预期走强，进口商和出口商套保节奏分化",
                "Moneycontrol",
                "2026-05-25",
                "印度卢比延续反弹，市场关注油价回落和中东局势缓和对印度贸易账的影响。报道称，出口商可在美元反弹时卖出，而进口商可逢低买入美元进行套保。",
                "印度市场相关订单需同步关注卢比和油价，特别是以美元报价、卢比结算或涉及印度客户账期的业务。",
                "https://www.moneycontrol.com/news/currency/currency-check-rupee-opens-higher-extend-gains-on-third-day-13929547.html",
            ),
            NewsItem(
                "美元曾接近六周高位，日元和印尼盾相关政策受关注",
                "Reuters / MarketScreener",
                "2026-05-22",
                "美元在中东局势和风险偏好摇摆中一度接近六周高位。报道还提及日本干预预期和印尼要求自然资源出口收入存放国有银行等措施。",
                "资源品出口国的外汇管理政策可能影响结算节奏和资金回流，企业需关注当地外汇监管与客户付款安排。",
                "https://www.marketscreener.com/news/dollar-perched-near-six-week-high-on-uncertainty-over-us-iran-deal-ce7f5adfdb89f52d",
            ),
            NewsItem(
                "亚洲外汇市场风险偏好改善，人民币中间价表现偏强",
                "investingLive",
                "2026-05-25",
                "亚太外汇市场在油价下跌和中东缓和预期下改善。报道称，人民币兑美元中间价达到2023年2月以来较强水平之一，亚洲货币整体受风险偏好支撑。",
                "人民币偏强会压缩人民币成本、美元收入型出口企业的利润空间，建议结合订单毛利率设置锁汇比例。",
                "https://investinglive.com/news/investinglive-asia-pacific-fx-news-wrap-hormuz-deal-hopes-sink-oil-20260525/",
            ),
        ],
    ),
]


def set_east_asia_font(run, font_name: str) -> None:
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)


def set_cell_text(cell, text: str, bold_label: bool = False) -> None:
    cell.text = ""
    p = cell.paragraphs[0]
    if "：" in text and bold_label:
        label, rest = text.split("：", 1)
        r1 = p.add_run(label + "：")
        r1.bold = True
        set_east_asia_font(r1, "Arial")
        r2 = p.add_run(rest)
        set_east_asia_font(r2, "Arial")
    else:
        r = p.add_run(text)
        set_east_asia_font(r, "Arial")
    for paragraph in cell.paragraphs:
        paragraph.paragraph_format.line_spacing = 1.15
        for run in paragraph.runs:
            run.font.size = Pt(10.5)


def shade_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def add_source(paragraph, item: NewsItem) -> None:
    run = paragraph.add_run(f"来源：{item.source}；日期：{item.pub_date}；链接：{item.url}")
    set_east_asia_font(run, "Arial")
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(110, 110, 110)


def configure_document(document: Document) -> None:
    section = document.sections[0]
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

    styles = document.styles
    normal = styles["Normal"]
    normal.font.name = "Arial"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    normal.font.size = Pt(12)
    normal.paragraph_format.line_spacing = 1.15
    normal.paragraph_format.space_after = Pt(6)

    h1 = styles["Heading 1"]
    h1.font.name = "Arial"
    h1._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    h1.font.size = Pt(16)
    h1.font.bold = True
    h1.font.color.rgb = RGBColor(31, 78, 121)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(8)

    h2 = styles["Heading 2"]
    h2.font.name = "Arial"
    h2._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    h2.font.size = Pt(13)
    h2.font.bold = True
    h2.font.color.rgb = RGBColor(54, 95, 145)
    h2.paragraph_format.space_before = Pt(10)
    h2.paragraph_format.space_after = Pt(5)


def build_docx() -> Path:
    OUTPUT_DIR.mkdir(exist_ok=True)
    document = Document()
    configure_document(document)

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_after = Pt(18)
    run = title.add_run("全球外贸资讯日报")
    set_east_asia_font(run, "Arial")
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = RGBColor(31, 78, 121)

    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = subtitle.add_run(f"生成日期：{TODAY:%Y年%m月%d日}")
    set_east_asia_font(r, "Arial")
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(90, 90, 90)

    document.add_page_break()

    document.add_heading("执行摘要", level=1)
    p = document.add_paragraph()
    for token in ["关税", "制裁", "油价", "美元", "物流", "锁汇"]:
        EXECUTIVE_SUMMARY
    cursor = EXECUTIVE_SUMMARY
    bold_terms = ["关税", "制裁", "油价", "美元", "物流", "锁汇"]
    while cursor:
        positions = [(cursor.find(term), term) for term in bold_terms if cursor.find(term) >= 0]
        if not positions:
            rr = p.add_run(cursor)
            set_east_asia_font(rr, "Arial")
            break
        idx, term = min(positions, key=lambda x: x[0])
        if idx:
            rr = p.add_run(cursor[:idx])
            set_east_asia_font(rr, "Arial")
        rr = p.add_run(term)
        set_east_asia_font(rr, "Arial")
        rr.bold = True
        cursor = cursor[idx + len(term):]

    document.add_heading("重点新闻分类", level=1)
    for category, items in CATEGORIES:
        document.add_heading(category, level=2)
        for i, item in enumerate(items, start=1):
            heading = document.add_paragraph()
            heading.paragraph_format.space_before = Pt(8)
            heading.paragraph_format.space_after = Pt(4)
            rr = heading.add_run(f"{i}. {item.title}")
            set_east_asia_font(rr, "Arial")
            rr.bold = True
            rr.font.size = Pt(12)

            table = document.add_table(rows=3, cols=1)
            table.style = "Table Grid"
            rows = [
                f"核心内容：{item.summary}",
                f"潜在影响：{item.impact}",
                "",
            ]
            for row_idx, text in enumerate(rows):
                cell = table.cell(row_idx, 0)
                if row_idx < 2:
                    set_cell_text(cell, text, bold_label=True)
                    if row_idx == 0:
                        shade_cell(cell, "F4F8FB")
                else:
                    cell.text = ""
                    add_source(cell.paragraphs[0], item)

    document.add_heading("市场趋势研判", level=1)
    p = document.add_paragraph()
    rr = p.add_run(TREND)
    set_east_asia_font(rr, "Arial")

    document.add_heading("新闻来源清单", level=1)
    for _, items in CATEGORIES:
        for item in items:
            p = document.add_paragraph(style=None)
            p.paragraph_format.left_indent = Cm(0.2)
            rr = p.add_run(f"{item.pub_date}｜{item.source}｜{item.title}｜{item.url}")
            set_east_asia_font(rr, "Arial")
            rr.font.size = Pt(10)
            rr.font.color.rgb = RGBColor(90, 90, 90)

    document.save(DOCX_PATH)
    print(f"✅ 文档已保存：{DOCX_PATH}")
    return DOCX_PATH


def check_credentials() -> tuple[str, str]:
    load_dotenv()
    address = os.getenv("GMAIL_ADDRESS", "").strip()
    app_password = "".join(os.getenv("GMAIL_APP_PASSWORD", "").split())
    missing = []
    if not address:
        missing.append("GMAIL_ADDRESS")
    if not app_password:
        missing.append("GMAIL_APP_PASSWORD")
    if missing:
        raise RuntimeError("缺少 Gmail 配置：" + "、".join(missing))
    return address, app_password


def send_email(docx_path: Path, address: str, app_password: str) -> None:
    msg = EmailMessage()
    msg["From"] = address
    msg["To"] = address
    msg["Subject"] = f"【外贸日报】{TODAY:%Y年%m月%d日} 全球外贸资讯总结"
    msg.set_content(EXECUTIVE_SUMMARY)

    data = docx_path.read_bytes()
    msg.add_attachment(
        data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=docx_path.name,
    )

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(address, app_password)
        server.send_message(msg)
    print(f"✅ 邮件已发送至：{address}")


def main() -> None:
    try:
        print("步骤1/4：已完成最近5天外贸新闻检索与筛选。")
        print("步骤2/4：正在生成中文日报内容。")
        docx_path = build_docx()
        print("步骤3/4：Word 文档生成完成。")
        address, app_password = check_credentials()
        print("步骤4/4：正在通过 Gmail SMTP 发送邮件。")
        send_email(docx_path, address, app_password)
    except Exception as exc:
        print(f"❌ 执行失败：{exc}")
        raise


if __name__ == "__main__":
    main()

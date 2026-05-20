from invoice_app.parser import parse_invoice_date, parse_item_names, parse_party_details


def test_parse_invoice_date_accepts_spaces_inside_chinese_date():
    text = "开票日期：\n2026年 04月 27日"

    assert parse_invoice_date(text) == "2026年04月27日"


def test_parse_party_details_splits_individual_buyer_and_company_seller():
    text = """
电子发票（普通发票） 发票号码：
开票日期：
购
买
方
信
息
统一社会信用代码/纳税人识别号：
销
售
方
信
息
统一社会信用代码/纳税人识别号：
名称： 名称：
开票人：
26442000001633847956
2026年02月10日
个人 北京五湖瑞顺科技有限公司广东分公司
91440101MA9Y9RFM26
¥ 4746.57 ¥ 617.06
"""

    buyer_name, buyer_tax_code, seller_name, seller_tax_code = parse_party_details(
        text,
        "26442000001633847956",
    )

    assert buyer_name == "个人"
    assert buyer_tax_code is None
    assert seller_name == "北京五湖瑞顺科技有限公司广东分公司"
    assert seller_tax_code == "91440101MA9Y9RFM26"



def test_parse_item_names_supports_merged_tax_rate_and_amount_token():
    text = """
项目名称 规格型号 单 位 数 量 单 价 金 额 税率/征收率 税 额
*现代服务*技术服务费 1%18629.70 186.30
销方开户银行:赣州银行股份有限公司滨江支行; 银行账号:2841000103080013836
合 计
"""

    assert parse_item_names(text) == ["*现代服务*技术服务费"]

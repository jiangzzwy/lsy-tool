# Configuration for the Excel-to-Word generation tool

# Source Excel file path
SOURCE_EXCEL = "demands/工单列表_20260609.xlsx"

# Output directory
OUTPUT_DIR = "output"

# Template files
TEMPLATES = {
    "外卖": "demands/（上海京东到家友恒 模版）案件线索移送函J20265002926.docx",
    "三方公司": "demands/（三方公司）案件线索移送函J20265159185.docx",
    "三方个人": "demands/（三方个人）案件线索移送函J20261179295.docx",
    "自营开票": "demands/（自营开票）案件线索移送函J20262053109.docx",
}

# Ledger template
LEDGER_TEMPLATE = "demands/台账(1）.xlsx"

# --- API Configuration for enterprise registration authority lookup ---
# Supported providers: "tianyancha_web", "tianyancha", "qichacha", "mock"
#   tianyancha_web: scrape tianyancha.com (no API key needed, rate-limited)
#   tianyancha: official Tianyancha API (requires token)
#   qichacha: official Qichacha API (requires key+secret)
#   mock: address heuristic only (no network)
API_PROVIDER = "tianyancha_web"

# Tianyancha API
TIANYANCHA_API_TOKEN = ""
TIANYANCHA_API_URL = "https://openapi.tianyancha.com/services/v3/open/baseinfo"

# Qichacha API
QICHACHA_API_KEY = ""
QICHACHA_SECRET_KEY = ""
QICHACHA_API_URL = "https://api.qichacha.com/Enterprise/GetEnterpriseInfo"

# Column mapping (1-based index)
COL = {
    "A": 1,   # 序号
    "B": 2,   # 任务单号
    "C": 3,   # 登记单号
    "D": 4,   # 登记时间
    "E": 5,   # 工单来源
    "F": 6,   # 姓名
    "G": 7,   # 手机号码
    "H": 8,   # 三方/自营
    "I": 9,   # 订单号
    "J": 10,  # 商品编号
    "K": 11,  # 商品名称
    "L": 12,  # 店铺名称
    "M": 13,  # 商品类别
    "N": 14,  # 企业名称
    "O": 15,  # 企业地址
    "P": 16,  # 统一社会信用代码
}

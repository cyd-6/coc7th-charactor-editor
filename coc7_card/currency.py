"""Annual USD-based quotes imported from the user's historical workbook."""
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext
import json
import hashlib
from pathlib import Path

DATA_PATH = Path(__file__).resolve().parents[1] / "assets/rates/annual.json"
YEAR_START, YEAR_END = 1920, 2026


def load_exchange_data(path=DATA_PATH):
    raw = Path(path).read_bytes()
    data = json.loads(raw)
    data["data_sha256"] = hashlib.sha256(raw).hexdigest()
    if data["schema_version"] != 1:
        raise ValueError("不支持的年度汇率数据版本。")
    return data


def currency_options(quotes, year, *, available=True):
    if isinstance(year, bool) or not isinstance(year, int) or not YEAR_START <= year <= YEAR_END:
        raise ValueError("汇率年份仅支持 1920—2026 年，请输入范围内的整数年份。")
    return [item for item in quotes if item["year"] == year and item["available"] == available]


def legacy_currency_code(code, year):
    # Only aliases identifying the same historical unit are safe to migrate.
    if code == "FRF" and 1920 <= year <= 1959:
        return "FRF_OLD"
    if code == "GRR" and 1925 <= year <= 1947:
        return "RM"
    return code


def parse_usd(value, fallback):
    if value is None or str(value).strip() == "":
        return Decimal(str(fallback))
    cleaned = str(value).strip().replace("$", "").replace(",", "").replace("USD", "").strip()
    try:
        amount = Decimal(cleaned)
    except InvalidOperation as exc:
        raise ValueError("美元金额请输入数字，可包含千位逗号；说明文字请填写在资产说明中。") from exc
    if not amount.is_finite() or amount < 0 or amount > Decimal("1e15"):
        raise ValueError("美元金额须为 0 到 1,000,000,000,000,000 之间的有限数值。")
    return amount


def format_amount(amount, rate):
    with localcontext() as context:
        context.prec = 60
        return f"{(Decimal(str(amount)) * Decimal(str(rate))).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP):,.2f}"


def usd_reference(credit, year):
    credit = max(0, min(99, credit))
    if credit == 0:
        values = (0.5, 0.5, 0)
    elif credit <= 9:
        values = (2, credit, credit * 10)
    elif credit <= 49:
        values = (10, credit * 2, credit * 50)
    elif credit <= 89:
        values = (50, credit * 5, credit * 500)
    elif credit <= 98:
        values = (250, credit * 20, credit * 2000)
    else:
        values = (5000, 50000, 5000000)
    return tuple(value * (20 if year == 2026 else 1) for value in values)

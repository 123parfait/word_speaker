import re


_SMALL_NUMBERS = {
    0: "zero",
    1: "one",
    2: "two",
    3: "three",
    4: "four",
    5: "five",
    6: "six",
    7: "seven",
    8: "eight",
    9: "nine",
    10: "ten",
    11: "eleven",
    12: "twelve",
    13: "thirteen",
    14: "fourteen",
    15: "fifteen",
    16: "sixteen",
    17: "seventeen",
    18: "eighteen",
    19: "nineteen",
}

_TENS = {
    20: "twenty",
    30: "thirty",
    40: "forty",
    50: "fifty",
    60: "sixty",
    70: "seventy",
    80: "eighty",
    90: "ninety",
}

_SCALES = [
    (1_000_000_000, "billion"),
    (1_000_000, "million"),
    (1_000, "thousand"),
]

_UNIT_MAP = {
    "km/h": "kilometres per hour",
    "kph": "kilometres per hour",
    "mph": "miles per hour",
    "kg": "kilograms",
    "mg": "milligrams",
    "km": "kilometres",
    "cm": "centimetres",
    "mm": "millimetres",
    "ml": "millilitres",
    "ft": "feet",
    "in": "inches",
    "g": "grams",
    "m": "metres",
    "l": "litres",
    "m2": "square metres",
    "m3": "cubic metres",
    "m^2": "square metres",
    "m^3": "cubic metres",
    "\u00b0c": "degrees Celsius",
    "\u00b0f": "degrees Fahrenheit",
}

_CURRENCY_INFO = {
    "\u00a3": ("pounds", "pence"),
    "$": ("dollars", "cents"),
    "\u20ac": ("euros", "cents"),
    "\u00a5": ("yen", "sen"),
    "\u20b9": ("rupees", "paise"),
    "\u20a9": ("won", "jeon"),
    "\u20bd": ("roubles", "kopeks"),
    "\u20ba": ("lira", "kurus"),
}

_CURRENCY_CODE_INFO = {
    "GBP": ("pounds", "pence"),
    "USD": ("US dollars", "cents"),
    "EUR": ("euros", "cents"),
    "JPY": ("yen", "sen"),
    "CNY": ("yuan", "fen"),
    "RMB": ("yuan", "fen"),
    "INR": ("rupees", "paise"),
    "KRW": ("won", "jeon"),
    "AUD": ("Australian dollars", "cents"),
    "CAD": ("Canadian dollars", "cents"),
    "NZD": ("New Zealand dollars", "cents"),
    "SGD": ("Singapore dollars", "cents"),
    "HKD": ("Hong Kong dollars", "cents"),
    "CHF": ("Swiss francs", "rappen"),
}

_UNIT_PATTERN = "|".join(
    sorted((re.escape(key) for key in _UNIT_MAP), key=len, reverse=True)
)
_CURRENCY_SYMBOL_PATTERN = "|".join(re.escape(key) for key in _CURRENCY_INFO)
_CURRENCY_CODE_PATTERN = "|".join(sorted(_CURRENCY_CODE_INFO, key=len, reverse=True))


def _integer_to_words(value):
    number = int(value)
    if number < 0:
        return f"minus {_integer_to_words(abs(number))}"
    if number < 20:
        return _SMALL_NUMBERS[number]
    if number < 100:
        tens = (number // 10) * 10
        remainder = number % 10
        return _TENS[tens] if remainder == 0 else f"{_TENS[tens]}-{_SMALL_NUMBERS[remainder]}"
    if number < 1000:
        hundreds = number // 100
        remainder = number % 100
        prefix = f"{_SMALL_NUMBERS[hundreds]} hundred"
        return prefix if remainder == 0 else f"{prefix} and {_integer_to_words(remainder)}"
    for scale_value, scale_label in _SCALES:
        if number >= scale_value:
            major = number // scale_value
            remainder = number % scale_value
            prefix = f"{_integer_to_words(major)} {scale_label}"
            if remainder == 0:
                return prefix
            connector = " and " if remainder < 100 else " "
            return f"{prefix}{connector}{_integer_to_words(remainder)}"
    return str(number)


def _number_to_words(raw_number):
    text = str(raw_number or "").replace(",", "").strip()
    if not text:
        return ""
    negative = text.startswith("-")
    if negative:
        text = text[1:]
    if "." in text:
        left, right = text.split(".", 1)
        left_words = _integer_to_words(int(left or "0"))
        right_digits = " ".join(_SMALL_NUMBERS[int(ch)] for ch in right if ch.isdigit())
        words = f"{left_words} point {right_digits}".strip()
    else:
        words = _integer_to_words(int(text))
    return f"minus {words}" if negative else words


def _year_to_words(raw_year):
    year = int(str(raw_year).replace(",", ""))
    if year < 1000 or year > 2099:
        return _integer_to_words(year)
    if 2000 <= year <= 2009:
        remainder = year % 1000
        return "two thousand" if remainder == 0 else f"two thousand and {_integer_to_words(remainder)}"
    if 2010 <= year <= 2099:
        suffix = year % 100
        return f"twenty {_integer_to_words(suffix)}"
    first = year // 100
    last = year % 100
    if last == 0:
        return f"{_integer_to_words(first)} hundred"
    if last < 10:
        return f"{_integer_to_words(first)} oh {_integer_to_words(last)}"
    return f"{_integer_to_words(first)} {_integer_to_words(last)}"


def _currency_amount_to_words(amount, unit, minor_unit):
    cleaned = str(amount or "").replace(",", "")
    if "." in cleaned:
        major, minor = cleaned.split(".", 1)
        major_text = _integer_to_words(int(major or "0"))
        minor = (minor + "00")[:2]
        spoken = f"{major_text} {unit}"
        if int(minor):
            spoken = f"{spoken} and {_integer_to_words(int(minor))} {minor_unit}"
        return spoken
    return f"{_integer_to_words(int(cleaned or '0'))} {unit}"


def _normalise_currency_symbol(match):
    symbol = match.group("symbol")
    amount = match.group("amount")
    unit, minor_unit = _CURRENCY_INFO.get(symbol, ("pounds", "pence"))
    return _currency_amount_to_words(amount, unit, minor_unit)


def _normalise_currency_code(match):
    code = str(match.group("code") or "").upper()
    amount = match.group("amount")
    unit, minor_unit = _CURRENCY_CODE_INFO.get(code, ("pounds", "pence"))
    return _currency_amount_to_words(amount, unit, minor_unit)


def _normalise_percentage(match):
    return f"{_number_to_words(match.group('number'))} percent"


def _normalise_unit(match):
    number = _number_to_words(match.group("number"))
    unit = _UNIT_MAP.get(str(match.group("unit") or "").lower(), match.group("unit"))
    return f"{number} {unit}"


def _normalise_year(match):
    return _year_to_words(match.group("year"))


def _normalise_number(match):
    return _number_to_words(match.group("number"))


def normalize_ielts_tts_text(text):
    spoken = re.sub(r"\s+", " ", str(text or "").strip())
    if not spoken:
        return spoken

    spoken = re.sub(
        rf"(?P<symbol>{_CURRENCY_SYMBOL_PATTERN})\s*(?P<amount>\d[\d,]*(?:\.\d+)?)",
        _normalise_currency_symbol,
        spoken,
    )
    spoken = re.sub(
        rf"\b(?P<code>{_CURRENCY_CODE_PATTERN})\s*(?P<amount>\d[\d,]*(?:\.\d+)?)\b",
        _normalise_currency_code,
        spoken,
        flags=re.IGNORECASE,
    )
    spoken = re.sub(
        r"(?P<number>\d[\d,]*(?:\.\d+)?)\s*%",
        _normalise_percentage,
        spoken,
    )
    spoken = re.sub(
        rf"(?P<number>\d[\d,]*(?:\.\d+)?)\s*(?P<unit>{_UNIT_PATTERN})\b",
        _normalise_unit,
        spoken,
        flags=re.IGNORECASE,
    )
    spoken = re.sub(
        r"\b(?P<year>(?:19|20)\d{2})\b",
        _normalise_year,
        spoken,
    )
    spoken = re.sub(
        r"\b(?P<number>\d[\d,]*(?:\.\d+)?)\b",
        _normalise_number,
        spoken,
    )
    spoken = re.sub(r"\s+", " ", spoken).strip()
    return spoken

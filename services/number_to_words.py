import math

_ONES = [
    "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
    "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
    "Seventeen", "Eighteen", "Nineteen",
]

_TENS = [
    "", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety",
]


def to_legal_dollar_string(amount: float) -> str:
    """Convert a dollar amount to legal format:
    'Five Hundred Thousand and 00/100 Dollars ($500,000.00)'
    """
    dollars = int(math.floor(amount))
    cents = int(round((amount - dollars) * 100))
    words = convert_to_words(dollars)
    number_formatted = f"{amount:,.2f}"
    return f"{words} and {cents:02d}/100 Dollars (${number_formatted})"


def convert_to_words(number: int) -> str:
    """Convert an integer to words (e.g. 500000 -> 'Five Hundred Thousand')."""
    if number == 0:
        return "Zero"

    parts = []

    if number >= 1_000_000_000:
        parts.append(_convert_group(number // 1_000_000_000) + " Billion")
        number %= 1_000_000_000

    if number >= 1_000_000:
        parts.append(_convert_group(number // 1_000_000) + " Million")
        number %= 1_000_000

    if number >= 1_000:
        parts.append(_convert_group(number // 1_000) + " Thousand")
        number %= 1_000

    if number > 0:
        parts.append(_convert_group(number))

    return " ".join(parts)


def _convert_group(number: int) -> str:
    parts = []

    if number >= 100:
        parts.append(_ONES[number // 100] + " Hundred")
        number %= 100

    if number >= 20:
        ten = _TENS[number // 10]
        one = "-" + _ONES[number % 10] if number % 10 > 0 else ""
        parts.append(ten + one)
    elif number > 0:
        parts.append(_ONES[number])

    return " ".join(parts)

"""
Automated test suite for the LOI Generator.
Generates documents with various configurations and validates the output.

Run: python tests/test_document_generator.py
"""

import os
import sys

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docx import Document
from lxml import etree

from services.document_generator import DocumentGenerator
from services.loi_form_data import (
    LoiFormData,
    DepositStructure,
    DueDiligenceType,
    CommissionType,
    SignatureBlockType,
    SignatureEntity,
)
from services.number_to_words import to_legal_dollar_string, convert_to_words

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "Templates", "LOI_Template.docx")

_passed = 0
_failed = 0
_failures = []


def _assert(name: str, condition: bool):
    global _passed, _failed
    if condition:
        print(f"  PASS: {name}")
        _passed += 1
    else:
        print(f"  FAIL: {name}")
        _failed += 1
        _failures.append(name)


def _qn(tag: str) -> str:
    prefix, local = tag.split(":")
    ns = {"w": W_NS}
    return f"{{{ns[prefix]}}}{local}"


def _generate_and_extract(form: LoiFormData):
    gen = DocumentGenerator()
    buf = gen.generate(TEMPLATE_PATH, form)
    doc = Document(buf)
    body = doc.element.body
    return body, doc


def _get_visible_text(body) -> str:
    """Get only visible text (w:t), excluding deleted text (w:delText)."""
    texts = []
    for t in body.iter(_qn("w:t")):
        texts.append(t.text or "")
    return "".join(texts)


def _get_all_text(body) -> str:
    """Get all text including tracked-change deletions."""
    texts = []
    for t in body.iter(_qn("w:t")):
        texts.append(t.text or "")
    for dt in body.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}delText"):
        texts.append(dt.text or "")
    return "".join(texts)


def _get_header_texts(doc) -> list:
    texts = []
    for section in doc.sections:
        h = section.header
        if h:
            parts = []
            for t in h._element.iter(_qn("w:t")):
                parts.append(t.text or "")
            for dt in h._element.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}delText"):
                parts.append(dt.text or "")
            texts.append("".join(parts))
    return texts


def _base_form() -> LoiFormData:
    return LoiFormData(
        date="March 1, 2026",
        property_address="123 Main St, Springfield, IL 62701",
        seller_address_line1="Acme Properties LLC",
        seller_address_line2="100 Commerce Dr",
        seller_address_line3="Springfield, IL 62701",
        attention_name="John Doe",
        salutation="Mr. Doe",
        seller_name="Acme Properties LLC",
        purchase_price=500000,
        initial_deposit=10000,
        additional_deposit=10000,
        monthly_release_amount=5000,
        legal_reimbursement_amount=5000,
        extension_deposit_amount=5000,
        monthly_closing_extension_deposit=25000,
        due_diligence_days=120,
        governmental_approvals_days=150,
        assemblage_days=90,
        closing_days=30,
        closing_extension_months=6,
        lease_end_date="May 31, 2026",
        lease_termination_days=60,
        deposit_structure=DepositStructure.GOVERNMENTAL_APPROVALS_GOING_HARD,
        include_legal_reimbursement=False,
        due_diligence_type=DueDiligenceType.STANDARD,
        include_closing_extension=False,
        commission_type=CommissionType.SELLER_PAYS_LISTING_AGENT,
        include_option_to_extend=True,
        include_existing_leases=True,
        include_delivered_vacant=False,
        include_lease_termination=False,
        include_right_to_negotiate_with_tenants=False,
        include_seller_rollover=False,
        signature_block_type=SignatureBlockType.INDIVIDUAL,
        seller_name_signature="John Doe",
        signature_entities=[SignatureEntity(company_name="Test Corp")],
        parcel_ids=["12-345-678"],
        prepared_by_first_name="Jake",
        prepared_by_last_name="Miller",
    )


# =========================================================================
# Tests
# =========================================================================

def test_basic_field_replacement():
    form = _base_form()
    body, doc = _generate_and_extract(form)
    text = _get_all_text(body)
    headers = _get_header_texts(doc)

    _assert("Basic: Date in body", "March 1, 2026" in text)
    _assert("Basic: Property address in body", "123 Main St, Springfield, IL 62701" in text)
    _assert("Basic: Property address in header", any("123 Main St, Springfield, IL 62701" in h for h in headers))
    _assert("Basic: Date in header", any("March 1, 2026" in h for h in headers))
    _assert("Basic: Seller name in body", "Acme Properties LLC" in text)
    _assert("Basic: Attention name", "John Doe" in text)
    _assert("Basic: Salutation", "Mr. Doe" in text)
    _assert("Basic: Purchase price words", "Five Hundred Thousand" in text)
    _assert("Basic: Purchase price number", "500,000" in text)


def test_address_line_collapse_fixed():
    # All 3 lines filled
    form1 = _base_form()
    body1, _ = _generate_and_extract(form1)
    text1 = _get_all_text(body1)
    _assert("Address: All 3 lines present",
            "Acme Properties LLC" in text1 and "100 Commerce Dr" in text1 and "Springfield, IL 62701" in text1)

    # Middle line empty
    form2 = _base_form()
    form2.seller_address_line2 = ""
    body2, _ = _generate_and_extract(form2)
    vis2 = _get_visible_text(body2)
    _assert("Address: Empty middle line - no unreplaced placeholder", "[____________________]" not in vis2)
    _assert("Address: Empty middle line - Line 1 present", "Acme Properties LLC" in vis2)
    _assert("Address: Empty middle line - Line 3 present", "Springfield, IL 62701" in vis2)

    # Only first line
    form3 = _base_form()
    form3.seller_address_line2 = ""
    form3.seller_address_line3 = ""
    body3, _ = _generate_and_extract(form3)
    vis3 = _get_visible_text(body3)
    _assert("Address: Only Line 1 - no unreplaced placeholder", "[____________________]" not in vis3)
    _assert("Address: Only Line 1 - value present", "Acme Properties LLC" in vis3)


def test_deposit_scenarios():
    # Governmental Approvals
    form1 = _base_form()
    form1.deposit_structure = DepositStructure.GOVERNMENTAL_APPROVALS_GOING_HARD
    form1.initial_deposit = 15000
    form1.additional_deposit = 25000
    body1, _ = _generate_and_extract(form1)
    text1 = _get_all_text(body1)
    _assert("Deposit Gov: Initial deposit in doc", "Fifteen Thousand" in text1)

    # Due Diligence
    form2 = _base_form()
    form2.deposit_structure = DepositStructure.DUE_DILIGENCE_GOING_HARD
    form2.initial_deposit = 10000
    form2.additional_deposit = 20000
    body2, _ = _generate_and_extract(form2)
    text2 = _get_all_text(body2)
    _assert("Deposit DD: Initial deposit in doc", "Ten Thousand" in text2)

    # Monthly
    form3 = _base_form()
    form3.deposit_structure = DepositStructure.MONTHLY_GOING_HARD
    form3.initial_deposit = 10000
    form3.monthly_release_amount = 7500
    body3, _ = _generate_and_extract(form3)
    text3 = _get_all_text(body3)
    _assert("Deposit Monthly: Monthly release in doc", "Seven Thousand Five Hundred" in text3)


def test_due_diligence_scenarios():
    form1 = _base_form()
    form1.due_diligence_type = DueDiligenceType.STANDARD
    body1, _ = _generate_and_extract(form1)
    text1 = _get_all_text(body1)
    _assert("DD Standard: DD days in doc", "one hundred twenty (120)" in text1)
    _assert("DD Standard: Gov approvals days", "one hundred fifty (150)" in text1)

    form2 = _base_form()
    form2.due_diligence_type = DueDiligenceType.WITH_ASSEMBLAGE
    body2, _ = _generate_and_extract(form2)
    text2 = _get_all_text(body2)
    _assert("DD Assemblage: Assemblage days in doc", "ninety (90)" in text2)


def test_closing_extension():
    form1 = _base_form()
    form1.include_closing_extension = True
    form1.closing_extension_months = 6
    form1.monthly_closing_extension_deposit = 25000
    body1, _ = _generate_and_extract(form1)
    text1 = _get_all_text(body1)
    _assert("Closing Ext: Extension months in doc", "six (6)" in text1)

    form2 = _base_form()
    form2.include_closing_extension = False
    body2, _ = _generate_and_extract(form2)
    vis2 = _get_visible_text(body2)
    _assert("Closing Ext: No extension - no stale placeholder", "[six (6)]" not in vis2)


def test_commission_scenarios():
    form1 = _base_form()
    form1.commission_type = CommissionType.SELLER_PAYS_LISTING_AGENT
    _generate_and_extract(form1)
    _assert("Commission: Seller Pays generates OK", True)

    form2 = _base_form()
    form2.commission_type = CommissionType.SUBTEXT_PAYS
    form2.broker_name = "Jones Realty"
    body2, _ = _generate_and_extract(form2)
    text2 = _get_all_text(body2)
    _assert("Commission: Broker name in doc", "Jones Realty" in text2)

    form3 = _base_form()
    form3.commission_type = CommissionType.NO_BROKERS
    _generate_and_extract(form3)
    _assert("Commission: No Brokers generates OK", True)


def test_lease_scenarios():
    form1 = _base_form()
    form1.include_existing_leases = True
    form1.lease_end_date = "December 31, 2026"
    body1, _ = _generate_and_extract(form1)
    text1 = _get_all_text(body1)
    _assert("Lease: End date in doc", "December 31, 2026" in text1)

    form2 = _base_form()
    form2.include_existing_leases = False
    form2.include_delivered_vacant = True
    _generate_and_extract(form2)
    _assert("Lease: Delivered Vacant generates OK", True)

    form3 = _base_form()
    form3.include_lease_termination = True
    form3.lease_termination_days = 60
    body3, _ = _generate_and_extract(form3)
    text3 = _get_all_text(body3)
    _assert("Lease: Termination days in doc", "sixty (60)" in text3)

    form4 = _base_form()
    form4.include_right_to_negotiate_with_tenants = True
    _generate_and_extract(form4)
    _assert("Lease: Right to Negotiate generates OK", True)


def test_option_to_extend():
    form1 = _base_form()
    form1.include_option_to_extend = True
    form1.extension_deposit_amount = 8000
    body1, _ = _generate_and_extract(form1)
    text1 = _get_all_text(body1)
    _assert("Option to Extend: Extension deposit in doc", "Eight Thousand" in text1)

    form2 = _base_form()
    form2.include_option_to_extend = False
    _generate_and_extract(form2)
    _assert("Option to Extend: Excluded generates OK", True)


def test_signature_block():
    form1 = _base_form()
    form1.signature_block_type = SignatureBlockType.INDIVIDUAL
    form1.seller_name_signature = "John Q. Seller"
    body1, _ = _generate_and_extract(form1)
    text1 = _get_all_text(body1)
    _assert("Signature: Individual - seller name", "John Q. Seller" in text1)

    form2 = _base_form()
    form2.signature_block_type = SignatureBlockType.COMPANY_ENTITY
    form2.signature_entities = [SignatureEntity(company_name="Alpha Corp")]
    body2, _ = _generate_and_extract(form2)
    text2 = _get_all_text(body2)
    _assert("Signature: Company name in doc", "Alpha Corp" in text2)

    form3 = _base_form()
    form3.signature_block_type = SignatureBlockType.COMPANY_ENTITY
    form3.signature_entities = [
        SignatureEntity(company_name="Alpha Corp"),
        SignatureEntity(company_name="Beta LLC"),
    ]
    body3, _ = _generate_and_extract(form3)
    text3 = _get_all_text(body3)
    _assert("Signature: Multiple entities", "Alpha Corp" in text3 and "Beta LLC" in text3)
    _assert("Signature: Collectively phrasing", "collectively" in text3)


def test_parcel_ids():
    form1 = _base_form()
    form1.parcel_ids = ["12-345-678"]
    body1, _ = _generate_and_extract(form1)
    text1 = _get_all_text(body1)
    _assert("Parcel: Single ID in doc", "12-345-678" in text1)

    form2 = _base_form()
    form2.parcel_ids = ["12-345-678", "98-765-432"]
    body2, _ = _generate_and_extract(form2)
    text2 = _get_all_text(body2)
    _assert("Parcel: Multiple IDs", "12-345-678" in text2 and "98-765-432" in text2)


def test_stale_value_guards():
    # AdditionalDeposit stale on Monthly
    form = _base_form()
    form.deposit_structure = DepositStructure.MONTHLY_GOING_HARD
    form.additional_deposit = 99999
    form.monthly_release_amount = 5000
    body, _ = _generate_and_extract(form)
    vis = _get_visible_text(body)
    _assert("Stale: AdditionalDeposit not injected on Monthly", "Ninety-Nine Thousand" not in vis)

    # BrokerName stale on SellerPays
    form2 = _base_form()
    form2.commission_type = CommissionType.SELLER_PAYS_LISTING_AGENT
    form2.broker_name = "Stale Broker Corp"
    body2, _ = _generate_and_extract(form2)
    vis2 = _get_visible_text(body2)
    _assert("Stale: BrokerName not injected on SellerPays", "Stale Broker Corp" not in vis2)

    # AssemblageDays stale on Standard
    form3 = _base_form()
    form3.due_diligence_type = DueDiligenceType.STANDARD
    form3.assemblage_days = 999
    body3, _ = _generate_and_extract(form3)
    vis3 = _get_visible_text(body3)
    _assert("Stale: AssemblageDays not injected on Standard DD", "nine hundred ninety-nine (999)" not in vis3)


def test_number_to_words():
    _assert("NTW: Zero", convert_to_words(0) == "Zero")
    _assert("NTW: One", convert_to_words(1) == "One")
    _assert("NTW: 10", convert_to_words(10) == "Ten")
    _assert("NTW: 21", convert_to_words(21) == "Twenty-One")
    _assert("NTW: 100", convert_to_words(100) == "One Hundred")
    _assert("NTW: 500000", convert_to_words(500000) == "Five Hundred Thousand")
    _assert("NTW: 1000000", convert_to_words(1000000) == "One Million")
    _assert("NTW: Legal format $10K", to_legal_dollar_string(10000) == "Ten Thousand and 00/100 Dollars ($10,000.00)")


def test_explicit_guards():
    # ClosingExtensionMonths when disabled
    form = _base_form()
    form.include_closing_extension = False
    form.closing_extension_months = 12
    form.monthly_closing_extension_deposit = 50000
    body, _ = _generate_and_extract(form)
    vis = _get_visible_text(body)
    _assert("Guard: ClosingExtensionMonths not replaced when disabled", "twelve (12)" not in vis)
    _assert("Guard: MonthlyClosingExtensionDeposit not replaced when disabled", "Fifty Thousand and 00/100" not in vis)

    # AssemblageDays on Standard
    form2 = _base_form()
    form2.due_diligence_type = DueDiligenceType.STANDARD
    form2.assemblage_days = 45
    body2, _ = _generate_and_extract(form2)
    vis2 = _get_visible_text(body2)
    _assert("Guard: AssemblageDays not replaced on Standard DD", "forty-five (45)" not in vis2)

    # LeaseTerminationDays when disabled
    form3 = _base_form()
    form3.include_lease_termination = False
    form3.lease_termination_days = 90
    form3.assemblage_days = 45
    body3, _ = _generate_and_extract(form3)
    vis3 = _get_visible_text(body3)
    _assert("Guard: LeaseTerminationDays not replaced when disabled", "ninety (90)" not in vis3)

    # BrokerName when NoBrokers
    form4 = _base_form()
    form4.commission_type = CommissionType.NO_BROKERS
    form4.broker_name = "Ghost Broker"
    body4, _ = _generate_and_extract(form4)
    vis4 = _get_visible_text(body4)
    _assert("Guard: BrokerName not replaced when NoBrokers", "Ghost Broker" not in vis4)

    # LeaseEndDate when existing leases disabled
    form5 = _base_form()
    form5.include_existing_leases = False
    form5.include_delivered_vacant = True
    form5.lease_end_date = "January 1, 2099"
    body5, _ = _generate_and_extract(form5)
    vis5 = _get_visible_text(body5)
    _assert("Guard: LeaseEndDate not replaced when disabled", "January 1, 2099" not in vis5)


# =========================================================================
# Main
# =========================================================================
if __name__ == "__main__":
    if not os.path.exists(TEMPLATE_PATH):
        print(f"ERROR: Template not found at {TEMPLATE_PATH}")
        sys.exit(1)

    print("=== LOI Generator Test Suite (Python) ===\n")

    test_basic_field_replacement()
    test_address_line_collapse_fixed()
    test_deposit_scenarios()
    test_due_diligence_scenarios()
    test_closing_extension()
    test_commission_scenarios()
    test_lease_scenarios()
    test_option_to_extend()
    test_signature_block()
    test_parcel_ids()
    test_stale_value_guards()
    test_number_to_words()
    test_explicit_guards()

    print(f"\n=== Results: {_passed} passed, {_failed} failed ===")
    if _failures:
        print("\nFailed tests:")
        for f in _failures:
            print(f"  - {f}")

    sys.exit(1 if _failed > 0 else 0)

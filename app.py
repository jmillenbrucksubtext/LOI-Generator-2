import streamlit as st
from datetime import datetime

from services.loi_form_data import (
    LoiFormData,
    DepositStructure,
    DueDiligenceType,
    CommissionType,
    SignatureBlockType,
    SignatureEntity,
)
from services.number_to_words import to_legal_dollar_string, convert_to_words
from services.document_generator import DocumentGenerator
import base64
import os

# ---------------------------------------------------------------------------
# Page config — wide layout for side-by-side
# ---------------------------------------------------------------------------
st.set_page_config(page_title="LOI Generator", layout="wide")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def fmt_period(value: int, unit: str = "days") -> str:
    return f"{convert_to_words(value).lower()} ({value}) {unit}"


def fmt_dollar(amount: float) -> str:
    return to_legal_dollar_string(amount)


def money_input(label: str, default: float = 0.0, key: str = None) -> float:
    """Text input that displays dollar amounts with comma separators."""
    input_key = key or f"money_{label}"
    val_key = f"_mval_{input_key}"

    if val_key not in st.session_state:
        st.session_state[val_key] = default
        st.session_state[input_key] = f"{default:,.2f}"

    def _reformat():
        raw = st.session_state[input_key]
        try:
            parsed = float(raw.replace(",", "").replace("$", "").strip())
            st.session_state[val_key] = max(0.0, parsed)
        except (ValueError, TypeError):
            pass
        st.session_state[input_key] = f"{st.session_state[val_key]:,.2f}"

    st.text_input(label, key=input_key, on_change=_reformat)
    return st.session_state[val_key]


def dollar_preview(amount: float):
    if amount > 0:
        st.markdown(
            f'<div class="legal-preview">{fmt_dollar(amount)}</div>',
            unsafe_allow_html=True,
        )


def period_preview(value: int, unit: str = "days"):
    if value > 0:
        st.markdown(f'<div class="legal-preview">{fmt_period(value, unit)}</div>', unsafe_allow_html=True)


def _v(text: str, placeholder: str = "___") -> str:
    """Show tracked-change style: strikethrough old + red new if filled, green placeholder if not."""
    if text and text.strip():
        return (
            f'<span class="tc-del">{placeholder}</span>'
            f'<span class="tc-ins">{text.strip()}</span>'
        )
    return f'<span class="tc-empty">{placeholder}</span>'


def _vn(text: str, placeholder: str = "___") -> str:
    """Show just the value (no tracked-change strikethrough) — for when placeholder isn't meaningful."""
    if text and text.strip():
        return f'<span class="tc-ins">{text.strip()}</span>'
    return f'<span class="tc-empty">{placeholder}</span>'


# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    .section-header {
        border-top: 3px solid #c1d100;
        padding-top: 0.5rem;
        margin-top: 1rem;
        margin-bottom: 0.5rem;
        font-weight: 600;
        font-size: 1.1rem;
    }
    .legal-preview {
        color: #9e9e90;
        font-size: 0.85rem;
        margin-top: 0.25rem;
    }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    /* Document preview — scrollable page viewer */
    .doc-scroll {
        max-height: 85vh;
        overflow-y: auto;
        border: 1px solid #555;
        border-radius: 4px;
        background: #444;
        padding: 12px;
    }
    .doc-page {
        background: #fff;
        color: #222;
        font-family: 'Times New Roman', Times, serif;
        font-size: 10pt;
        line-height: 1.4;
        /* US Letter: 8.5"×11" at 96dpi = 816×1056px, 1" margins = 96px */
        width: 816px;
        min-height: 1056px;
        padding: 96px;
        margin: 0 auto 12px auto;
        box-shadow: 0 2px 8px rgba(0,0,0,0.3);
        position: relative;
        box-sizing: border-box;
    }
    .doc-page:last-child {
        margin-bottom: 0;
    }
    .doc-page p {
        margin: 0.4em 0;
        text-align: justify;
    }
    .doc-page .hdr-right {
        text-align: right;
        font-size: 9pt;
        margin: 0;
    }
    .doc-page .addr-line {
        margin: 0.1em 0;
    }
    .doc-page .re-line {
        margin: 0.8em 0;
    }
    .doc-page .section-item {
        margin: 0.6em 0;
        padding-left: 3em;
        text-indent: -3em;
    }
    .doc-page .section-item .sec-label {
        font-weight: bold;
        text-decoration: underline;
    }
    .doc-page .closing-text {
        text-indent: 2em;
    }
    .doc-page .sig-block {
        margin-top: 1.5em;
    }
    .doc-page .sig-line {
        margin: 0.2em 0;
    }
    .doc-page .exhibit-header {
        text-align: center;
        font-weight: bold;
        margin-top: 1em;
    }
    .doc-page .exhibit-center {
        text-align: center;
    }
    .doc-page .cont-header {
        text-align: right;
        font-size: 10pt;
        margin-bottom: 1.5em;
        color: #222;
        line-height: 1.6;
    }
    .doc-page hr {
        border: none;
        border-top: 1px solid #ccc;
        margin: 0.8em 0;
    }
    .doc-page .photo-preview {
        display: block;
        max-width: 100%;
        max-height: 500px;
        margin: 1em auto;
    }
    /* Tracked change styles */
    .tc-del {
        color: #c0392b;
        text-decoration: line-through;
        font-size: 0.95em;
    }
    .tc-ins {
        color: #c0392b;
    }
    .tc-empty {
        color: #27ae60;
    }
    /* Make preview column sticky */
    [data-testid="stVerticalBlock"] > div:has(.doc-scroll) {
        position: sticky;
        top: 2rem;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.image("https://subtextliving.com/wp-content/uploads/2023/08/subtext-primary-logo.svg", width=220)
st.markdown("#### LOI Generator")

# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------
if "parcel_ids" not in st.session_state:
    st.session_state.parcel_ids = [""]
if "entities" not in st.session_state:
    st.session_state.entities = [{"company_name": ""}]
if "generated_file" not in st.session_state:
    st.session_state.generated_file = None
    st.session_state.generated_filename = None

# ---------------------------------------------------------------------------
# Two-column layout: Form (left) | Preview (right)
# ---------------------------------------------------------------------------
form_col, preview_col = st.columns([1, 1], gap="large")

# ============================= LEFT: FORM =================================
with form_col:

    # --- Letter Details ---
    st.markdown('<div class="section-header">Letter Details</div>', unsafe_allow_html=True)
    c1, c2 = st.columns([1, 2])
    with c1:
        _now = datetime.now()
        date_val = st.text_input("Date", value=f"{_now.strftime('%B')} {_now.day}, {_now.year}")
    with c2:
        property_address = st.text_input("Property Address", placeholder="Address, City, State Zip Code")

    seller_addr1 = st.text_input("Seller Address Line 1 (Seller Entity / Individual)")
    seller_addr2 = st.text_input("Seller Address Line 2 (Owner Address)")
    seller_addr3 = st.text_input("Seller Address Line 3 (City, State Zip Code)")

    ca, cb, cc = st.columns(3)
    with ca:
        attention_name = st.text_input("Attention Name", placeholder="John Smith")
    with cb:
        salutation = st.text_input("Salutation", placeholder="Mr. Smith")
    with cc:
        seller_name = st.text_input("Seller Name", placeholder="ABC Properties LLC")

    # --- B. Purchase Price ---
    st.markdown('<div class="section-header">B. Purchase Price</div>', unsafe_allow_html=True)
    purchase_price = money_input("Purchase Price ($)", default=0.0, key="purchase_price")
    dollar_preview(purchase_price)

    # --- C. Deposit ---
    st.markdown('<div class="section-header">C. Deposit</div>', unsafe_allow_html=True)
    deposit_options = [e.value for e in DepositStructure]
    deposit_choice = st.radio("Deposit Scenario", deposit_options, index=0, horizontal=True)
    deposit_structure = DepositStructure(deposit_choice)

    cd1, cd2 = st.columns(2)
    with cd1:
        initial_deposit = money_input("Initial Deposit ($)", default=10000.0, key="initial_deposit")
        dollar_preview(initial_deposit)
    with cd2:
        if deposit_structure in (DepositStructure.GOVERNMENTAL_APPROVALS_GOING_HARD, DepositStructure.DUE_DILIGENCE_GOING_HARD):
            additional_deposit = money_input("Additional Deposit ($)", default=10000.0, key="additional_deposit")
            dollar_preview(additional_deposit)
        else:
            additional_deposit = 10000.0
        if deposit_structure == DepositStructure.MONTHLY_GOING_HARD:
            monthly_release = money_input("Monthly Release Amount ($)", default=5000.0, key="monthly_release")
            dollar_preview(monthly_release)
        else:
            monthly_release = 5000.0

    st.divider()
    include_legal_reimb = st.checkbox("Include Legal Reimbursement Fee")
    if include_legal_reimb:
        legal_reimb_amount = money_input("Legal Reimbursement Amount ($)", default=5000.0, key="legal_reimb")
        dollar_preview(legal_reimb_amount)
    else:
        legal_reimb_amount = 5000.0

    # --- E. Due Diligence ---
    st.markdown('<div class="section-header">E. Due Diligence and Governmental Approvals</div>', unsafe_allow_html=True)
    dd_options = [e.value for e in DueDiligenceType]
    dd_choice = st.radio("Due Diligence Type", dd_options, index=0, horizontal=True)
    dd_type = DueDiligenceType(dd_choice)

    ce1, ce2 = st.columns(2)
    with ce1:
        dd_days = st.number_input("Due Diligence Period (days)", min_value=0, value=120, step=1)
        period_preview(dd_days)
    with ce2:
        ga_days = st.number_input("Governmental Approvals Period (days)", min_value=0, value=150, step=1)
        period_preview(ga_days)

    if dd_type == DueDiligenceType.WITH_ASSEMBLAGE:
        assemblage_days = st.number_input("Assemblage Period (days)", min_value=0, value=90, step=1)
        period_preview(assemblage_days)
    else:
        assemblage_days = 90

    # --- F. Closing ---
    st.markdown('<div class="section-header">F. Closing</div>', unsafe_allow_html=True)
    closing_days = st.number_input("Closing Period (days)", min_value=0, value=30, step=1)
    period_preview(closing_days)

    include_closing_ext = st.checkbox("Include Closing Extension")
    if include_closing_ext:
        cf1, cf2 = st.columns(2)
        with cf1:
            closing_ext_months = st.number_input("Closing Extension Duration (months)", min_value=0, value=6, step=1)
            period_preview(closing_ext_months, "months")
        with cf2:
            monthly_closing_ext_deposit = money_input("Monthly Closing Extension Deposit ($)", default=25000.0, key="monthly_closing_ext")
            dollar_preview(monthly_closing_ext_deposit)
    else:
        closing_ext_months = 6
        monthly_closing_ext_deposit = 25000.0

    # --- G. Commissions ---
    st.markdown('<div class="section-header">G. Commissions</div>', unsafe_allow_html=True)
    comm_options = [e.value for e in CommissionType]
    comm_choice = st.radio("Commission Type", comm_options, index=0, horizontal=True)
    commission_type = CommissionType(comm_choice)

    if commission_type == CommissionType.SUBTEXT_PAYS:
        broker_name = st.text_input("Broker Name", placeholder="Jones Realty")
    else:
        broker_name = ""

    # --- H. Option to Extend ---
    st.markdown('<div class="section-header">H. Option to Extend</div>', unsafe_allow_html=True)
    include_option_extend = st.checkbox("Include Option to Extend", value=True)
    if include_option_extend:
        ext_deposit = money_input("Extension Deposit Amount ($)", default=5000.0, key="ext_deposit")
        dollar_preview(ext_deposit)
    else:
        ext_deposit = 5000.0

    # --- I. Leases ---
    st.markdown('<div class="section-header">I. Leases</div>', unsafe_allow_html=True)
    include_existing_leases = st.checkbox("Existing Leases with End Date", value=True)
    if include_existing_leases:
        lease_end_date = st.text_input("Lease End Date", value="May 31, 2026")
    else:
        lease_end_date = "May 31, 2026"

    include_delivered_vacant = st.checkbox("Delivered Vacant")
    include_lease_termination = st.checkbox("Lease Termination Provision")
    if include_lease_termination:
        lease_term_days = st.number_input("Lease Termination Days", min_value=0, value=60, step=1)
        period_preview(lease_term_days)
    else:
        lease_term_days = 60

    include_negotiate_tenants = st.checkbox("Right to Negotiate with Existing Tenants")
    include_seller_rollover = st.checkbox("Seller Rollover Option")

    # --- Signature ---
    st.markdown('<div class="section-header">Signature</div>', unsafe_allow_html=True)
    sig_options = [e.value for e in SignatureBlockType]
    sig_choice = st.radio("Signature Block Type", sig_options, index=0, horizontal=True)
    sig_type = SignatureBlockType(sig_choice)

    if sig_type == SignatureBlockType.INDIVIDUAL:
        seller_name_sig = st.text_input("Seller Name (Signature Line)")
    else:
        seller_name_sig = ""
        st.markdown("**Company / Entity Name(s)**")
        entities_to_remove = None
        for i, entity in enumerate(st.session_state.entities):
            col_ent, col_btn = st.columns([6, 1])
            with col_ent:
                st.session_state.entities[i]["company_name"] = st.text_input(
                    f"Entity {i+1}", value=entity["company_name"],
                    key=f"entity_{i}", label_visibility="collapsed"
                )
            with col_btn:
                if len(st.session_state.entities) > 1:
                    if st.button("X", key=f"rm_entity_{i}"):
                        entities_to_remove = i
        if entities_to_remove is not None:
            st.session_state.entities.pop(entities_to_remove)
            st.rerun()
        if st.button("+ Add Entity"):
            st.session_state.entities.append({"company_name": ""})
            st.rerun()

    # --- Exhibit A ---
    st.markdown('<div class="section-header">Exhibit A</div>', unsafe_allow_html=True)
    st.markdown("**Parcel ID Number(s)**")
    parcels_to_remove = None
    for i, pid in enumerate(st.session_state.parcel_ids):
        col_p, col_pb = st.columns([6, 1])
        with col_p:
            st.session_state.parcel_ids[i] = st.text_input(
                f"Parcel {i+1}", value=pid,
                key=f"parcel_{i}", placeholder="12-345-678", label_visibility="collapsed"
            )
        with col_pb:
            if len(st.session_state.parcel_ids) > 1:
                if st.button("X", key=f"rm_parcel_{i}"):
                    parcels_to_remove = i
    if parcels_to_remove is not None:
        st.session_state.parcel_ids.pop(parcels_to_remove)
        st.rerun()
    if st.button("+ Add Parcel ID"):
        st.session_state.parcel_ids.append("")
        st.rerun()

    uploaded_photo = st.file_uploader("Property Photo (Exhibit A)", type=["png", "jpg", "jpeg", "gif", "bmp"])
    if uploaded_photo and uploaded_photo.size > 10 * 1024 * 1024:
        st.error("Photo must be under 10 MB.")
        uploaded_photo = None
    elif uploaded_photo:
        st.caption(f"{uploaded_photo.name} ({uploaded_photo.size / 1024:.0f} KB)")

    # --- Prepared By ---
    st.markdown('<div class="section-header">Prepared By</div>', unsafe_allow_html=True)
    cpb1, cpb2 = st.columns(2)
    with cpb1:
        prep_first = st.text_input("First Name", placeholder="First name", key="prep_first")
    with cpb2:
        prep_last = st.text_input("Last Name", placeholder="Last name", key="prep_last")
    st.caption("This name will appear as the author of all tracked changes in the document.")

    # --- Generate ---
    st.divider()
    if st.session_state.generated_file is not None:
        st.download_button(
            label=f"Download: {st.session_state.generated_filename}",
            data=st.session_state.generated_file,
            file_name=st.session_state.generated_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )

    if st.button("Generate LOI", type="primary", use_container_width=True):
        missing = []
        if not property_address.strip():
            missing.append("Property Address")
        if not seller_name.strip():
            missing.append("Seller Name")
        if purchase_price <= 0:
            missing.append("Purchase Price")

        if missing:
            st.error(f"Please fill in: {', '.join(missing)}")
        else:
            form = LoiFormData(
                date=date_val,
                seller_address_line1=seller_addr1,
                seller_address_line2=seller_addr2,
                seller_address_line3=seller_addr3,
                attention_name=attention_name,
                property_address=property_address,
                salutation=salutation,
                seller_name=seller_name,
                purchase_price=purchase_price if purchase_price > 0 else None,
                initial_deposit=initial_deposit if initial_deposit > 0 else None,
                additional_deposit=additional_deposit if additional_deposit > 0 else None,
                monthly_release_amount=monthly_release if monthly_release > 0 else None,
                legal_reimbursement_amount=legal_reimb_amount if legal_reimb_amount > 0 else None,
                extension_deposit_amount=ext_deposit if ext_deposit > 0 else None,
                monthly_closing_extension_deposit=monthly_closing_ext_deposit if monthly_closing_ext_deposit > 0 else None,
                due_diligence_days=dd_days,
                governmental_approvals_days=ga_days,
                assemblage_days=assemblage_days,
                closing_days=closing_days,
                closing_extension_months=closing_ext_months,
                lease_end_date=lease_end_date,
                lease_termination_days=lease_term_days,
                broker_name=broker_name,
                seller_name_signature=seller_name_sig,
                signature_entities=[SignatureEntity(company_name=e["company_name"]) for e in st.session_state.entities],
                parcel_ids=list(st.session_state.parcel_ids),
                property_photo_bytes=uploaded_photo.read() if uploaded_photo else None,
                property_photo_content_type=uploaded_photo.type if uploaded_photo else None,
                property_photo_filename=uploaded_photo.name if uploaded_photo else None,
                deposit_structure=deposit_structure,
                include_legal_reimbursement=include_legal_reimb,
                due_diligence_type=dd_type,
                include_closing_extension=include_closing_ext,
                commission_type=commission_type,
                include_option_to_extend=include_option_extend,
                include_existing_leases=include_existing_leases,
                include_delivered_vacant=include_delivered_vacant,
                include_lease_termination=include_lease_termination,
                include_right_to_negotiate_with_tenants=include_negotiate_tenants,
                include_seller_rollover=include_seller_rollover,
                signature_block_type=sig_type,
                prepared_by_first_name=prep_first,
                prepared_by_last_name=prep_last,
            )

            script_dir = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(script_dir, "Templates", "LOI_Template.docx")

            if not os.path.exists(template_path):
                st.error(f"Template not found at: {template_path}")
            else:
                try:
                    with st.spinner("Generating LOI..."):
                        generator = DocumentGenerator()
                        result = generator.generate(template_path, form)

                    entity_name = seller_name.strip() if seller_name.strip() else "Unknown"
                    f_initial = prep_first.strip()[0].upper() if prep_first.strip() else ""
                    l_initial = prep_last.strip()[0].upper() if prep_last.strip() else ""
                    initials = f_initial + l_initial
                    date_str = datetime.now().strftime("%Y%m%d")
                    if initials:
                        filename = f"LOI Draft_{entity_name}_{date_str} {initials}.docx"
                    else:
                        filename = f"LOI Draft_{entity_name}_{date_str}.docx"

                    st.session_state.generated_file = result.getvalue()
                    st.session_state.generated_filename = filename
                    st.success(f"LOI generated: {filename}")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error generating LOI: {e}")


# ============================ RIGHT: PREVIEW ==============================
with preview_col:
    st.markdown("**Live Preview**")

    # --- Build tracked-change values ---
    pp_words_val = convert_to_words(int(purchase_price)) if purchase_price > 0 else ""
    pp_num_val = f"{int(purchase_price):,}" if purchase_price > 0 else ""
    init_dep_val = fmt_dollar(initial_deposit) if initial_deposit > 0 else ""
    add_dep_val = fmt_dollar(additional_deposit) if additional_deposit > 0 else ""
    monthly_val = fmt_dollar(monthly_release) if monthly_release > 0 else ""
    dd_period_val = fmt_period(dd_days) if dd_days > 0 else ""
    ga_period_val = fmt_period(ga_days) if ga_days > 0 else ""
    asm_period_val = fmt_period(assemblage_days) if assemblage_days > 0 else ""
    cl_period_val = fmt_period(closing_days) if closing_days > 0 else ""
    ext_months_val = fmt_period(closing_ext_months, "months") if closing_ext_months > 0 else ""
    ext_dep_val = fmt_dollar(ext_deposit) if ext_deposit > 0 else ""
    monthly_ext_val = fmt_dollar(monthly_closing_ext_deposit) if monthly_closing_ext_deposit > 0 else ""
    lr_val = fmt_dollar(legal_reimb_amount) if legal_reimb_amount > 0 else ""
    lt_days_val = fmt_period(lease_term_days) if lease_term_days > 0 else ""

    p = []  # preview parts
    p.append('<div class="doc-scroll">')

    # ==================== PAGE 1 ====================
    p.append('<div class="doc-page">')

    # -- Header --
    p.append('<div style="display:flex;justify-content:space-between;align-items:flex-start;">')
    p.append('<img src="https://subtextliving.com/wp-content/uploads/2023/08/subtext-primary-logo.svg" style="height:40px;">')
    p.append('<div style="text-align:right;font-size:9pt;">3000 Locust Street<br>St. Louis, MO 63103<br>Phone 314-721-5559 Fax 314-667-3121</div>')
    p.append('</div>')
    p.append('<hr>')

    # -- Date --
    p.append(f'<p>{_v(date_val, "[Date]")}</p>')

    # -- Address block --
    p.append(f'<p class="addr-line">{_v(seller_addr1, "[____________________]")}</p>')
    p.append(f'<p class="addr-line">{_v(seller_addr2, "[____________________]")}</p>')
    p.append(f'<p class="addr-line">{_v(seller_addr3, "[____________________]")}</p>')
    p.append(f'<p class="addr-line">Attn: {_v(attention_name, "[_______________]")}</p>')

    # -- Re: line --
    p.append(f'<p class="re-line"><b>Re: &nbsp;&nbsp;Proposal for the acquisition of {_v(property_address, "[Address, City, State]")} (&ldquo;Property&rdquo;)</b></p>')

    # -- Salutation --
    p.append(f'<p>{_v(salutation, "[Mr./Mrs./Ms._________]")}:</p>')

    # -- Opening paragraph --
    seller_display = _v(seller_name, "[______________]")
    prop_display = _v(property_address, "[Address, City, State]")
    p.append(
        f'<p style="text-indent:2em;">On behalf of Subtext Acquisitions, LLC, a Missouri limited liability company, or its assignee '
        f'(&ldquo;Purchaser&rdquo;), we are pleased to submit this non-binding proposal to purchase the above-referenced '
        f'Property from {seller_display} (&ldquo;Seller&rdquo;) on the terms and conditions set forth herein.</p>'
    )

    p.append('<p style="text-indent:2em;">The terms of the proposed sale are as follows:</p>')

    # -- A. Property --
    p.append(
        '<p class="section-item"><b>A.</b> &nbsp;&nbsp;&nbsp;<span class="sec-label">Property.</span> '
        'The proposal is for the Property described herein together with all improvements, rights of way, '
        'easements, hereditaments and appurtenances in any way related to or benefiting the Property. '
        'The Property is legally described and generally depicted on <u>Exhibit A</u> attached hereto.</p>'
    )

    # -- B. Purchase Price --
    pp_words_tc = _v(pp_words_val, "[_________________]") if pp_words_val else '<span class="tc-empty">[_________________]</span>'
    pp_num_tc = _v(pp_num_val, "[_________]") if pp_num_val else '<span class="tc-empty">[_________]</span>'
    p.append(
        f'<p class="section-item"><b>B.</b> &nbsp;&nbsp;&nbsp;<span class="sec-label">Purchase Price.</span> '
        f'{pp_words_tc} and 00/100 Dollars (${pp_num_tc}.00) '
        f'(the &ldquo;Purchase Price&rdquo;), which shall be paid at the Closing (defined below) by cash, '
        f'certified check or wire transfer.</p>'
    )

    # -- C. Deposit --
    init_tc = _v(init_dep_val, "[Ten Thousand and 00/100 Dollars ($10,000.00)]")
    add_tc = _v(add_dep_val, "[Ten Thousand and 00/100 Dollars ($10,000.00)]")
    monthly_tc = _v(monthly_val, "[Five Thousand and 00/100 Dollars ($5,000.00)]")

    if deposit_structure == DepositStructure.GOVERNMENTAL_APPROVALS_GOING_HARD:
        deposit_text = (
            f'Within five (5) business days following mutual execution of the Purchase Agreement '
            f'(as defined below), {init_tc} (&ldquo;Initial Deposit&rdquo;) shall be delivered to First American Title Insurance, '
            f'National Commercial Services (&ldquo;Title Company&rdquo;). If, prior to the expiration of the Due Diligence Period '
            f'(as defined below), Purchaser determines not to pursue the transaction, Purchaser may terminate the Purchase Agreement, '
            f'and the Initial Deposit shall be fully and promptly refunded to Purchaser. Otherwise, Purchaser shall deposit an additional '
            f'sum into escrow with the Title Company in the amount of {add_tc} (&ldquo;Additional Deposit&rdquo;, which together with '
            f'the Initial Deposit, is referred to herein as the (&ldquo;Earnest Money&rdquo;). The Earnest Money shall be non-refundable '
            f'to the Purchaser subject to (i) Purchaser&rsquo;s receipt of Governmental Approvals, during the Governmental Approvals Period '
            f'(as defined below), (ii) a default by Seller under the Purchase Agreement, or (iii) a casualty or a condemnation, each as '
            f'shall be further defined in the Purchase Agreement. If the Purchaser has not obtained the Governmental Approvals during the '
            f'Governmental Approvals Period, Purchaser may terminate the Purchase Agreement, and the Earnest Money shall be fully and promptly '
            f'refunded to Purchaser. Otherwise, the Earnest Money shall become non-refundable to the Purchaser upon expiration of the '
            f'Governmental Approvals Period (except as expressly set forth in the Purchase Agreement). The Earnest Money shall be applied '
            f'towards the Purchase Price at Closing.'
        )
    elif deposit_structure == DepositStructure.DUE_DILIGENCE_GOING_HARD:
        deposit_text = (
            f'Within five (5) business days following mutual execution of the Purchase Agreement '
            f'(as defined below), {init_tc} (&ldquo;Initial Deposit&rdquo;) shall be delivered to First American Title Insurance, '
            f'National Commercial Services (&ldquo;Title Company&rdquo;). If, prior to the expiration of the Due Diligence Period '
            f'(as defined below), Purchaser determines not to pursue the transaction, Purchaser may terminate the Purchase Agreement, '
            f'and the Initial Deposit shall be fully and promptly refunded to Purchaser. Otherwise, upon waiver of the Due Diligence Period, '
            f'the Initial Deposit shall become non-refundable to the Purchaser, and Purchaser shall deposit an additional sum into escrow '
            f'with the Title Company in the amount of {add_tc} (&ldquo;Additional Deposit&rdquo;, which together with the Initial Deposit, '
            f'is referred to herein as the (&ldquo;Earnest Money&rdquo;). The Additional Deposit shall be non-refundable to the Purchaser '
            f'subject to (i) Purchaser&rsquo;s receipt of the Governmental Approvals during the Governmental Approvals Period (as defined below), '
            f'(ii) a default by Seller under the Purchase Agreement, or (iii) a casualty or a condemnation, each as shall be further defined '
            f'in the Purchase Agreement. If Purchaser has not obtained the Governmental Approvals during the Governmental Approvals Period, then '
            f'Purchaser may terminate the Purchase Agreement, and the Additional Deposit shall be fully and promptly refunded to Purchaser. '
            f'Otherwise, the Additional Deposit shall become non-refundable to the Purchaser upon expiration of the Governmental Approvals Period '
            f'(except as expressly set forth in the Purchase Agreement). The Earnest Money shall be applied towards the Purchase Price at Closing.'
        )
    else:  # Monthly Going Hard
        deposit_text = (
            f'Within five (5) business days following mutual execution of the Purchase Agreement '
            f'(as defined below), {init_tc} (the &ldquo;Initial Deposit&rdquo;) shall be delivered to First American Title Insurance, '
            f'National Commercial Services (the &ldquo;Title Company&rdquo;). If, prior to the expiration of the Due Diligence Period '
            f'(as defined below), Purchaser determines not to pursue the transaction, Purchaser may terminate the Purchase Agreement and '
            f'the Initial Deposit shall be fully and promptly refunded to Purchaser. Otherwise, Purchaser shall deposit an additional sum '
            f'into escrow with the Title Company in the amount of {add_tc} (the &ldquo;Additional Deposit&rdquo;, which together with '
            f'the Initial Deposit, is referred to herein as the &ldquo;Earnest Money&rdquo;) and shall be non-refundable to the Purchaser, '
            f'subject to (i) Purchaser&rsquo;s receipt of the Governmental Approvals during the Governmental Approvals Period (as defined below), '
            f'(ii) a default by Seller under the Purchase Agreement, or (iii) a casualty or a condemnation, each as shall be further defined '
            f'in the Purchase Agreement. On the 1st of each month following waiver of the Due Diligence Period, {monthly_tc} of the Earnest Money '
            f'(collectively, the &ldquo;Monthly Releases&rdquo;), shall become non-refundable to the Purchaser, subject to a default by Seller '
            f'under the Purchase Agreement, and shall be immediately released to the Seller by Title Company. If Purchaser has not obtained the '
            f'Governmental Approvals during the Governmental Approvals Period, Purchaser may terminate the Purchase Agreement and the Earnest Money, '
            f'less the Monthly Releases paid to date, shall be fully and promptly refunded to Purchaser. Otherwise, the Earnest Money shall become '
            f'non-refundable to the Purchaser upon expiration of the Governmental Approvals Period (except as expressly set forth in the Purchase '
            f'Agreement). The Earnest Money shall be applied towards the Purchase Price at Closing.'
        )

    p.append(f'<p class="section-item"><b>C.</b> &nbsp;&nbsp;&nbsp;<span class="sec-label">Deposit.</span> {deposit_text}</p>')

    # -- Legal Reimbursement (optional) --
    if include_legal_reimb:
        lr_tc = _v(lr_val, "[Five Thousand and 00/100 Dollars ($5,000.00)]")
        p.append(
            f'<p style="text-indent:2em;"><b>Legal Reimbursement Fee.</b> Upon mutual execution of the Purchase Agreement, '
            f'{lr_tc} shall be immediately released to the Seller by the Title Company (&ldquo;Legal Reimbursement Fee&rdquo;) '
            f'and shall be non-refundable to the Purchaser, subject to a default by Seller under the Purchase Agreement, '
            f'a casualty or a condemnation.</p>'
        )

    # ==================== PAGE 2 ====================
    p.append('</div>')  # end page 1
    p.append('<div class="doc-page">')
    p.append(f'<p class="cont-header">{_vn(property_address, "[Address, City, State]")}<br>{_vn(date_val, "[Date]")}<br>Page 2</p>')

    # -- D. Exclusivity --
    p.append(
        '<p class="section-item"><b>D.</b> &nbsp;&nbsp;&nbsp;<span class="sec-label">Exclusivity; Seller&rsquo;s Covenants.</span> '
        'Upon mutual execution of this proposal, Purchaser shall have thirty (30) days (&ldquo;Exclusivity Period&rdquo;) to negotiate '
        'on an exclusive basis a definitive purchase agreement for the Property (&ldquo;Purchase Agreement&rdquo;), which Purchase Agreement '
        'shall be prepared by Purchaser and provided to Seller promptly after execution of this proposal. During the Exclusivity Period and '
        'while the Purchase Agreement is in effect, subject to Paragraph I, Seller shall (a) not negotiate with any other party or sell or '
        'lease, offer to sell or lease, accept an offer to purchase or lease or solicit or respond to solicitations for the sale or lease of '
        'any portion of or interest in the Property, (b) not execute any contracts, leases, easements or other documents or grant any rights to '
        'or affecting the Property or take any other material actions affecting the Property without Purchaser&rsquo;s prior written consent, and '
        '(c) continue to manage and maintain the Property as it is currently being managed and maintained.</p>'
    )

    # -- E. Due Diligence --
    dd_tc = _v(dd_period_val, "[one hundred twenty (120)]")
    ga_tc = _v(ga_period_val, "[one hundred fifty (150)]")

    if dd_type == DueDiligenceType.STANDARD:
        dd_text = (
            f'Purchaser shall have a period of {dd_tc} days after mutual execution of the Purchase Agreement '
            f'(&ldquo;Due Diligence Period&rdquo;) to inspect and investigate the Property to the extent deemed necessary or desirable '
            f'by Purchaser in its sole and absolute discretion. Upon expiration of the Due Diligence Period, Purchaser shall have a period of '
            f'{ga_tc} days (&ldquo;Governmental Approvals Period&rdquo;) to procure all zoning approvals, permits, consents, authorizations, '
            f'variances, waivers, licenses, certificates and other approvals from any governmental or quasi-governmental authority with respect '
            f'to the Property necessary or desirable, in Purchaser&rsquo;s sole and absolute discretion (collectively, &ldquo;Governmental Approvals&rdquo;).'
        )
        p.append(f'<p class="section-item"><b>E.</b> &nbsp;&nbsp;&nbsp;<span class="sec-label">Due Diligence and Governmental Approvals.</span> {dd_text}</p>')
    else:
        asm_tc = _v(asm_period_val, "[ninety (90)]")
        dd_text = (
            f'Purchaser shall have a period of {asm_tc} days after mutual execution of the Purchase Agreement '
            f'(&ldquo;Assemblage Period&rdquo;) to execute purchase agreements with all necessary adjacent parcels, in Purchaser&rsquo;s '
            f'sole and absolute discretion. Upon expiration of the Assemblage Period, Purchaser shall have a period of {dd_tc} days '
            f'(&ldquo;Due Diligence Period&rdquo;) to inspect and investigate the Property to the extent deemed necessary or desirable by '
            f'Purchaser in its sole and absolute discretion. Upon expiration of the Due Diligence Period, Purchaser shall have a period of '
            f'{ga_tc} days (&ldquo;Governmental Approvals Period&rdquo;) to procure all zoning approvals, permits, consents, authorizations, '
            f'variances, waivers, licenses, certificates and other approvals from any governmental or quasi-governmental authority with respect '
            f'to the Property necessary or desirable, in Purchaser&rsquo;s sole and absolute discretion (collectively, &ldquo;Governmental Approvals&rdquo;).'
        )
        p.append(f'<p class="section-item"><b>E.</b> &nbsp;&nbsp;&nbsp;<span class="sec-label">Assemblage; Due Diligence; Governmental Approvals.</span> {dd_text}</p>')

    # -- F. Closing --
    cl_tc = _v(cl_period_val, "[thirty (30)]")
    closing_text = f'Within {cl_tc} days after the expiration of the Governmental Approvals Period (&ldquo;Closing&rdquo;).'

    if include_closing_ext:
        ext_m_tc = _v(ext_months_val, "[six (6)]")
        ext_d_tc = _v(monthly_ext_val, "[Twenty-Five Thousand and 00/100 Dollars ($25,000.00)]")
        closing_text += (
            f' Notwithstanding the foregoing, Purchaser shall have the right to extend the Closing on a month-to-month basis '
            f'for up to a total of {ext_m_tc} months (each a &ldquo;Closing Extension&rdquo;) by delivering to the Title Company '
            f'a deposit in the amount of {ext_d_tc} for each month of extension (&ldquo;Monthly Closing Extension Deposit&rdquo;). '
            f'The Monthly Closing Extension Deposits shall be non-refundable to the Purchaser when made, subject to a default by '
            f'Seller under the Purchase Agreement, a casualty or a condemnation, but shall be applicable to the Purchase Price.'
        )

    p.append(f'<p class="section-item"><b>F.</b> &nbsp;&nbsp;&nbsp;<span class="sec-label">Closing.</span> {closing_text}</p>')

    # -- G. Commissions --
    if commission_type == CommissionType.SELLER_PAYS_LISTING_AGENT:
        comm_text = ('Seller to pay the commission set forth in the listing agreement at Closing, and Purchaser and Seller '
                     'agree no other brokerage commissions are due as a result of this transaction.')
    elif commission_type == CommissionType.SUBTEXT_PAYS:
        broker_tc = _v(broker_name, "[_____________]")
        comm_text = (f'Purchaser to pay a commission at Closing to {broker_tc} pursuant to a separate agreement, '
                     f'and Purchaser and Seller agree no other brokerage commissions are due as a result of this transaction.')
    else:
        comm_text = 'Purchaser and Seller agree that no brokerage commission is due as a result of this transaction.'

    p.append(f'<p class="section-item"><b>G.</b> &nbsp;&nbsp;&nbsp;<span class="sec-label">Commissions.</span> {comm_text}</p>')

    # -- H. Option to Extend (optional) --
    if include_option_extend:
        ext_dep_tc = _v(ext_dep_val, "[Five Thousand and 00/100 Dollars ($5,000.00)]")
        p.append(
            f'<p class="section-item"><b>H.</b> &nbsp;&nbsp;&nbsp;<span class="sec-label">Option to Extend.</span> '
            f'Purchaser shall have a total of two (2) sixty (60) day extension options (each an &ldquo;Option to Extend&rdquo;). '
            f'Each Option to Extend may be applied to the Due Diligence Period or the Governmental Approvals Period by delivering '
            f'written notice to the Seller prior to the expiration of such term (&ldquo;Extension Notice&rdquo;) and delivering to the '
            f'Title Company an amount equal to {ext_dep_tc} (&ldquo;Extension Deposit&rdquo;) within five (5) business days of the '
            f'Extension Notice. The Extension Deposits shall be non-refundable to the Purchaser when made, subject to a default by '
            f'Seller under the Purchase Agreement but shall be applicable to the Purchase Price.</p>'
        )

    # -- I. Leases --
    lease_sentences = []
    if include_existing_leases:
        le_tc = _v(lease_end_date, "[May 31, 2026]")
        lease_sentences.append(
            f'Purchaser and Seller agree there are existing leases on the Property and that all existing leases end on or before {le_tc}.'
        )
    if include_delivered_vacant:
        lease_sentences.append(
            'Seller agrees to terminate, effective as of Closing, all leases such that the Property shall be delivered vacant at Closing, '
            'and Seller shall be responsible for all monetary penalties/fees and or settlements related to the termination of leases.'
        )
    if include_lease_termination:
        lt_tc = _v(lt_days_val, "[sixty (60)]")
        lease_sentences.append(
            f'All new leases and renewals of existing leases entered into after the date hereof, shall include a {lt_tc} day '
            f'termination provision that shall be further defined in the Purchase Agreement.'
        )
    if include_negotiate_tenants:
        lease_sentences.append(
            'Upon mutual execution of this proposal, Purchaser shall have the right to negotiate directly with the existing tenants '
            'regarding a potential lease amendment, provided that Seller shall have the right to facilitate such introductions.'
        )

    if lease_sentences:
        p.append(f'<p class="section-item"><b>I.</b> &nbsp;&nbsp;&nbsp;<span class="sec-label">Leases.</span> {" ".join(lease_sentences)}</p>')

    # -- Seller Rollover (optional) --
    if include_seller_rollover:
        p.append(
            '<p class="section-item"><span class="sec-label">Seller Rollover Option for Project Equity.</span> '
            'In concert with the execution of the Purchase Agreement, Seller shall have the right, but not an obligation, to invest '
            'up to one hundred (100%) percent of his/her/its/their sale price as equity in the Purchaser&rsquo;s intended project '
            '(&ldquo;Project&rdquo;), with such investment being through a limited liability company or other entity type selected by '
            'the Purchaser in good faith, that will be an indirect owner of the Property, whereby Seller shall be a passive investor '
            '(no day to day or major decision rights) in the Project, but with Seller being entitled to substantially the same '
            '&ldquo;limited partner&rdquo; returns on equity as the to be selected capital partner for the Project will receive '
            '(&ldquo;Capital Partner&rdquo;).</p>'
        )

    # -- J. Confidentiality --
    p.append(
        '<p class="section-item"><b>J.</b> &nbsp;&nbsp;&nbsp;<span class="sec-label">Confidentiality.</span> '
        'During the Exclusivity Period and, while the Purchase Agreement is in effect, Purchaser and Seller shall keep all negotiations '
        'and communications between the parties regarding the potential purchase of the Property confidential and shall not disclose '
        'any matter related to such negotiations and communications to any third party.</p>'
    )

    # -- K. Governing Law --
    p.append(
        '<p class="section-item"><b>K.</b> &nbsp;&nbsp;&nbsp;<span class="sec-label">Governing Law.</span> '
        'This proposal shall be governed by the laws of the State of Missouri, without giving effect to conflict of laws principles.</p>'
    )

    # -- L. Miscellaneous --
    p.append(
        '<p class="section-item"><b>L.</b> &nbsp;&nbsp;&nbsp;<span class="sec-label">Miscellaneous.</span> '
        'This proposal may be signed in counterparts with the same effect as if executed on a single document. This proposal constitutes '
        'the entire agreement between the parties concerning the subject matter hereof and supersedes all prior representations, '
        'understandings or agreements, whether oral or written.</p>'
    )

    # ==================== PAGE 3 ====================
    p.append('</div>')  # end page 2
    p.append('<div class="doc-page">')
    p.append(f'<p class="cont-header">{_vn(property_address, "[Address, City, State]")}<br>{_vn(date_val, "[Date]")}<br>Page 3</p>')

    # -- Closing paragraph --
    p.append(
        '<p class="closing-text">If the foregoing is acceptable, please indicate by executing a copy of this proposal and returning '
        'it to the undersigned. We look forward to working with you on this transaction.</p>'
    )

    # -- Signature --
    p.append('<p class="sig-block">Sincerely,</p>')
    p.append('<p class="sig-line" style="margin-top:2em;">Subtext Acquisitions, LLC</p>')
    p.append('<p class="sig-line" style="margin-top:2em;">____________________________________</p>')
    p.append('<p class="sig-line">Richard Birner, Vice President of Land Acquisitions</p>')

    # -- Seller signature --
    p.append('<p class="sig-block" style="margin-top:1.5em;">Agreed and Accepted by Seller this _____ day of ___________, 2026</p>')

    if sig_type == SignatureBlockType.INDIVIDUAL:
        p.append('<p class="sig-line" style="margin-top:2em;">____________________________________</p>')
        seller_sig_tc = _vn(seller_name_sig, "[SELLER NAME]")
        p.append(f'<p class="sig-line">{seller_sig_tc}</p>')
    else:
        for ent in st.session_state.entities:
            name = ent["company_name"]
            company_tc = _vn(name, "[COMPANY NAME]")
            p.append(f'<p class="sig-line" style="margin-top:1.5em;"><b>{company_tc}</b></p>')
            p.append('<p class="sig-line">By: &nbsp;&nbsp;________________________</p>')
            p.append('<p class="sig-line">Name: &nbsp;&nbsp;________________________</p>')
            p.append('<p class="sig-line">Title: &nbsp;&nbsp;________________________</p>')

    # ==================== PAGE 4: EXHIBIT A ====================
    p.append('</div>')  # end page 3
    p.append('<div class="doc-page">')
    p.append(f'<p class="cont-header">{_vn(property_address, "[Address, City, State]")}<br>{_vn(date_val, "[Date]")}<br>Page 4</p>')

    p.append('<p class="exhibit-header">EXHIBIT A</p>')
    p.append('<p class="exhibit-center" style="margin-top:2em;">DEPICTION OF PROPERTY</p>')
    p.append('<p class="exhibit-center" style="margin-top:1em;">PARCEL ID NUMBER(S)</p>')
    parcel_list = "<br>".join(pid for pid in st.session_state.parcel_ids if pid.strip())
    if parcel_list:
        p.append(f'<p class="exhibit-center" style="margin-top:0.5em;">{parcel_list}</p>')
    else:
        p.append('<p class="exhibit-center" style="margin-top:0.5em;"><span class="tc-empty">[]</span></p>')

    # -- Property photo --
    if uploaded_photo:
        photo_bytes = uploaded_photo.read()
        uploaded_photo.seek(0)  # reset for document generation
        b64 = base64.b64encode(photo_bytes).decode()
        mime = uploaded_photo.type or "image/jpeg"
        p.append(f'<img class="photo-preview" src="data:{mime};base64,{b64}" alt="Property Photo">')

    p.append('</div>')  # end page 4
    p.append('</div>')  # end doc-scroll

    st.markdown("\n".join(p), unsafe_allow_html=True)

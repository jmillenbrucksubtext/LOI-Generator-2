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


def dollar_preview(amount: float):
    if amount > 0:
        st.markdown(f'<div class="legal-preview">{fmt_dollar(amount)}</div>', unsafe_allow_html=True)


def period_preview(value: int, unit: str = "days"):
    if value > 0:
        st.markdown(f'<div class="legal-preview">{fmt_period(value, unit)}</div>', unsafe_allow_html=True)


def _v(text: str, fallback: str = "___") -> str:
    """Return text or a placeholder for the preview."""
    return text.strip() if text.strip() else f'<span style="color:#c1d100">{fallback}</span>'


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

    /* Document preview styling */
    .doc-preview {
        background: #fff;
        color: #222;
        font-family: 'Times New Roman', Times, serif;
        font-size: 11pt;
        line-height: 1.5;
        padding: 2rem 2.5rem;
        border-radius: 4px;
        max-height: 85vh;
        overflow-y: auto;
        border: 1px solid #555;
    }
    .doc-preview p {
        margin: 0.6em 0;
        text-indent: 2em;
    }
    .doc-preview p.no-indent {
        text-indent: 0;
    }
    .doc-preview .doc-header {
        text-align: left;
        text-indent: 0;
        margin-bottom: 0.3em;
    }
    .doc-preview .doc-section {
        font-weight: bold;
        text-indent: 0;
        margin-top: 1em;
    }
    .doc-preview .doc-re {
        text-indent: 0;
        margin: 0.8em 0;
    }
    .doc-preview .doc-sig {
        text-indent: 0;
        margin-top: 2em;
    }
    .doc-preview .highlight {
        background: #fffde0;
        padding: 0 2px;
        border-radius: 2px;
    }
    .doc-preview hr {
        border: none;
        border-top: 1px solid #ccc;
        margin: 1em 0;
    }
    /* Make preview column sticky */
    [data-testid="stVerticalBlock"] > div:has(.doc-preview) {
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
        date_val = st.text_input("Date", value=datetime.now().strftime("%B %d, %Y"))
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
    purchase_price = st.number_input("Purchase Price ($)", min_value=0.0, value=0.0, step=1000.0, format="%.2f")
    dollar_preview(purchase_price)

    # --- C. Deposit ---
    st.markdown('<div class="section-header">C. Deposit</div>', unsafe_allow_html=True)
    deposit_options = [e.value for e in DepositStructure]
    deposit_choice = st.radio("Deposit Scenario", deposit_options, index=0, horizontal=True)
    deposit_structure = DepositStructure(deposit_choice)

    cd1, cd2 = st.columns(2)
    with cd1:
        initial_deposit = st.number_input("Initial Deposit ($)", min_value=0.0, value=10000.0, step=1000.0, format="%.2f")
        dollar_preview(initial_deposit)
    with cd2:
        if deposit_structure in (DepositStructure.GOVERNMENTAL_APPROVALS_GOING_HARD, DepositStructure.DUE_DILIGENCE_GOING_HARD):
            additional_deposit = st.number_input("Additional Deposit ($)", min_value=0.0, value=10000.0, step=1000.0, format="%.2f")
            dollar_preview(additional_deposit)
        else:
            additional_deposit = 10000.0
        if deposit_structure == DepositStructure.MONTHLY_GOING_HARD:
            monthly_release = st.number_input("Monthly Release Amount ($)", min_value=0.0, value=5000.0, step=1000.0, format="%.2f")
            dollar_preview(monthly_release)
        else:
            monthly_release = 5000.0

    st.divider()
    include_legal_reimb = st.checkbox("Include Legal Reimbursement Fee")
    if include_legal_reimb:
        legal_reimb_amount = st.number_input("Legal Reimbursement Amount ($)", min_value=0.0, value=5000.0, step=1000.0, format="%.2f")
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
            monthly_closing_ext_deposit = st.number_input("Monthly Closing Extension Deposit ($)", min_value=0.0, value=25000.0, step=1000.0, format="%.2f")
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
        ext_deposit = st.number_input("Extension Deposit Amount ($)", min_value=0.0, value=5000.0, step=1000.0, format="%.2f")
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

    # Build preview values
    pp_str = fmt_dollar(purchase_price) if purchase_price > 0 else "___"
    pp_words = convert_to_words(int(purchase_price)) if purchase_price > 0 else "___"
    pp_num = f"${int(purchase_price):,}" if purchase_price > 0 else "$___"
    init_dep_str = fmt_dollar(initial_deposit) if initial_deposit > 0 else "___"
    add_dep_str = fmt_dollar(additional_deposit) if additional_deposit > 0 else "___"
    monthly_str = fmt_dollar(monthly_release) if monthly_release > 0 else "___"

    # Deposit paragraph based on scenario
    if deposit_structure == DepositStructure.GOVERNMENTAL_APPROVALS_GOING_HARD:
        deposit_para = (
            f'<b>C. Deposit.</b> Within five (5) business days after mutual execution of the PSA, '
            f'Purchaser shall deposit {_v(init_dep_str)} as an initial deposit. '
            f'An additional {_v(add_dep_str)} shall become non-refundable upon expiration of '
            f'the Governmental Approvals Period.'
        )
    elif deposit_structure == DepositStructure.DUE_DILIGENCE_GOING_HARD:
        deposit_para = (
            f'<b>C. Deposit.</b> Within five (5) business days after mutual execution of the PSA, '
            f'Purchaser shall deposit {_v(init_dep_str)} as an initial deposit. '
            f'An additional {_v(add_dep_str)} shall become non-refundable upon expiration of '
            f'the Due Diligence Period.'
        )
    else:
        deposit_para = (
            f'<b>C. Deposit.</b> Within five (5) business days after mutual execution of the PSA, '
            f'Purchaser shall deposit {_v(init_dep_str)} as an initial deposit. '
            f'Thereafter, {_v(monthly_str)} shall be released monthly after Due Diligence.'
        )

    # Legal reimbursement
    legal_reimb_para = ""
    if include_legal_reimb:
        lr_str = fmt_dollar(legal_reimb_amount) if legal_reimb_amount > 0 else "___"
        legal_reimb_para = f'<p>Additionally, Purchaser shall pay a legal reimbursement fee of {_v(lr_str)} at PSA execution.</p>'

    # Due Diligence paragraph
    dd_period = fmt_period(dd_days) if dd_days > 0 else "___"
    ga_period = fmt_period(ga_days) if ga_days > 0 else "___"
    if dd_type == DueDiligenceType.STANDARD:
        dd_para = (
            f'<b>E. Due Diligence and Governmental Approvals.</b> Purchaser shall have a due diligence period of '
            f'{_v(dd_period)} from the Effective Date. The Governmental Approvals Period shall be '
            f'{_v(ga_period)} from the Effective Date.'
        )
    else:
        asm_period = fmt_period(assemblage_days) if assemblage_days > 0 else "___"
        dd_para = (
            f'<b>E. Due Diligence and Governmental Approvals.</b> Purchaser shall have an assemblage period of '
            f'{_v(asm_period)} from the Effective Date, followed by a due diligence period of '
            f'{_v(dd_period)}. The Governmental Approvals Period shall be '
            f'{_v(ga_period)} from the Effective Date.'
        )

    # Closing paragraph
    cl_period = fmt_period(closing_days) if closing_days > 0 else "___"
    closing_para = f'<b>F. Closing.</b> Closing shall occur within {_v(cl_period)} after expiration of the Governmental Approvals Period.'
    if include_closing_ext:
        ext_months = fmt_period(closing_ext_months, "months") if closing_ext_months > 0 else "___"
        ext_dep_str = fmt_dollar(monthly_closing_ext_deposit) if monthly_closing_ext_deposit > 0 else "___"
        closing_para += (
            f' Purchaser shall have the option to extend closing for {_v(ext_months)} '
            f'with a monthly deposit of {_v(ext_dep_str)}.'
        )

    # Commission paragraph
    if commission_type == CommissionType.SELLER_PAYS_LISTING_AGENT:
        comm_para = '<b>G. Commissions.</b> Seller shall be responsible for paying any commission owed to the listing agent.'
    elif commission_type == CommissionType.SUBTEXT_PAYS:
        comm_para = f'<b>G. Commissions.</b> Purchaser shall pay the commission to {_v(broker_name)}.'
    else:
        comm_para = '<b>G. Commissions.</b> No brokers are involved in this transaction.'

    # Option to extend
    option_para = ""
    if include_option_extend:
        ext_d_str = fmt_dollar(ext_deposit) if ext_deposit > 0 else "___"
        option_para = (
            f'<p><b>H. Option to Extend.</b> Purchaser shall have the option to extend the Governmental Approvals Period '
            f'by depositing an additional {_v(ext_d_str)}.</p>'
        )

    # Leases paragraph
    lease_parts = []
    if include_existing_leases:
        lease_parts.append(f'Purchaser acknowledges existing leases with an end date of {_v(lease_end_date)}.')
    if include_delivered_vacant:
        lease_parts.append('The Property shall be delivered vacant at closing.')
    if include_lease_termination:
        lt_days = fmt_period(lease_term_days) if lease_term_days > 0 else "___"
        lease_parts.append(f'Seller shall terminate all leases within {_v(lt_days)} prior to closing.')
    if include_negotiate_tenants:
        lease_parts.append('Purchaser shall have the right to negotiate with existing tenants prior to closing.')
    lease_para = f'<b>I. Leases.</b> {" ".join(lease_parts)}' if lease_parts else ""

    # Seller rollover
    rollover_para = ""
    if include_seller_rollover:
        rollover_para = '<p>Seller shall have the option to contribute land and receive LP level returns in the project.</p>'

    # Signature block
    if sig_type == SignatureBlockType.INDIVIDUAL:
        sig_block = f"""
        <p class="no-indent" style="margin-top:2em;">Agreed and Accepted:</p>
        <p class="no-indent" style="margin-top:1.5em;">____________________________<br>{_v(seller_name_sig, "Seller Name")}</p>
        """
    else:
        entities_html = ""
        for ent in st.session_state.entities:
            name = ent["company_name"] or "___"
            entities_html += f'<p class="no-indent" style="margin-top:1.5em;"><b>{name}</b><br>By: ____________________________<br>Name: ____________________________<br>Title: ____________________________</p>'
        sig_block = f'<p class="no-indent" style="margin-top:2em;">Agreed and Accepted:</p>{entities_html}'

    # Parcel IDs
    parcel_list = "<br>".join(p for p in st.session_state.parcel_ids if p.strip()) or "___"

    # Assemble the preview
    preview_html = f"""
    <div class="doc-preview">
        <p class="doc-header">{_v(date_val, "Date")}</p>
        <p class="doc-header">{_v(seller_addr1, "Seller Entity")}</p>
        <p class="doc-header">{_v(seller_addr2, "Seller Address")}</p>
        <p class="doc-header">{_v(seller_addr3, "City, State Zip")}</p>
        <p class="doc-header" style="margin-top:0.5em;">Attention: {_v(attention_name, "Name")}</p>

        <p class="doc-re"><b>Re: Letter of Intent — {_v(property_address, "Property Address")}</b></p>

        <p>Dear {_v(salutation, "Mr./Mrs./Ms.")}:</p>

        <p>This Letter of Intent sets forth the basic terms pursuant to which Subtext LLC, or its assignee
        (&ldquo;Purchaser&rdquo;), would be willing to enter into a Purchase and Sale Agreement (&ldquo;PSA&rdquo;) with
        {_v(seller_name, "Seller Name")} (&ldquo;Seller&rdquo;) for the purchase of the property located at
        {_v(property_address, "Property Address")} (the &ldquo;Property&rdquo;).</p>

        <p class="doc-section">Key Terms:</p>

        <p><b>B. Purchase Price.</b> The purchase price shall be {_v(pp_words)} and 00/100 Dollars ({_v(pp_num)}).</p>

        <p>{deposit_para}</p>
        {legal_reimb_para}

        <p>{dd_para}</p>

        <p>{closing_para}</p>

        <p>{comm_para}</p>

        {option_para}

        {"<p>" + lease_para + "</p>" if lease_para else ""}

        {rollover_para}

        <hr>

        <p class="doc-section">EXHIBIT A</p>
        <p class="no-indent"><b>Parcel ID(s):</b><br>{parcel_list}</p>
        {"<p class='no-indent'><em>[ Property Photo Attached ]</em></p>" if uploaded_photo else ""}

        {sig_block}
    </div>
    """

    st.markdown(preview_html, unsafe_allow_html=True)

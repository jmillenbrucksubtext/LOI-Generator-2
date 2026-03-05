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
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(page_title="LOI Generator", layout="centered")


# ---------------------------------------------------------------------------
# Helper: format period preview (matches document output)
# ---------------------------------------------------------------------------
def format_period_preview(value: int, unit: str = "days") -> str:
    """Preview how a number will appear in the document, e.g. 'one hundred twenty (120) days'."""
    words = convert_to_words(value).lower()
    return f"{words} ({value}) {unit}"


def dollar_preview(amount: float):
    """Render a dollar amount preview below an input."""
    if amount > 0:
        st.markdown(
            f'<div class="legal-preview">{to_legal_dollar_string(amount)}</div>',
            unsafe_allow_html=True,
        )


def period_preview(value: int, unit: str = "days"):
    """Render a period preview below an input."""
    if value > 0:
        st.markdown(
            f'<div class="legal-preview">{format_period_preview(value, unit)}</div>',
            unsafe_allow_html=True,
        )


# ---------------------------------------------------------------------------
# Custom CSS for Subtext branding
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
    /* Tighter spacing on dynamic list buttons */
    .stButton > button[kind="secondary"] {
        padding: 0.25rem 0.6rem;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.image("https://subtextliving.com/wp-content/uploads/2023/08/subtext-primary-logo.svg", width=220)
st.markdown("#### LOI Generator")

# ---------------------------------------------------------------------------
# Session state initialization for dynamic lists
# ---------------------------------------------------------------------------
if "parcel_ids" not in st.session_state:
    st.session_state.parcel_ids = [""]
if "entities" not in st.session_state:
    st.session_state.entities = [{"company_name": ""}]
if "generated_file" not in st.session_state:
    st.session_state.generated_file = None
    st.session_state.generated_filename = None

# ---------------------------------------------------------------------------
# Letter Details
# ---------------------------------------------------------------------------
st.markdown('<div class="section-header">Letter Details</div>', unsafe_allow_html=True)

col1, col2 = st.columns([1, 2])
with col1:
    date_val = st.text_input("Date", value=datetime.now().strftime("%B %d, %Y"))
with col2:
    property_address = st.text_input("Property Address", placeholder="Address, City, State Zip Code")

seller_addr1 = st.text_input("Seller Address Line 1 (Seller Entity / Individual)")
seller_addr2 = st.text_input("Seller Address Line 2 (Owner Address)")
seller_addr3 = st.text_input("Seller Address Line 3 (City, State Zip Code)")

col_a, col_b, col_c = st.columns(3)
with col_a:
    attention_name = st.text_input("Attention Name", placeholder="John Smith")
with col_b:
    salutation = st.text_input("Salutation", placeholder="Mr. Smith")
with col_c:
    seller_name = st.text_input("Seller Name", placeholder="ABC Properties LLC")

# ---------------------------------------------------------------------------
# B. Purchase Price
# ---------------------------------------------------------------------------
st.markdown('<div class="section-header">B. Purchase Price</div>', unsafe_allow_html=True)

purchase_price = st.number_input("Purchase Price ($)", min_value=0.0, value=0.0, step=1000.0, format="%.2f")
dollar_preview(purchase_price)

# ---------------------------------------------------------------------------
# C. Deposit
# ---------------------------------------------------------------------------
st.markdown('<div class="section-header">C. Deposit</div>', unsafe_allow_html=True)

deposit_options = [e.value for e in DepositStructure]
deposit_choice = st.radio("Deposit Scenario", deposit_options, index=0, horizontal=True)
deposit_structure = DepositStructure(deposit_choice)

col_d1, col_d2 = st.columns(2)
with col_d1:
    initial_deposit = st.number_input("Initial Deposit ($)", min_value=0.0, value=10000.0, step=1000.0, format="%.2f")
    dollar_preview(initial_deposit)

with col_d2:
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

# ---------------------------------------------------------------------------
# E. Due Diligence and Governmental Approvals
# ---------------------------------------------------------------------------
st.markdown('<div class="section-header">E. Due Diligence and Governmental Approvals</div>', unsafe_allow_html=True)

dd_options = [e.value for e in DueDiligenceType]
dd_choice = st.radio("Due Diligence Type", dd_options, index=0, horizontal=True)
dd_type = DueDiligenceType(dd_choice)

col_e1, col_e2 = st.columns(2)
with col_e1:
    dd_days = st.number_input("Due Diligence Period (days)", min_value=0, value=120, step=1)
    period_preview(dd_days)
with col_e2:
    ga_days = st.number_input("Governmental Approvals Period (days)", min_value=0, value=150, step=1)
    period_preview(ga_days)

if dd_type == DueDiligenceType.WITH_ASSEMBLAGE:
    assemblage_days = st.number_input("Assemblage Period (days)", min_value=0, value=90, step=1)
    period_preview(assemblage_days)
else:
    assemblage_days = 90

# ---------------------------------------------------------------------------
# F. Closing
# ---------------------------------------------------------------------------
st.markdown('<div class="section-header">F. Closing</div>', unsafe_allow_html=True)

closing_days = st.number_input("Closing Period (days)", min_value=0, value=30, step=1)
period_preview(closing_days)

include_closing_ext = st.checkbox("Include Closing Extension")
if include_closing_ext:
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        closing_ext_months = st.number_input("Closing Extension Duration (months)", min_value=0, value=6, step=1)
        period_preview(closing_ext_months, "months")
    with col_f2:
        monthly_closing_ext_deposit = st.number_input("Monthly Closing Extension Deposit ($)", min_value=0.0, value=25000.0, step=1000.0, format="%.2f")
        dollar_preview(monthly_closing_ext_deposit)
else:
    closing_ext_months = 6
    monthly_closing_ext_deposit = 25000.0

# ---------------------------------------------------------------------------
# G. Commissions
# ---------------------------------------------------------------------------
st.markdown('<div class="section-header">G. Commissions</div>', unsafe_allow_html=True)

comm_options = [e.value for e in CommissionType]
comm_choice = st.radio("Commission Type", comm_options, index=0, horizontal=True)
commission_type = CommissionType(comm_choice)

if commission_type == CommissionType.SUBTEXT_PAYS:
    broker_name = st.text_input("Broker Name", placeholder="Jones Realty")
else:
    broker_name = ""

# ---------------------------------------------------------------------------
# H. Option to Extend
# ---------------------------------------------------------------------------
st.markdown('<div class="section-header">H. Option to Extend</div>', unsafe_allow_html=True)

include_option_extend = st.checkbox("Include Option to Extend", value=True)
if include_option_extend:
    ext_deposit = st.number_input("Extension Deposit Amount ($)", min_value=0.0, value=5000.0, step=1000.0, format="%.2f")
    dollar_preview(ext_deposit)
else:
    ext_deposit = 5000.0

# ---------------------------------------------------------------------------
# I. Leases
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# Signature
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# Exhibit A
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# Prepared By
# ---------------------------------------------------------------------------
st.markdown('<div class="section-header">Prepared By</div>', unsafe_allow_html=True)

col_pb1, col_pb2 = st.columns(2)
with col_pb1:
    prep_first = st.text_input("First Name", placeholder="First name", key="prep_first")
with col_pb2:
    prep_last = st.text_input("Last Name", placeholder="Last name", key="prep_last")

st.caption("This name will appear as the author of all tracked changes in the document.")

# ---------------------------------------------------------------------------
# Generate
# ---------------------------------------------------------------------------
st.divider()

# Show download button if a file was already generated this session
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
    # Validation
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
        # Build form data
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

        # Find template
        script_dir = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(script_dir, "Templates", "LOI_Template.docx")

        if not os.path.exists(template_path):
            st.error(f"Template not found at: {template_path}")
        else:
            try:
                with st.spinner("Generating LOI..."):
                    generator = DocumentGenerator()
                    result = generator.generate(template_path, form)

                # Build filename
                entity_name = seller_name.strip() if seller_name.strip() else "Unknown"
                f_initial = prep_first.strip()[0].upper() if prep_first.strip() else ""
                l_initial = prep_last.strip()[0].upper() if prep_last.strip() else ""
                initials = f_initial + l_initial
                date_str = datetime.now().strftime("%Y%m%d")
                if initials:
                    filename = f"LOI Draft_{entity_name}_{date_str} {initials}.docx"
                else:
                    filename = f"LOI Draft_{entity_name}_{date_str}.docx"

                # Store in session state so download button persists
                st.session_state.generated_file = result.getvalue()
                st.session_state.generated_filename = filename

                st.success(f"LOI generated: {filename}")
                st.rerun()
            except Exception as e:
                st.error(f"Error generating LOI: {e}")

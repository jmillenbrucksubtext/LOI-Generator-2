from dataclasses import dataclass, field
from enum import Enum
from typing import Optional


class DepositStructure(Enum):
    GOVERNMENTAL_APPROVALS_GOING_HARD = "Governmental Approvals Going Hard"
    DUE_DILIGENCE_GOING_HARD = "Due Diligence Going Hard"
    MONTHLY_GOING_HARD = "Monthly Going Hard"


class DueDiligenceType(Enum):
    STANDARD = "Standard"
    WITH_ASSEMBLAGE = "With Assemblage Period"


class CommissionType(Enum):
    SELLER_PAYS_LISTING_AGENT = "Seller Pays Listing Agent"
    SUBTEXT_PAYS = "Subtext Pays Commission"
    NO_BROKERS = "No Brokers Involved"


class ClosingExtensionType(Enum):
    NONE = "No Closing Extension"
    MONTH_TO_MONTH = "Month-to-Month Extensions"
    SINGLE = "One Closing Extension"


class SignatureBlockType(Enum):
    INDIVIDUAL = "Individual Seller"
    COMPANY_ENTITY = "Company / Entity"


@dataclass
class PropertyPhoto:
    photo_bytes: bytes = b""
    content_type: str = "image/jpeg"
    filename: str = "photo.jpg"


@dataclass
class SignatureEntity:
    company_name: str = ""


@dataclass
class LoiFormData:
    # Party Information
    date: str = ""
    seller_address_line1: str = ""
    seller_address_line2: str = ""
    seller_address_line3: str = ""
    attention_name: str = ""
    property_address: str = ""
    header_address: str = ""  # If set, used in page headers instead of property_address
    salutation: str = ""
    seller_name: str = ""

    # Financial Terms
    purchase_price: Optional[float] = None
    initial_deposit: Optional[float] = 10000
    additional_deposit: Optional[float] = 10000
    monthly_release_amount: Optional[float] = 5000
    legal_reimbursement_amount: Optional[float] = 5000
    extension_deposit_amount: Optional[float] = 5000
    monthly_closing_extension_deposit: Optional[float] = 25000

    # Timeline
    due_diligence_days: Optional[int] = 120
    governmental_approvals_days: Optional[int] = 150
    assemblage_days: Optional[int] = 90
    closing_days: Optional[int] = 30
    closing_extension_months: Optional[int] = 6

    # Lease Terms
    lease_end_date: str = "May 31, 2026"
    lease_termination_days: Optional[int] = 60

    # Commission
    broker_name: str = ""

    # Signature
    seller_name_signature: str = ""
    signature_entities: list = field(default_factory=lambda: [SignatureEntity()])
    parcel_ids: list = field(default_factory=lambda: [""])

    # Property Photos
    property_photos: list = field(default_factory=list)  # list[PropertyPhoto]
    # Legacy single-photo fields (still supported for backwards compatibility)
    property_photo_bytes: Optional[bytes] = None
    property_photo_content_type: Optional[str] = None
    property_photo_filename: Optional[str] = None

    # Scenario Options
    deposit_structure: DepositStructure = DepositStructure.GOVERNMENTAL_APPROVALS_GOING_HARD
    include_legal_reimbursement: bool = False
    due_diligence_type: DueDiligenceType = DueDiligenceType.STANDARD
    closing_extension_type: ClosingExtensionType = ClosingExtensionType.NONE
    commission_type: CommissionType = CommissionType.SELLER_PAYS_LISTING_AGENT
    include_option_to_extend: bool = True
    num_extension_options: Optional[int] = 2
    extension_option_days: Optional[int] = 60
    include_existing_leases: bool = True
    include_delivered_vacant: bool = False
    include_lease_termination: bool = False
    include_right_to_negotiate_with_tenants: bool = False
    include_seller_rollover: bool = False
    signature_block_type: SignatureBlockType = SignatureBlockType.INDIVIDUAL

    # Prepared By
    prepared_by_first_name: str = ""
    prepared_by_last_name: str = ""

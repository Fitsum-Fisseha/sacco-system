import streamlit as st
import pandas as pd
import os
import shutil
import base64
import math
from io import StringIO, BytesIO
from datetime import datetime
import hashlib
import secrets

# ============================================================
# CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="ተስፋ የገንዘብ ቁጠባና ብድር ማህበር",
    page_icon="🏦",
    layout="wide"
)

DB_FILE = "sacco_database.xlsx"

# 12 months - Ethiopian fiscal/year cycle requested
MONTHS = [
    "መስከረም",
    "ጥቅምት",
    "ኅዳር",
    "ታኅሣሥ",
    "ጥር",
    "የካቲት",
    "መጋቢት",
    "ሚያዝያ",
    "ግንቦት",
    "ሰኔ",
    "ሐምሌ",
    "ነሐሴ"
]

SEED_VERSION = "phase3-auth-v1"

REGISTRATION_FEE = 500.0
LOAN_FEE_RATE = 0.10
INTEREST_RATE = 0.02

LOAN_TERMS = [3, 6, 12, 24, 36]

# ============================================================
# PHASE 3 - USERS, ROLES & AUDIT
# ============================================================

DEFAULT_ADMIN_USERNAME = "admin"
DEFAULT_ADMIN_PASSWORD = "admin123"

ROLE_PERMISSIONS = {
    "Admin": {
        "🏠 ዋና ዳሽቦርድ",
        "👤 አባል መመዝገቢያ",
        "💰 የወር ቁጠባ ማስገቢያ",
        "💵 የብድር አገልግሎት",
        "📅 የብድር ክፍያ መመዝገቢያ",
        "📊 ጠቅላላ ሪፖርት",
        "✏️ የአባላት መረጃ ማስተካከያ",
        "👥 ተጠቃሚዎች አስተዳደር",
    },
    "Savings": {
        "🏠 ዋና ዳሽቦርድ",
        "💰 የወር ቁጠባ ማስገቢያ",
        "📊 ጠቅላላ ሪፖርት",
    },
    "Loan": {
        "🏠 ዋና ዳሽቦርድ",
        "💵 የብድር አገልግሎት",
        "📅 የብድር ክፍያ መመዝገቢያ",
        "📊 ጠቅላላ ሪፖርት",
    },
    "Reports": {
        "🏠 ዋና ዳሽቦርድ",
        "📊 ጠቅላላ ሪፖርት",
    },
}

ALL_ROLES = list(ROLE_PERMISSIONS.keys())


# ============================================================
# 136 MEMBERS
# ============================================================

MEMBER_DATA = """
1	ወርቅነህ ጉታ	1000	1000	1000
2	ኢሳያስ አይሳ	0	0	0
3	ፍቅራለም ዳኪቶ	1500	1500	1500
4	ትክክል ከበደ	2000	2000	2000
5	መስታውት ወንድሙ	2000	2000	2000
6	አምባው መዳረሻ	1500	1500	1500
7	አቻምየለሽ ገሪቶ	8000	8000	8000
8	መስከረም አያሌው	1000	1000	1000
9	አዲሳለም ገዛኸኝ	1000	1000	1000
10	አሰቴር ጋሪፎ	500	500	500
11	ወልዴ ታደሰ	1000	2250	1000
12	መስፍን ኃይሌ	500	500	500
13	አድማሱ ሙላቱ	500	500	500
14	ዋሲሁን አድነው	1000	1000	1000
15	ስኳሬ ማሞ	1100	1100	1100
16	ተሸለ አንጉሎ	13000	13000	13000
17	ታደለች ሰይድ	700	700	700
18	አበበ ዳጫቸው	1000	1000	1000
19	አዲሱ አቡዬ	1000	1000	1000
20	ገለታ ጌታቸው	6000	6000	6000
21	ጌታቸው ታደሰ	500	500	500
22	ቡዛለም ጸሀይ	500	500	500
23	ሽብሩ ዳዲሞ	500	6000	500
24	ይልቃል ካሳ	1000	1000	1000
25	አበበች ደሳለኝ	500	500	500
26	ሙሉቀን ታዬ	6000	6000	6000
27	ግርማ ገ/ሚካኤል	1000	1000	1000
28	ሥጋቱ አሰፋ	1000	1000	1000
29	ልኡል ሰገድ ሙንዬ	0	0	0
30	ሀብታሙ ኃይሌ	500	500	500
31	ግዛው አርሰኖ	300	300	300
32	ክፍሌ አበበ	1000	1000	1000
33	እምሩ እሚቶ	500	500	500
34	መከተ ማሞ	1000	1000	1000
35	ጀመረ ቆጭቶ	500	500	500
36	ዳግም ደምሴ	2000	2000	2000
37	ገነት ኃይሌ	500	500	500
38	አዲሱ አገሎ	1000	1000	1000
39	አክሊሉ አቾሞ	500	0	0
40	ቆጭቶ ወ/ሥላሴ	500	500	500
41	ትግሌ ታምሩ	500	1000	1000
42	ታሪኳ አንገሎ	600	600	600
43	ዮሐንስ ታከለ	1000	1000	1000
44	ታምሩ ጋሎ	500	500	500
45	ታምሩ ደስኖ	5000	5000	5000
46	ኬሮ ማሞ	700	700	0
47	አክሊሉ ሻረው	500	500	500
48	አጦ አምቦ	500	500	500
49	አክሊሉ ገበዬሁ	500	500	500
50	መሠረት መቹሎ	1000	1000	1000
51	ማዘንጊያሽ በፍቃዱ	500	500	500
52	ሽመልስ ይመር	500	500	500
53	ሙሉጌታ ሃይሌ	0	0	0
54	መላኩ ወ/ሚካኤል	1000	1000	1000
55	አባተ ገብሬ	0	0	0
56	አክሊሉ ኃይሌ	500	500	500
57	አስረስ አደም	1000	1000	1000
58	የሺዋስ ሙላቱ	500	500	500
59	ግዛው ደንበል	0	0	0
60	ኪሮስ ምትኩ	10000	10000	2000
61	ዘሪቱ ይማም	3000	3000	3000
62	ኤሊያስ አዳሾ	500	500	500
63	ትዕግስት ዓለሙ	1000	1000	1000
64	አዲሱ ደመቀ	1000	1000	1000
65	አበበ ቦጋለ	1000	1000	1000
66	እሸቱ በዛብህ	500	500	500
67	ዳርጌ እሸቱ	500	500	500
68	ቡዛዬሁ ኃይሌ	500	500	500
69	ካሳሁን ወ/ጻድቅ	1000	1000	1000
70	አልማዝ ኃይሌ	0	0	0
71	ለገሰ ዘውዴ	200	200	200
72	እህታለም ታዬ	2000	2000	2000
73	አያሌው ከልክሌ	0	0	0
74	ሀብታሙ አሰፋ	0	0	0
75	ዘሪሁን ዲላሞ	300	300	300
76	ባህሩ ገባቦ	0	0	0
77	ገሰሰ ገበዬሁ	2000	2000	2000
78	ጀመረ አድራሮ	600	600	600
79	ይታይህ ቁምላቸው	400	400	400
80	መንግሥቱ መጫሎ	500	500	500
81	አስራት በዛብህ	2000	2000	2000
82	ጌታቸው ገ/ማሪያም	300	1133.33	300
83	ይድነቃቸው አበበ	3000	3000	3000
84	አህመድ የሱፍ	1000	1000	1000
85	ጀማል መሃመድ	0	0	0
86	አምንቴ አዳሾ	500	500	500
87	ምንትዋብ አምበሎ	1000	1000	1000
88	ቢሻሽ በቃሉ	1000	1000	1000
89	ሰለሞን ሾደኖ	2000	2000	2000
90	ቡዛዬሁ ጋሪፎ	2000	2000	2000
91	እምቢበል ቆጭቶ	1000	1000	1000
92	የሺጥላ ገላን	1000	1000	1000
93	ዮናስ ካዳኔ	1000	1000	1000
94	ቃላአብ ጥኡም	0	0	0
95	ዓለሙ ብርሃኑ	1000	1000	1000
96	ሠላሙ ዳሪቆ	500	500	500
97	ሞሲሳ ቤኛ	2000	2000	2000
98	ተካልኝ አደሞ	500	500	500
99	ወርቅነሽ አንገሎ	1000	1000	1000
100	ባሳዝነው በላይ	2000	3000	3000
101	ፍቃዱ ጉታ	1000	1000	1000
102	ውድነሽ ቀጸላ	1000	1000	1000
103	አመለወርቅ መስፍን	1000	1000	1000
104	የኋላእሸት በቃሉ	500	500	500
105	ሙሴ ካሳሁን	0	0	0
106	ሳልልሽ ጸዳ	2000	2000	2000
107	ተራመድ ገረመው	1000	1000	1000
108	ሚሊዮን አትርሴ	1000	1000	1000
109	ዓለሙ ገ/ማሪያም	1000	1000	1000
110	ዘካሪያስ ቶላ	2000	2000	2000
111	አለማዬሁ ተስፋዬ	2000	2000	2000
112	አጥናፉ ገይቶ	500	500	500
113	ተናኜ አንገሎ	500	500	500
114	ሻመቶ እንደሻው	1000	1000	1000
115	ማርታ አበበ	1000	1000	1000
116	አስናቀች ነገዎ	1000	1000	1000
117	አዲስዓለም አደዮ	1000	1000	1000
118	ታፈሰ ታምሩ	0	0	0
119	ወርቁ ጋዎቶ	700	700	700
120	አንዱዓለም ተስፋዬ	1000	1000	1000
121	አጥናፉ አዴሎ	5000	5000	5000
122	አሸናፊ አበበ	1000	1000	1000
123	ደስታ ሙላት	2000	2000	2000
124	ገዛኸኝ ገደኖ	2000	2000	2000
125	ታደሰ ወ/ሚካኤል	1000	1000	1000
126	አማኑኤል አድነው	1000	1000	1000
127	ሽብሩ ሻወኖ	3000	3000	3000
128	ዳንኤል መስፍን	0	0	0
129	መስፍን ማሞ	1000	1000	1000
130	ባይለየኘ ክብረት	1000	1000	1000
131	ጠና ደጋጋ	0	1000	0
132	ኤርሚያስ ግዛው	0	1000	1000
133	መብራቴ ሻረው	0	500	500
134	ፍጹም ፍስሃ	0	10000	10000
135	ይከበር አለምነህ	0	0	7000
136	ግርማ ግዛው	0	0	1000
"""


# ============================================================
# HELPERS
# ============================================================

def money(value):
    try:
        return f"{float(value):,.2f}"
    except Exception:
        return "0.00"


def safe_float(value, default=0.0):
    try:
        if pd.isna(value):
            return default
        return float(value)
    except Exception:
        return default


def safe_int(value, default=0):
    try:
        if pd.isna(value):
            return default
        return int(float(value))
    except Exception:
        return default


def make_default_national_id(member_number):
    """Generate a temporary 12-digit National ID for existing members."""
    return f"{safe_int(member_number):012d}"


def clean_national_id(national_id):
    """Remove spaces so the National ID is stored as exactly 12 digits."""
    return str(national_id or "").replace(" ", "").strip()


def valid_national_id(national_id):
    """National ID must contain exactly 12 digits."""
    national_id = clean_national_id(national_id)
    return len(national_id) == 12 and national_id.isdigit()


def format_national_id(national_id):
    """Display a 12-digit National ID as 0000 0000 0000."""
    national_id = clean_national_id(national_id)
    if len(national_id) == 12:
        return f"{national_id[:4]} {national_id[4:8]} {national_id[8:]}"
    return national_id


def normalize_member(member):
    """Ensure old and new records have all required fields."""

    for month in MONTHS:
        member[month] = safe_float(member.get(month, 0))

    member["የአባል ቁጥር"] = safe_int(
        member.get("የአባል ቁጥር", 0)
    )

    member["ስም"] = str(member.get("ስም", "") or "")

    national_id = clean_national_id(
        member.get("ብሔራዊ መታወቂያ", "")
    )

    # Existing registered members may have blank or old numeric temporary IDs.
    # Give them temporary sequential 12-digit IDs based on member number.
    if not valid_national_id(national_id):
        national_id = make_default_national_id(
            member["የአባል ቁጥር"]
        )

    member["ብሔራዊ መታወቂያ"] = national_id

    member["ፎቶ"] = str(member.get("ፎቶ", "") or "")

    member["የመመዝገቢያ ክፍያ (ብር)"] = safe_float(
        member.get("የመመዝገቢያ ክፍያ (ብር)", REGISTRATION_FEE)
    )

    member["የተበደረ ብር"] = safe_float(
        member.get("የተበደረ ብር", 0)
    )

    member["10% የብድር ክፍያ (ብር)"] = safe_float(
        member.get("10% የብድር ክፍያ (ብር)", 0)
    )

    member["በእጅ የተሰጠ 90% (ብር)"] = safe_float(
        member.get("በእጅ የተሰጠ 90% (ብር)", 0)
    )

    member["የብድር ጊዜ (ወር)"] = safe_int(
        member.get("የብድር ጊዜ (ወር)", 0)
    )

    member["የብድር ወርሃዊ ክፍያ (ብር)"] = safe_float(
        member.get("የብድር ወርሃዊ ክፍያ (ብር)", 0)
    )

    member["የቀረው ዋና ብድር (ብር)"] = safe_float(
        member.get(
            "የቀረው ዋና ብድር (ብር)",
            member.get("የቀረው ዕዳ (ብር)", 0)
        )
    )

    member["የቀረው ዕዳ (ብር)"] = safe_float(
        member.get(
            "የቀረው ዕዳ (ብር)",
            member.get("የቀረው ዋና ብድር (ብር)", 0)
        )
    )

    member["የተከፈለ ወለድ (ብር)"] = safe_float(
        member.get("የተከፈለ ወለድ (ብር)", 0)
    )

    member["የተከፈለ ዋና ብድር (ብር)"] = safe_float(
        member.get("የተከፈለ ዋና ብድር (ብር)", 0)
    )

    member["የተበደረበት ቀን"] = str(
        member.get("የተበደረበት ቀን", "") or ""
    )

    return member


# ============================================================
# PHASE 3 - AUTHENTICATION HELPERS
# ============================================================

def hash_password(password, salt=None):
    """Hash a password with PBKDF2-HMAC-SHA256."""
    if salt is None:
        salt = secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        str(password).encode("utf-8"),
        salt.encode("utf-8"),
        120000
    ).hex()
    return f"{salt}${digest}"


def verify_password(password, stored_hash):
    try:
        salt, digest = str(stored_hash).split("$", 1)
        check = hashlib.pbkdf2_hmac(
            "sha256",
            str(password).encode("utf-8"),
            salt.encode("utf-8"),
            120000
        ).hex()
        return secrets.compare_digest(check, digest)
    except Exception:
        return False


def create_initial_users():
    return [{
        "username": DEFAULT_ADMIN_USERNAME,
        "password_hash": hash_password(DEFAULT_ADMIN_PASSWORD),
        "role": "Admin",
        "active": "Yes",
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }]


def user_can_access(page_name):
    user = st.session_state.get("current_user")
    if not user:
        return False
    return page_name in ROLE_PERMISSIONS.get(user.get("role"), set())


def current_username():
    user = st.session_state.get("current_user")
    return str(user.get("username", "system")) if user else "system"


def log_action(audit_log, action, details=""):
    audit_log.append({
        "ቀን": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "ተጠቃሚ": current_username(),
        "ተግባር": str(action),
        "ዝርዝር": str(details),
    })


def find_user(users, username):
    target = str(username or "").strip().lower()
    return next(
        (u for u in users if str(u.get("username", "")).strip().lower() == target),
        None
    )


# ============================================================
# INITIAL MEMBERS
# ============================================================

def create_initial_members():

    members = []

    df = pd.read_csv(
        StringIO(MEMBER_DATA.strip()),
        sep="\t",
        header=None,
        names=[
            "የአባል ቁጥር",
            "ስም",
            "መስከረም",
            "ጥቅምት",
            "ኅዳር"
        ]
    )

    for _, row in df.iterrows():

        member = {
            "የአባል ቁጥር": int(row["የአባል ቁጥር"]),
            "ስም": str(row["ስም"]),
            "ብሔራዊ መታወቂያ": make_default_national_id(
                int(row["የአባል ቁጥር"])
            ),
            "ፎቶ": "",

            # The original three values supplied by the user
            # are placed into the first three months.
            "መስከረም": safe_float(row["መስከረም"]),
            "ጥቅምት": safe_float(row["ጥቅምት"]),
            "ኅዳር": safe_float(row["ኅዳር"])
        }

        for month in MONTHS[3:]:
            member[month] = 0.0

        # All 136 existing members already paid the registration fee.
        member["የመመዝገቢያ ክፍያ (ብር)"] = REGISTRATION_FEE

        member["የተበደረ ብር"] = 0.0
        member["10% የብድር ክፍያ (ብር)"] = 0.0
        member["በእጅ የተሰጠ 90% (ብር)"] = 0.0
        member["የብድር ጊዜ (ወር)"] = 0
        member["የብድር ወርሃዊ ክፍያ (ብር)"] = 0.0
        member["የቀረው ዋና ብድር (ብር)"] = 0.0
        member["የቀረው ዕዳ (ብር)"] = 0.0
        member["የተከፈለ ወለድ (ብር)"] = 0.0
        member["የተከፈለ ዋና ብድር (ብር)"] = 0.0
        member["የተበደረበት ቀን"] = ""

        members.append(normalize_member(member))

    return members


# ============================================================
# LOAN CALCULATION
# ============================================================

def calculate_monthly_payment(principal, months, rate=INTEREST_RATE):

    principal = safe_float(principal)

    if principal <= 0 or months <= 0:
        return 0.0

    if rate == 0:
        return principal / months

    payment = (
        principal
        * rate
        * ((1 + rate) ** months)
        / (((1 + rate) ** months) - 1)
    )

    return payment


def calculate_current_interest(member):

    remaining_principal = safe_float(
        member.get("የቀረው ዋና ብድር (ብር)", 0)
    )

    return remaining_principal * INTEREST_RATE


# ============================================================
# LOAN PAYMENT
# ============================================================

def apply_loan_payment(member, payment_amount):

    payment_amount = safe_float(payment_amount)

    if payment_amount <= 0:
        return {
            "success": False,
            "message": "የክፍያ መጠን ከ0 በላይ መሆን አለበት።"
        }

    remaining_principal = safe_float(
        member.get("የቀረው ዋና ብድር (ብር)", 0)
    )

    if remaining_principal <= 0:
        return {
            "success": False,
            "message": "ይህ አባል የሚከፈል የቀረ ብድር የለውም።"
        }

    # Current month's interest is based ONLY on remaining principal.
    interest_due = remaining_principal * INTEREST_RATE

    # First pay current interest.
    interest_paid = min(payment_amount, interest_due)

    remaining_payment = payment_amount - interest_paid

    # Then pay principal.
    principal_paid = min(
        remaining_payment,
        remaining_principal
    )

    actual_payment = interest_paid + principal_paid

    new_remaining_principal = max(
        0.0,
        remaining_principal - principal_paid
    )

    member["የቀረው ዋና ብድር (ብር)"] = new_remaining_principal
    member["የቀረው ዕዳ (ብር)"] = new_remaining_principal

    member["የተከፈለ ወለድ (ብር)"] = (
        safe_float(member.get("የተከፈለ ወለድ (ብር)", 0))
        + interest_paid
    )

    member["የተከፈለ ዋና ብድር (ብር)"] = (
        safe_float(member.get("የተከፈለ ዋና ብድር (ብር)", 0))
        + principal_paid
    )

    return {
        "success": True,
        "payment": actual_payment,
        "interest": interest_paid,
        "principal": principal_paid,
        "remaining": new_remaining_principal
    }


# ============================================================
# PHASE 2 - BACKUP & RECEIPT HELPERS
# ============================================================

def create_database_backup():
    """Create a timestamped backup of the current Excel database."""
    if not os.path.exists(DB_FILE):
        return None

    backup_dir = "sacco_backups"
    os.makedirs(backup_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = os.path.join(
        backup_dir,
        f"sacco_database_backup_{timestamp}.xlsx"
    )

    shutil.copy2(DB_FILE, backup_file)
    return backup_file


def cleanup_old_backups(max_backups=30):
    """Keep the newest backup files only."""
    backup_dir = "sacco_backups"
    if not os.path.isdir(backup_dir):
        return

    files = [
        os.path.join(backup_dir, f)
        for f in os.listdir(backup_dir)
        if f.lower().endswith(".xlsx")
    ]

    files.sort(key=lambda x: os.path.getmtime(x), reverse=True)

    for old_file in files[max_backups:]:
        try:
            os.remove(old_file)
        except OSError:
            pass


def make_receipt_number(payment_history):
    """Generate a simple sequential receipt number."""
    highest = 0
    for item in payment_history:
        value = str(item.get("ደረሰኝ ቁጥር", "")).strip()
        if value.startswith("RC-"):
            try:
                highest = max(highest, int(value.split("-")[-1]))
            except ValueError:
                pass
    return f"RC-{highest + 1:06d}"


def generate_receipt_html(receipt):
    """Create a clean printable HTML receipt."""
    def esc(value):
        return (str(value)
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace('"', "&quot;"))

    return f"""<!DOCTYPE html>
<html lang="am">
<head>
<meta charset="UTF-8">
<title>Receipt {esc(receipt.get('ደረሰኝ ቁጥር',''))}</title>
<style>
body {{ font-family: Arial, 'Noto Sans Ethiopic', sans-serif; margin: 40px; }}
.receipt {{ max-width: 720px; margin: auto; border: 1px solid #ccc; padding: 30px; }}
h1 {{ text-align: center; margin-bottom: 5px; }}
.sub {{ text-align: center; color: #555; margin-bottom: 25px; }}
table {{ width: 100%; border-collapse: collapse; }}
td {{ padding: 10px; border-bottom: 1px solid #eee; }}
td:first-child {{ font-weight: bold; width: 45%; }}
.total {{ font-size: 20px; font-weight: bold; }}
.footer {{ margin-top: 30px; text-align: center; color: #666; }}
</style>
</head>
<body>
<div class="receipt">
<h1>🏦 ተስፋ የገንዘብ ቁጠባና ብድር ማህበር</h1>
<div class="sub">የብድር ክፍያ ደረሰኝ</div>
<table>
<tr><td>ደረሰኝ ቁጥር</td><td>{esc(receipt.get('ደረሰኝ ቁጥር',''))}</td></tr>
<tr><td>ቀን</td><td>{esc(receipt.get('ቀን',''))}</td></tr>
<tr><td>የአባል ቁጥር</td><td>{esc(receipt.get('የአባል ቁጥር',''))}</td></tr>
<tr><td>ስም</td><td>{esc(receipt.get('ስም',''))}</td></tr>
<tr><td>ብሔራዊ መታወቂያ</td><td>{esc(format_national_id(receipt.get('ብሔራዊ መታወቂያ','')))}</td></tr>
<tr><td>የተከፈለ ጠቅላላ</td><td class="total">{money(receipt.get('የተከፈለ ጠቅላላ',0))} ብር</td></tr>
<tr><td>ወለድ</td><td>{money(receipt.get('ወለድ',0))} ብር</td></tr>
<tr><td>ዋና ብድር</td><td>{money(receipt.get('ዋና ብድር',0))} ብር</td></tr>
<tr><td>የቀረ ዋና ብድር</td><td>{money(receipt.get('የቀረ ዋና ብድር',0))} ብር</td></tr>
</table>
<div class="footer">እናመሰግናለን።</div>
</div>
</body>
</html>"""


# ============================================================
# EXCEL SAVE
# ============================================================

def save_data_to_excel(members, payment_history=None, users=None, audit_log=None):

    if payment_history is None:
        payment_history = []

    if users is None:
        users = st.session_state.get("users", create_initial_users())

    if audit_log is None:
        audit_log = st.session_state.get("audit_log", [])

    # Phase 2 safety: backup the existing database before replacing it.
    if os.path.exists(DB_FILE):
        try:
            create_database_backup()
            cleanup_old_backups(max_backups=30)
        except Exception:
            # Never block normal saving if backup creation fails.
            pass

    member_df = pd.DataFrame(members)

    # Keep important columns in an organized order.
    preferred_columns = [
        "የአባል ቁጥር",
        "ስም",
        "ብሔራዊ መታወቂያ",
        "ፎቶ",
        "የመመዝገቢያ ክፍያ (ብር)"
    ]

    preferred_columns += MONTHS

    preferred_columns += [
        "የተበደረ ብር",
        "10% የብድር ክፍያ (ብር)",
        "በእጅ የተሰጠ 90% (ብር)",
        "የብድር ጊዜ (ወር)",
        "የብድር ወርሃዊ ክፍያ (ብር)",
        "የቀረው ዋና ብድር (ብር)",
        "የቀረው ዕዳ (ብር)",
        "የተከፈለ ወለድ (ብር)",
        "የተከፈለ ዋና ብድር (ብር)",
        "የተበደረበት ቀን"
    ]

    existing_columns = [
        col for col in preferred_columns
        if col in member_df.columns
    ]

    remaining_columns = [
        col for col in member_df.columns
        if col not in existing_columns
    ]

    member_df = member_df[
        existing_columns + remaining_columns
    ]

    history_df = pd.DataFrame(payment_history)

    with pd.ExcelWriter(
        DB_FILE,
        engine="openpyxl"
    ) as writer:

        member_df.to_excel(
            writer,
            sheet_name="አባላት",
            index=False
        )

        history_df.to_excel(
            writer,
            sheet_name="የብድር_ክፍያ_ታሪክ",
            index=False
        )

        pd.DataFrame(users).to_excel(
            writer,
            sheet_name="ተጠቃሚዎች",
            index=False
        )

        pd.DataFrame(audit_log).to_excel(
            writer,
            sheet_name="የስርዓት_ታሪክ",
            index=False
        )


# ============================================================
# EXCEL LOAD
# ============================================================

def load_database():

    if not os.path.exists(DB_FILE):

        members = create_initial_members()
        payment_history = []

        users = create_initial_users()
        audit_log = []

        save_data_to_excel(
            members,
            payment_history,
            users,
            audit_log
        )

        return members, payment_history, users, audit_log

    try:

        excel = pd.ExcelFile(DB_FILE)

        # -----------------------------
        # Members
        # -----------------------------

        if "አባላት" in excel.sheet_names:

            df = pd.read_excel(
                DB_FILE,
                sheet_name="አባላት"
            )

        else:

            df = pd.read_excel(
                DB_FILE,
                sheet_name=excel.sheet_names[0]
            )

        members = []

        for _, row in df.iterrows():

            member = row.to_dict()

            # Remove completely empty rows
            if not str(member.get("ስም", "")).strip():
                continue

            members.append(
                normalize_member(member)
            )

        # -----------------------------
        # Payment History
        # -----------------------------

        if "የብድር_ክፍያ_ታሪክ" in excel.sheet_names:

            history_df = pd.read_excel(
                DB_FILE,
                sheet_name="የብድር_ክፍያ_ታሪክ"
            )

            payment_history = history_df.fillna("").to_dict(
                orient="records"
            )

        else:

            payment_history = []

        # If no valid members exist, initialize.
        if len(members) == 0:

            members = create_initial_members()

            save_data_to_excel(
                members,
                payment_history
            )

        else:
            # Persist temporary National IDs for existing members
            # that were previously registered without an ID.
            ids_changed = False

            for member in members:
                expected_id = make_default_national_id(
                    member.get("የአባል ቁጥር", 0)
                )

                if not clean_national_id(
                    member.get("ብሔራዊ መታወቂያ", "")
                ):
                    member["ብሔራዊ መታወቂያ"] = expected_id
                    ids_changed = True

            if ids_changed:
                save_data_to_excel(
                    members,
                    payment_history
                )

        # -----------------------------
        # Users
        # -----------------------------
        if "ተጠቃሚዎች" in excel.sheet_names:
            users_df = pd.read_excel(
                DB_FILE,
                sheet_name="ተጠቃሚዎች"
            )
            users = users_df.fillna("").to_dict(orient="records")
        else:
            users = create_initial_users()

        # Guarantee an active admin exists if the users sheet is empty.
        if not users:
            users = create_initial_users()

        # -----------------------------
        # Audit Trail
        # -----------------------------
        if "የስርዓት_ታሪክ" in excel.sheet_names:
            audit_df = pd.read_excel(
                DB_FILE,
                sheet_name="የስርዓት_ታሪክ"
            )
            audit_log = audit_df.fillna("").to_dict(orient="records")
        else:
            audit_log = []

        # If Phase 3 sheets did not exist, write them once while
        # preserving all current members and payment history.
        if (
            "ተጠቃሚዎች" not in excel.sheet_names
            or "የስርዓት_ታሪክ" not in excel.sheet_names
        ):
            save_data_to_excel(
                members,
                payment_history,
                users,
                audit_log
            )

        return members, payment_history, users, audit_log

    except Exception as e:

        st.error(
            f"Excel database ሲከፈት ችግር ተፈጥሯል፦ {e}"
        )

        return create_initial_members(), [], create_initial_users(), []


# ============================================================
# DUPLICATE NATIONAL ID
# ============================================================

def national_id_exists(
    members,
    national_id,
    exclude_member_number=None
):

    national_id = clean_national_id(national_id)

    if not national_id:
        return False

    for member in members:

        if (
            exclude_member_number is not None
            and safe_int(member.get("የአባል ቁጥር"))
            == safe_int(exclude_member_number)
        ):
            continue

        if clean_national_id(
            member.get("ብሔራዊ መታወቂያ", "")
        ) == national_id:

            return True

    return False


# ============================================================
# MEMBER SEARCH / SELECTION
# ============================================================

def member_search_selector(members, label="አባል ይምረጡ", key="member_search"):
    """Search members by name, member number, or National ID."""
    search_text = st.text_input(
        "🔎 አባል ፈልግ (ስም / አባል ቁጥር / National ID)",
        placeholder="ለምሳሌ፦ 25 ወይም አበበ ወይም 0000 0000 0025",
        key=f"{key}_search"
    )

    search_text = str(search_text or "").strip().lower()
    search_id = clean_national_id(search_text)

    filtered = []
    for member in members:
        name = str(member.get("ስም", ""))
        number = str(safe_int(member.get("የአባል ቁጥር", 0)))
        national_id = clean_national_id(
            member.get("ብሔራዊ መታወቂያ", "")
        )

        if (
            not search_text
            or search_text in name.lower()
            or search_text in number
            or search_id in national_id.lower()
        ):
            filtered.append(member)

    if not filtered:
        st.warning("🔎 የፈለጉት አባል አልተገኘም።")
        return None, None

    options = [
        f'{safe_int(m.get("የአባል ቁጥር"))} - {m.get("ስም", "")} - {format_national_id(m.get("ብሔራዊ መታወቂያ", ""))}'
        for m in filtered
    ]

    selected = st.selectbox(
        label,
        options,
        key=f"{key}_select"
    )

    selected_number = int(selected.split(" - ")[0])
    selected_member = next(
        m for m in filtered
        if safe_int(m.get("የአባል ቁጥር")) == selected_number
    )

    return selected_member, selected_number


def recent_payment_duplicate(payment_history, member_number, amount, window_seconds=120):
    """Detect an accidental repeated loan payment within a short time window."""
    now = datetime.now()
    amount = safe_float(amount)

    for history in reversed(payment_history):
        if safe_int(history.get("የአባል ቁጥር", 0)) != safe_int(member_number):
            continue

        if abs(safe_float(history.get("የተከፈለ ጠቅላላ", 0)) - amount) > 0.01:
            continue

        try:
            payment_time = datetime.strptime(
                str(history.get("ቀን", "")),
                "%Y-%m-%d %H:%M:%S"
            )
        except (ValueError, TypeError):
            continue

        if 0 <= (now - payment_time).total_seconds() <= window_seconds:
            return True

    return False


# ============================================================
# TOTAL SAVINGS
# ============================================================

def total_savings(member):

    return sum(
        safe_float(member.get(month, 0))
        for month in MONTHS
    )


# ============================================================
# DOWNLOAD EXCEL
# ============================================================

def create_excel_download(members, payment_history):

    output = BytesIO()

    member_df = pd.DataFrame(members)
    history_df = pd.DataFrame(payment_history)

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        member_df.to_excel(
            writer,
            sheet_name="አባላት",
            index=False
        )

        history_df.to_excel(
            writer,
            sheet_name="የብድር_ክፍያ_ታሪክ",
            index=False
        )

    output.seek(0)

    return output.getvalue()


# ============================================================
# LOAD DATA INTO SESSION
# ============================================================

if "members" not in st.session_state:

    members, payment_history, users, audit_log = load_database()

    st.session_state.members = members
    st.session_state.payment_history = payment_history
    st.session_state.users = users
    st.session_state.audit_log = audit_log
    st.session_state.authenticated = False
    st.session_state.current_user = None


members = st.session_state.members
payment_history = st.session_state.payment_history
users = st.session_state.users
audit_log = st.session_state.audit_log


# ============================================================
# HEADER
# ============================================================

st.title("🏦 ተስፋ የገንዘብ ቁጠባና ብድር ማህበር")

st.caption(
    "የቁጠባ፣ የብድር፣ የክፍያ እና የአባላት አስተዳደር ስርዓት"
)


# ============================================================
# PHASE 3 - LOGIN GATE
# ============================================================

if not st.session_state.get("authenticated", False):

    st.title("🔐 የስርዓቱ መግቢያ")

    with st.form("login_form"):
        login_username = st.text_input("👤 Username")
        login_password = st.text_input("🔑 Password", type="password")
        login_submit = st.form_submit_button(
            "➡️ ግባ",
            type="primary"
        )

    if login_submit:
        found_user = find_user(users, login_username)

        if (
            found_user
            and str(found_user.get("active", "Yes")).lower() in ("yes", "true", "1")
            and verify_password(login_password, found_user.get("password_hash", ""))
        ):
            st.session_state.authenticated = True
            st.session_state.current_user = {
                "username": str(found_user.get("username", "")),
                "role": str(found_user.get("role", "Reports")),
            }
            log_action(
                audit_log,
                "LOGIN",
                f"Username={found_user.get('username')} Role={found_user.get('role')}"
            )
            save_data_to_excel(members, payment_history, users, audit_log)
            st.rerun()
        else:
            st.error("❌ Username ወይም Password ትክክል አይደለም።")

    st.info(
        f"መጀመሪያ ጊዜ ከሆነ፦ Username: `{DEFAULT_ADMIN_USERNAME}`  | "
        f"Password: `{DEFAULT_ADMIN_PASSWORD}`"
    )
    st.stop()


# ============================================================
# CURRENT USER / LOGOUT
# ============================================================

current_user = st.session_state.get("current_user") or {}
st.sidebar.success(
    f"👤 {current_user.get('username', '')}\n\n"
    f"🔐 Role: {current_user.get('role', '')}"
)

if st.sidebar.button("🚪 ውጣ"):
    log_action(audit_log, "LOGOUT", f"Username={current_username()}")
    save_data_to_excel(members, payment_history, users, audit_log)
    st.session_state.authenticated = False
    st.session_state.current_user = None
    st.rerun()


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title("📋 ምናሌ")

ALL_PAGES = [
    "🏠 ዋና ዳሽቦርድ",
    "👤 አባል መመዝገቢያ",
    "💰 የወር ቁጠባ ማስገቢያ",
    "💵 የብድር አገልግሎት",
    "📅 የብድር ክፍያ መመዝገቢያ",
    "📊 ጠቅላላ ሪፖርት",
    "✏️ የአባላት መረጃ ማስተካከያ",
    "👥 ተጠቃሚዎች አስተዳደር",
]

allowed_pages = [
    p for p in ALL_PAGES
    if user_can_access(p)
]

page = st.sidebar.radio(
    "ገጽ ይምረጡ",
    allowed_pages
)

st.sidebar.divider()

if st.sidebar.button("💾 የዳታቤዝ Backup ፍጠር"):
    if os.path.exists(DB_FILE):
        try:
            backup_file = create_database_backup()
            cleanup_old_backups(max_backups=30)
            st.sidebar.success("Backup ተፈጥሯል።")
        except Exception as e:
            st.sidebar.error(f"Backup ማድረግ አልተቻለም፦ {e}")
    else:
        st.sidebar.warning("እስካሁን Excel database የለም።")


# ============================================================
# 0. DASHBOARD
# ============================================================

if page == "🏠 ዋና ዳሽቦርድ":

    st.header("🏠 ዋና ዳሽቦርድ")

    total_members = len(members)
    total_savings_all = sum(total_savings(m) for m in members)
    total_loans = sum(safe_float(m.get("የተበደረ ብር", 0)) for m in members)
    total_remaining = sum(safe_float(m.get("የቀረው ዕዳ (ብር)", 0)) for m in members)
    total_interest = sum(safe_float(m.get("የተከፈለ ወለድ (ብር)", 0)) for m in members)
    total_loan_fees = sum(safe_float(m.get("10% የብድር ክፍያ (ብር)", 0)) for m in members)
    total_registration = sum(safe_float(m.get("የመመዝገቢያ ክፍያ (ብር)", 0)) for m in members)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("👥 አባላት", total_members)
    c2.metric("💰 ጠቅላላ ቁጠባ", f"{money(total_savings_all)} ብር")
    c3.metric("💵 የተበደረ", f"{money(total_loans)} ብር")
    c4.metric("📌 ቀሪ ዕዳ", f"{money(total_remaining)} ብር")

    st.divider()

    d1, d2, d3, d4 = st.columns(4)
    d1.metric("📈 የተከፈለ ወለድ", f"{money(total_interest)} ብር")
    d2.metric("💳 10% የብድር ክፍያ", f"{money(total_loan_fees)} ብር")
    d3.metric("🧾 መመዝገቢያ", f"{money(total_registration)} ብር")
    d4.metric("💼 ጠቅላላ ገቢ", f"{money(total_interest + total_loan_fees + total_registration)} ብር")

    st.divider()

    st.subheader("📊 የወር ቁጠባ ሁኔታ")
    monthly_totals = pd.DataFrame({
        "ወር": MONTHS,
        "ጠቅላላ ቁጠባ": [
            sum(safe_float(m.get(month, 0)) for m in members)
            for month in MONTHS
        ]
    })
    st.bar_chart(monthly_totals.set_index("ወር"))

    st.subheader("💳 የብድር ሁኔታ")
    loan_status = pd.DataFrame({
        "ሁኔታ": ["የተበደረ ዋና ብድር", "ቀሪ ዋና ብድር"],
        "መጠን": [
            total_loans,
            sum(safe_float(m.get("የቀረው ዋና ብድር (ብር)", 0)) for m in members)
        ]
    })
    st.bar_chart(loan_status.set_index("ሁኔታ"))

    st.subheader("🕒 የቅርብ ጊዜ የብድር ክፍያዎች")
    if payment_history:
        recent_df = pd.DataFrame(payment_history).tail(10).iloc[::-1]
        st.dataframe(recent_df, use_container_width=True, hide_index=True)
    else:
        st.info("እስካሁን የብድር ክፍያ የለም።")

    excel_data = create_excel_download(members, payment_history)
    st.download_button(
        "📥 የአሁኑን Excel ዳታ አውርድ",
        data=excel_data,
        file_name="sacco_database.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ============================================================
# 1. MEMBER REGISTRATION
# ============================================================

if page == "👤 አባል መመዝገቢያ":

    st.header("👤 አዲስ አባል መመዝገቢያ")

    st.info(
        f"የመመዝገቢያ ክፍያ፦ {money(REGISTRATION_FEE)} ብር"
    )

    col1, col2 = st.columns(2)

    with col1:

        name = st.text_input(
            "የአባል ስም"
        )

        national_id = st.text_input(
            "ብሔራዊ መታወቂያ *",
            placeholder="0000 0000 0000",
            help="12 ዲጂት መሆን አለበት። ለምሳሌ፦ 0000 0000 0001"
        )

    with col2:

        registration_fee = st.number_input(
            "የመመዝገቢያ ክፍያ (ብር)",
            min_value=0.0,
            value=REGISTRATION_FEE,
            step=50.0
        )

        photo = st.text_input(
            "የፎቶ መረጃ / Photo (ከፈለጉ)"
        )

    if st.button(
        "➕ አባል መመዝገብ",
        type="primary"
    ):

        if not name.strip():

            st.error("የአባል ስም ያስገቡ።")

        elif not valid_national_id(national_id):

            st.error(
                "ብሔራዊ መታወቂያ 12 ዲጂት መሆን አለበት። "
                "ለምሳሌ፦ 0000 0000 0001"
            )

        elif national_id_exists(
            members,
            clean_national_id(national_id)
        ):

            st.error(
                "ይህ ብሔራዊ መታወቂያ በሲስተሙ ውስጥ አስቀድሞ አለ።"
            )

        else:

            existing_numbers = [
                safe_int(m.get("የአባል ቁጥር"))
                for m in members
            ]

            new_number = (
                max(existing_numbers) + 1
                if existing_numbers
                else 1
            )

            new_member = {
                "የአባል ቁጥር": new_number,
                "ስም": name.strip(),
                "ብሔራዊ መታወቂያ": clean_national_id(national_id),
                "ፎቶ": photo.strip(),
                "የመመዝገቢያ ክፍያ (ብር)": registration_fee,

                "የተበደረ ብር": 0.0,
                "10% የብድር ክፍያ (ብር)": 0.0,
                "በእጅ የተሰጠ 90% (ብር)": 0.0,

                "የብድር ጊዜ (ወር)": 0,
                "የብድር ወርሃዊ ክፍያ (ብር)": 0.0,

                "የቀረው ዋና ብድር (ብር)": 0.0,
                "የቀረው ዕዳ (ብር)": 0.0,

                "የተከፈለ ወለድ (ብር)": 0.0,
                "የተከፈለ ዋና ብድር (ብር)": 0.0,
                "የተበደረበት ቀን": ""
            }

            for month in MONTHS:
                new_member[month] = 0.0

            members.append(
                normalize_member(new_member)
            )

            log_action(
                audit_log,
                "REGISTER_MEMBER",
                f"Member={new_number} Name={name.strip()}"
            )

            save_data_to_excel(
                members,
                payment_history
            )

            st.session_state.members = members
            st.session_state.audit_log = audit_log

            st.success(
                f"{name} በአባል ቁጥር {new_number} ተመዝግቧል። "
                f"የመመዝገቢያ ክፍያ {money(registration_fee)} ብር ተመዝግቧል።"
            )

            st.rerun()


# ============================================================
# 2. MONTHLY SAVINGS
# ============================================================

elif page == "💰 የወር ቁጠባ ማስገቢያ":

    st.header("💰 የወር ቁጠባ ማስገቢያ")

    member, selected_number = member_search_selector(
        members,
        key="savings_member"
    )

    if member is not None:

        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric(
                "የአባል ስም",
                member["ስም"]
            )

        with col2:
            st.metric(
                "ጠቅላላ ቁጠባ",
                f"{money(total_savings(member))} ብር"
            )

        with col3:
            st.metric(
                "የቀረ ብድር",
                f"{money(member.get('የቀረው ዕዳ (ብር)', 0))} ብር"
            )

        month = st.selectbox(
            "ወር ይምረጡ",
            MONTHS
        )

        existing_amount = safe_float(
            member.get(month, 0)
        )

        if existing_amount > 0:

            st.warning(
                f"⚠️ {month} ውስጥ {money(existing_amount)} ብር "
                "ቁጠባ አስቀድሞ አለ።"
            )

            mode = st.radio(
                "ምን ማድረግ ይፈልጋሉ?",
                [
                    "➕ አዲስ ቁጠባ ጨምር",
                    "✏️ ያለውን መጠን አስተካክል"
                ]
            )

        else:

            mode = "➕ አዲስ ቁጠባ ጨምር"

        amount = st.number_input(
            "የቁጠባ መጠን (ብር)",
            min_value=0.0,
            step=50.0
        )

        allow_savings_save = True
        if existing_amount > 0 and mode.startswith("➕"):
            allow_savings_save = st.checkbox(
                f"⚠️ {month} ላይ ቀድሞ {money(existing_amount)} ብር አለ። አዲሱን ቁጠባ በዚህ ወር ላይ ለመጨመር አረጋግጣለሁ።",
                key=f"confirm_savings_{selected_number}_{month}"
            )

        savings_payment_duplicate = False
        confirm_savings_payment_duplicate = True
        if safe_float(member.get("የቀረው ዕዳ (ብር)", 0)) > 0 and amount > 0:
            savings_payment_duplicate = recent_payment_duplicate(
                payment_history,
                selected_number,
                amount
            )
            if savings_payment_duplicate:
                st.warning(
                    "⚠️ ይህ አባል በቅርቡ ተመሳሳይ የብድር ክፍያ መጠን ከፍሏል።"
                )
                confirm_savings_payment_duplicate = st.checkbox(
                    "እርግጠኛ ነኝ፤ ይህን ክፍያ እንደገና መመዝገብ እፈልጋለሁ።",
                    key=f"confirm_savings_payment_duplicate_{selected_number}_{amount}"
                )

        if st.button(
            "💾 ቁጠባ መዝግብ",
            type="primary"
        ):

            if amount <= 0:

                st.error(
                    "የቁጠባ መጠን ከ0 በላይ ይሁን።"
                )

            elif not allow_savings_save:

                st.warning(
                    "አዲስ ቁጠባ ለመጨመር እባክዎ የማረጋገጫ ሳጥኑን ይምረጡ።"
                )

            elif savings_payment_duplicate and not confirm_savings_payment_duplicate:

                st.error(
                    "የተደጋጋሚ የብድር ክፍያ ለመከላከል እባክዎ ክፍያውን ያረጋግጡ።"
                )

            else:

                # ------------------------------------------------
                # If member has outstanding loan:
                # savings payment is first used for the loan.
                # Any excess goes to savings.
                # ------------------------------------------------

                remaining_debt = safe_float(
                    member.get("የቀረው ዕዳ (ብር)", 0)
                )

                if remaining_debt > 0:

                    result = apply_loan_payment(
                        member,
                        amount
                    )

                    if result["success"]:

                        loan_paid = result["payment"]

                        excess = max(
                            0.0,
                            amount - loan_paid
                        )

                        if excess > 0:

                            if mode.startswith("✏️"):

                                old_value = existing_amount

                                member[month] = excess

                                st.info(
                                    f"{month} ቁጠባ ከ "
                                    f"{money(old_value)} ወደ "
                                    f"{money(excess)} ብር ተስተካክሏል።"
                                )

                            else:

                                member[month] = (
                                    existing_amount + excess
                                )

                        payment_history.append({
                            "የአባል ቁጥር": selected_number,
                            "ስም": member["ስም"],
                            "ቀን": datetime.now().strftime(
                                "%Y-%m-%d %H:%M:%S"
                            ),
                            "የተከፈለ ጠቅላላ": loan_paid,
                            "ወለድ": result["interest"],
                            "ዋና ብድር": result["principal"],
                            "የቀረ ዋና ብድር": result["remaining"],
                            "የብድር ጊዜ": member.get(
                                "የብድር ጊዜ (ወር)", 0
                            ),
                            "ወርሃዊ ክፍያ": member.get(
                                "የብድር ወርሃዊ ክፍያ (ብር)", 0
                            )
                        })

                        log_action(
                            audit_log,
                            "LOAN_PAYMENT",
                            f"Member={selected_number} Amount={money(loan_paid)}"
                        )

                        save_data_to_excel(
                            members,
                            payment_history
                        )

                        st.session_state.members = members
                        st.session_state.audit_log = audit_log
                        st.session_state.payment_history = payment_history

                        st.success(
                            f"{money(loan_paid)} ብር ወደ ብድር ክፍያ ገብቷል።"
                        )

                        if excess > 0:
                            st.success(
                                f"{money(excess)} ብር ወደ {month} ቁጠባ ተጨምሯል።"
                            )

                        st.rerun()

                else:

                    old_value = existing_amount

                    if mode.startswith("✏️"):

                        # Replace existing amount.
                        member[month] = amount

                        st.success(
                            f"{month} የነበረው "
                            f"{money(old_value)} ብር ወደ "
                            f"{money(amount)} ብር ተስተካክሏል።"
                        )

                    else:

                        # Add new deposit.
                        member[month] = (
                            existing_amount + amount
                        )

                        st.success(
                            f"{money(amount)} ብር ወደ "
                            f"{month} ቁጠባ ተጨምሯል።"
                        )

                    log_action(
                        audit_log,
                        "SAVINGS_PAYMENT",
                        f"Member={selected_number} Month={month} Amount={money(amount)} Mode={mode}"
                    )

                    save_data_to_excel(
                        members,
                        payment_history
                    )

                    st.session_state.members = members
                    st.session_state.audit_log = audit_log

                    st.rerun()

        st.subheader("📋 የወራት ቁጠባ")

        savings_display = {
            "ወር": MONTHS,
            "ቁጠባ (ብር)": [
                safe_float(member.get(m, 0))
                for m in MONTHS
            ]
        }

        st.dataframe(
            pd.DataFrame(savings_display),
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# 3. LOAN SERVICE
# ============================================================

elif page == "💵 የብድር አገልግሎት":

    st.header("💵 የብድር አገልግሎት")

    member, selected_number = member_search_selector(
        members,
        key="loan_member"
    )

    if member is not None:

        savings = total_savings(member)

        current_debt = safe_float(
            member.get("የቀረው ዕዳ (ብር)", 0)
        )

        max_loan = savings * 4

        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric(
                "ጠቅላላ ቁጠባ",
                f"{money(savings)} ብር"
            )

        with col2:
            st.metric(
                "ከፍተኛ ብድር",
                f"{money(max_loan)} ብር"
            )

        with col3:
            st.metric(
                "የቀረ ብድር",
                f"{money(current_debt)} ብር"
            )

        if current_debt > 0:

            st.warning(
                "⚠️ ይህ አባል አሁንም ያልተከፈለ ብድር ስላለበት "
                "አዲስ ብድር መውሰድ አይችልም።"
            )

        else:

            loan_amount = st.number_input(
                "የብድር መጠን (ብር)",
                min_value=0.0,
                max_value=float(max_loan),
                step=100.0
            )

            term = st.selectbox(
                "የብድር ጊዜ",
                LOAN_TERMS,
                format_func=lambda x: f"{x} ወር"
            )

            monthly_payment = calculate_monthly_payment(
                loan_amount,
                term
            )

            loan_fee = loan_amount * LOAN_FEE_RATE
            cash_given = loan_amount - loan_fee

            st.subheader("🧮 የብድር ስሌት")

            calc1, calc2, calc3, calc4 = st.columns(4)

            with calc1:
                st.metric(
                    "የብድር መጠን",
                    f"{money(loan_amount)} ብር"
                )

            with calc2:
                st.metric(
                    "10% የብድር ክፍያ",
                    f"{money(loan_fee)} ብር"
                )

            with calc3:
                st.metric(
                    "በእጅ የሚሰጠው 90%",
                    f"{money(cash_given)} ብር"
                )

            with calc4:
                st.metric(
                    "የወር ክፍያ",
                    f"{money(monthly_payment)} ብር"
                )

            st.info(
                "ℹ️ የ10% ብር ከብድሩ ውስጥ የሚቆረጥ የብድር ክፍያ ነው። "
                "የአባሉ ዕዳ ግን የተበደረው ሙሉ መጠን ነው።"
            )

            if st.button(
                "💵 ብድር መፍቀድ",
                type="primary"
            ):

                if loan_amount <= 0:

                    st.error(
                        "የብድር መጠን ከ0 በላይ ይሁን።"
                    )

                elif loan_amount > max_loan:

                    st.error(
                        f"ከፍተኛው ብድር {money(max_loan)} ብር ነው።"
                    )

                else:

                    member["የተበደረ ብር"] = loan_amount

                    member["10% የብድር ክፍያ (ብር)"] = (
                        loan_fee
                    )

                    member["በእጅ የተሰጠ 90% (ብር)"] = (
                        cash_given
                    )

                    member["የብድር ጊዜ (ወር)"] = term

                    member["የብድር ወርሃዊ ክፍያ (ብር)"] = (
                        monthly_payment
                    )

                    member["የቀረው ዋና ብድር (ብር)"] = (
                        loan_amount
                    )

                    member["የቀረው ዕዳ (ብር)"] = (
                        loan_amount
                    )

                    member["የተከፈለ ወለድ (ብር)"] = 0.0
                    member["የተከፈለ ዋና ብድር (ብር)"] = 0.0

                    member["የተበደረበት ቀን"] = (
                        datetime.now().strftime("%Y-%m-%d")
                    )

                    log_action(
                        audit_log,
                        "LOAN_APPROVAL",
                        f"Member={selected_number} Amount={money(loan_amount)} Term={term}"
                    )

                    save_data_to_excel(
                        members,
                        payment_history
                    )

                    st.session_state.members = members
                    st.session_state.audit_log = audit_log

                    st.success(
                        f"{member['ስም']} {money(loan_amount)} ብር ብድር "
                        f"ተፈቅዶለታል።"
                    )

                    st.info(
                        f"10% የብድር ክፍያ = {money(loan_fee)} ብር | "
                        f"በእጅ የተሰጠ = {money(cash_given)} ብር"
                    )

                    st.rerun()


    # ============================================================
    # 4. LOAN REPAYMENT
    # ============================================================

elif page == "📅 የብድር ክፍያ መመዝገቢያ":

    st.header("📅 የብድር ክፍያ መመዝገቢያ")

    if "last_receipt" in st.session_state:
        last_receipt = st.session_state.last_receipt
        st.success(
            f"✅ የመጨረሻ ደረሰኝ: {last_receipt.get('ደረሰኝ ቁጥር', '')}"
        )
        receipt_html = generate_receipt_html(last_receipt)
        st.download_button(
            "🧾 ደረሰኝ አውርድ / Print",
            data=receipt_html.encode("utf-8"),
            file_name=f"{last_receipt.get('ደረሰኝ ቁጥር','receipt')}.html",
            mime="text/html",
            key="last_receipt_download"
        )

    member, selected_number = member_search_selector(
        members,
        key="repayment_member"
    )

    if member is None:
        st.stop()

    remaining = safe_float(
        member.get("የቀረው ዋና ብድር (ብር)", 0)
    )

    if remaining <= 0:

        st.success(
            "✅ ይህ አባል ያለበትን ብድር ሙሉ በሙሉ ከፍሏል።"
        )

    else:

        current_interest = calculate_current_interest(
            member
        )

        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric(
                "የቀረ ዋና ብድር",
                f"{money(remaining)} ብር"
            )

        with col2:
            st.metric(
                "የአሁኑ ወር 2% ወለድ",
                f"{money(current_interest)} ብር"
            )

        with col3:
            st.metric(
                "ወርሃዊ የተወሰነ ክፍያ",
                f"{money(member.get('የብድር ወርሃዊ ክፍያ (ብር)', 0))} ብር"
            )

        payment_amount = st.number_input(
            "የክፍያ መጠን (ብር)",
            min_value=0.0,
            step=50.0
        )

        if payment_amount > 0:

            estimated_interest = min(
                payment_amount,
                current_interest
            )

            estimated_principal = max(
                0.0,
                min(
                    remaining,
                    payment_amount - estimated_interest
                )
            )

            st.write(
                f"**የሚከፈለው ወለድ:** "
                f"{money(estimated_interest)} ብር"
            )

            st.write(
                f"**የሚከፈለው ዋና ብድር:** "
                f"{money(estimated_principal)} ብር"
            )

            duplicate_payment_warning = recent_payment_duplicate(
                payment_history,
                selected_number,
                payment_amount
            )

            confirm_repeat_payment = True
            if duplicate_payment_warning:
                st.warning(
                    "⚠️ ይህ አባል በቅርቡ ተመሳሳይ የክፍያ መጠን ከፍሏል። "
                    "ይህ የሁለት ጊዜ መመዝገብ ሊሆን ስለሚችል እባክዎ ያረጋግጡ።"
                )
                confirm_repeat_payment = st.checkbox(
                    "እርግጠኛ ነኝ፤ ይህን ክፍያ እንደገና መመዝገብ እፈልጋለሁ።",
                    key=f"confirm_repeat_payment_{selected_number}_{payment_amount}"
                )

            if st.button(
                "💾 ክፍያ መዝግብ",
                type="primary"
            ):

                if duplicate_payment_warning and not confirm_repeat_payment:
                    st.error(
                        "የተደጋጋሚ ክፍያ ለመከላከል እባክዎ ክፍያውን ያረጋግጡ።"
                    )
                    st.stop()

                result = apply_loan_payment(
                    member,
                    payment_amount
                )

                if result["success"]:

                    receipt_number = make_receipt_number(payment_history)
                    payment_time = datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )

                    payment_history.append({
                        "ደረሰኝ ቁጥር": receipt_number,
                        "የአባል ቁጥር": selected_number,
                        "ስም": member["ስም"],
                        "ብሔራዊ መታወቂያ": clean_national_id(
                            member.get("ብሔራዊ መታወቂያ", "")
                        ),
                        "ቀን": payment_time,
                        "የተከፈለ ጠቅላላ": result["payment"],
                        "ወለድ": result["interest"],
                        "ዋና ብድር": result["principal"],
                        "የቀረ ዋና ብድር": result["remaining"],
                        "የብድር ጊዜ": member.get(
                            "የብድር ጊዜ (ወር)", 0
                        ),
                        "ወርሃዊ ክፍያ": member.get(
                            "የብድር ወርሃዊ ክፍያ (ብር)", 0
                        )
                    })

                    log_action(
                        audit_log,
                        "LOAN_REPAYMENT",
                        f"Member={selected_number} Amount={money(result['payment'])} Receipt={receipt_number}"
                    )

                    save_data_to_excel(
                        members,
                        payment_history
                    )

                    st.session_state.members = members
                    st.session_state.payment_history = payment_history
                    st.session_state.audit_log = audit_log
                    st.session_state.last_receipt = payment_history[-1].copy()

                    st.success(
                        f"ክፍያው ተመዝግቧል። "
                        f"ወለድ: {money(result['interest'])} ብር | "
                        f"ዋና ብድር: {money(result['principal'])} ብር"
                    )

                    if result["remaining"] <= 0.01:

                        st.success(
                            "🎉 የአባሉ ብድር ሙሉ በሙሉ ተከፍሏል።"
                        )

                    st.rerun()

                else:

                    st.error(
                        result["message"]
                    )

    # --------------------------------------------------------
    # PAYMENT HISTORY
    # --------------------------------------------------------

    st.divider()

    st.subheader("📜 የዚህ አባል የብድር ክፍያ ታሪክ")

    member_history = [
        h for h in payment_history
        if safe_int(
            h.get("የአባል ቁጥር", 0)
        ) == selected_number
    ]

    if member_history:

        history_df = pd.DataFrame(
            member_history
        )

        st.dataframe(
            history_df,
            use_container_width=True,
            hide_index=True
        )

    else:

        st.info(
            "ይህ አባል እስካሁን የብድር ክፍያ ታሪክ የለውም።"
        )


# ============================================================
# 5. REPORT
# ============================================================

elif page == "📊 ጠቅላላ ሪፖርት":

    st.header("📊 ጠቅላላ ሪፖርት")

    total_members = len(members)

    total_savings_all = sum(
        total_savings(m)
        for m in members
    )

    total_registration_fees = sum(
        safe_float(
            m.get("የመመዝገቢያ ክፍያ (ብር)", 0)
        )
        for m in members
    )

    total_loans = sum(
        safe_float(
            m.get("የተበደረ ብር", 0)
        )
        for m in members
    )

    total_loan_fees = sum(
        safe_float(
            m.get("10% የብድር ክፍያ (ብር)", 0)
        )
        for m in members
    )

    total_interest = sum(
        safe_float(
            m.get("የተከፈለ ወለድ (ብር)", 0)
        )
        for m in members
    )

    total_remaining = sum(
        safe_float(
            m.get("የቀረው ዕዳ (ብር)", 0)
        )
        for m in members
    )

    total_income = (
        total_registration_fees
        + total_loan_fees
        + total_interest
    )

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.metric(
            "👥 አባላት",
            total_members
        )

    with c2:
        st.metric(
            "💰 ጠቅላላ ቁጠባ",
            f"{money(total_savings_all)} ብር"
        )

    with c3:
        st.metric(
            "💵 ጠቅላላ ብድር",
            f"{money(total_loans)} ብር"
        )

    with c4:
        st.metric(
            "📌 የቀረ ዕዳ",
            f"{money(total_remaining)} ብር"
        )

    st.divider()

    i1, i2, i3, i4 = st.columns(4)

    with i1:
        st.metric(
            "የመመዝገቢያ ክፍያ",
            f"{money(total_registration_fees)} ብር"
        )

    with i2:
        st.metric(
            "10% የብድር ክፍያ",
            f"{money(total_loan_fees)} ብር"
        )

    with i3:
        st.metric(
            "የተከፈለ ወለድ",
            f"{money(total_interest)} ብር"
        )

    with i4:
        st.metric(
            "ጠቅላላ የማህበሩ ገቢ",
            f"{money(total_income)} ብር"
        )

    st.info(
        "የማህበሩ ገቢ = የመመዝገቢያ ክፍያ "
        "+ 10% የብድር ክፍያ + የተከፈለ ወለድ።"
    )

    # --------------------------------------------------------
    # MEMBER REPORT
    # --------------------------------------------------------

    st.subheader("👥 የአባላት ሪፖርት")

    report_rows = []

    for member in members:

        report_rows.append({
            "የአባል ቁጥር": safe_int(
                member.get("የአባል ቁጥር")
            ),

            "ስም": member.get(
                "ስም", ""
            ),

            "ብሔራዊ መታወቂያ": format_national_id(
                member.get("ብሔራዊ መታወቂያ", "")
            ),

            "ጠቅላላ ቁጠባ": total_savings(
                member
            ),

            "የመመዝገቢያ ክፍያ": safe_float(
                member.get(
                    "የመመዝገቢያ ክፍያ (ብር)",
                    0
                )
            ),

            "የተበደረ ብር": safe_float(
                member.get(
                    "የተበደረ ብር",
                    0
                )
            ),

            "10% የብድር ክፍያ": safe_float(
                member.get(
                    "10% የብድር ክፍያ (ብር)",
                    0
                )
            ),

            "በእጅ የተሰጠ 90%": safe_float(
                member.get(
                    "በእጅ የተሰጠ 90% (ብር)",
                    0
                )
            ),

            "የብድር ጊዜ": safe_int(
                member.get(
                    "የብድር ጊዜ (ወር)",
                    0
                )
            ),

            "ወርሃዊ ክፍያ": safe_float(
                member.get(
                    "የብድር ወርሃዊ ክፍያ (ብር)",
                    0
                )
            ),

            "የተከፈለ ወለድ": safe_float(
                member.get(
                    "የተከፈለ ወለድ (ብር)",
                    0
                )
            ),

            "የቀረ ዋና ብድር": safe_float(
                member.get(
                    "የቀረው ዋና ብድር (ብር)",
                    0
                )
            )
        })

    report_df = pd.DataFrame(report_rows)

    # --------------------------------------------------------
    # MEMBER REPORT FILTER
    # --------------------------------------------------------

    filter_text = st.text_input(
        "🔎 አባል በስም፣ በአባል ቁጥር ወይም National ID ፈልግ",
        placeholder="ስም፣ አባል ቁጥር ወይም 0000 0000 0001 ያስገቡ..."
    )

    if filter_text.strip():

        search_text = filter_text.strip().lower()

        filtered_report_df = report_df[
            report_df["ስም"].astype(str).str.lower().str.contains(
                search_text, na=False
            )
            |
            report_df["የአባል ቁጥር"].astype(str).str.contains(
                search_text, na=False
            )
            |
            report_df["ብሔራዊ መታወቂያ"].astype(str).str.replace(" ", "", regex=False).str.contains(
                clean_national_id(search_text), na=False
            )
        ]

    else:

        filtered_report_df = report_df

    st.dataframe(
        filtered_report_df,
        use_container_width=True,
        hide_index=True
    )

    # --------------------------------------------------------
    # FULL PAYMENT HISTORY
    # --------------------------------------------------------

    st.subheader("📜 ሙሉ የብድር ክፍያ ታሪክ")

    if payment_history:

        history_df = pd.DataFrame(
            payment_history
        )

        st.dataframe(
            history_df,
            use_container_width=True,
            hide_index=True
        )

    else:

        st.info(
            "እስካሁን የብድር ክፍያ ታሪክ የለም።"
        )

    # --------------------------------------------------------
    # DOWNLOAD
    # --------------------------------------------------------

    st.divider()

    excel_data = create_excel_download(
        members,
        payment_history
    )

    st.download_button(
        "📥 ሙሉ Excel ሪፖርት አውርድ",
        data=excel_data,
        file_name="sacco_database.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


# ============================================================
# 6. EDIT MEMBERS
# ============================================================

elif page == "✏️ የአባላት መረጃ ማስተካከያ":

    st.header("✏️ የአባላት መረጃ ማስተካከያ")

    member, selected_number = member_search_selector(
        members,
        key="edit_member"
    )

    if member is not None:

        member_index = next(
            i
            for i, m in enumerate(members)
            if safe_int(m.get("የአባል ቁጥር")) == selected_number
        )

        col1, col2 = st.columns(2)

        with col1:

            new_name = st.text_input(
                "ስም",
                value=str(
                    member.get("ስም", "")
                )
            )

            new_national_id = st.text_input(
                "ብሔራዊ መታወቂያ *",
                value=format_national_id(
                    member.get(
                        "ብሔራዊ መታወቂያ",
                        ""
                    )
                ),
                placeholder="0000 0000 0000",
                help="12 ዲጂት መሆን አለበት። ለምሳሌ፦ 0000 0000 0001"
            )

        with col2:

            new_photo = st.text_input(
                "ፎቶ",
                value=str(
                    member.get(
                        "ፎቶ",
                        ""
                    )
                )
            )

            new_registration_fee = st.number_input(
                "የመመዝገቢያ ክፍያ",
                min_value=0.0,
                value=safe_float(
                    member.get(
                        "የመመዝገቢያ ክፍያ (ብር)",
                        REGISTRATION_FEE
                    )
                ),
                step=50.0
            )

        if st.button(
            "💾 ለውጥ አስቀምጥ",
            type="primary"
        ):

            if not new_name.strip():

                st.error(
                    "ስም ባዶ መሆን አይችልም።"
                )

            elif not valid_national_id(new_national_id):

                st.error(
                    "ብሔራዊ መታወቂያ 12 ዲጂት መሆን አለበት። "
                    "ለምሳሌ፦ 0000 0000 0001"
                )

            elif national_id_exists(
                members,
                clean_national_id(new_national_id),
                exclude_member_number=selected_number
            ):

                st.error(
                    "ይህ ብሔራዊ መታወቂያ ሌላ አባል ላይ ተመዝግቧል።"
                )

            else:

                member["ስም"] = new_name.strip()

                member["ብሔራዊ መታወቂያ"] = (
                    clean_national_id(new_national_id)
                )

                member["ፎቶ"] = new_photo.strip()

                member["የመመዝገቢያ ክፍያ (ብር)"] = (
                    new_registration_fee
                )

                log_action(
                    audit_log,
                    "EDIT_MEMBER",
                    f"Member={selected_number} Name={new_name.strip()}"
                )

                save_data_to_excel(
                    members,
                    payment_history
                )

                st.session_state.members = members
                st.session_state.audit_log = audit_log

                st.success(
                    "የአባሉ መረጃ ተስተካክሏል።"
                )

                st.rerun()

        st.divider()

        st.subheader("💰 የአባሉ ቁጠባ")

        savings_table = pd.DataFrame({
            "ወር": MONTHS,
            "ቁጠባ": [
                safe_float(
                    member.get(month, 0)
                )
                for month in MONTHS
            ]
        })

        st.dataframe(
            savings_table,
            use_container_width=True,
            hide_index=True
        )

        st.metric(
            "ጠቅላላ ቁጠባ",
            f"{money(total_savings(member))} ብር"
        )

        st.divider()

        st.subheader("🗑️ አባል ሰርዝ")

        st.warning(
            "⚠️ አባልን መሰረዝ የአባሉን መረጃ ከሲስተሙ ያስወግዳል።"
        )

        confirm_delete = st.checkbox(
            "⚠️ ይህን አባል መሰረዝ እፈልጋለሁ"
        )

        delete_member_number = st.text_input(
            f"ለማረጋገጥ የአባል ቁጥር {selected_number} ያስገቡ",
            placeholder=str(selected_number),
            key=f"delete_confirm_number_{selected_number}"
        )

        if confirm_delete:

            if st.button(
                "🗑️ አባል ሰርዝ",
                type="secondary"
            ):

                if delete_member_number.strip() != str(selected_number):
                    st.error(
                        f"ለማረጋገጥ የአባል ቁጥር {selected_number} በትክክል ያስገቡ።"
                    )
                    st.stop()

                deleted_name = member.get("ስም", "")
                members.pop(member_index)

                log_action(
                    audit_log,
                    "DELETE_MEMBER",
                    f"Member={selected_number} Name={deleted_name}"
                )

                save_data_to_excel(
                    members,
                    payment_history
                )

                st.session_state.members = members
                st.session_state.audit_log = audit_log

                st.success(
                    "አባሉ ተሰርዟል።"
                )

                st.rerun()


# ============================================================
# 7. USER MANAGEMENT / CHANGE PASSWORD
# ============================================================

elif page == "👥 ተጠቃሚዎች አስተዳደር":

    st.header("👥 ተጠቃሚዎች አስተዳደር")

    if current_user.get("role") != "Admin":
        st.error("❌ ይህን ገጽ ለመጠቀም Admin መሆን አለብዎት።")
        st.stop()

    st.subheader("➕ አዲስ ተጠቃሚ ፍጠር")

    with st.form("create_user_form"):
        new_username = st.text_input("Username")
        new_password = st.text_input("Password", type="password")
        new_role = st.selectbox("Role", ALL_ROLES[1:])
        create_user_submit = st.form_submit_button(
            "➕ ተጠቃሚ ፍጠር",
            type="primary"
        )

    if create_user_submit:

        if not new_username.strip():
            st.error("Username ያስገቡ።")

        elif len(new_password) < 6:
            st.error("Password ቢያንስ 6 ቁምፊዎች ይኑረው።")

        elif find_user(users, new_username):
            st.error("ይህ Username አስቀድሞ አለ።")

        else:
            users.append({
                "username": new_username.strip(),
                "password_hash": hash_password(new_password),
                "role": new_role,
                "active": "Yes",
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })

            log_action(
                audit_log,
                "CREATE_USER",
                f"Username={new_username.strip()} Role={new_role}"
            )

            save_data_to_excel(members, payment_history, users, audit_log)
            st.session_state.users = users
            st.session_state.audit_log = audit_log
            st.success("ተጠቃሚው ተፈጥሯል።")
            st.rerun()

    st.divider()

    st.subheader("👥 ያሉ ተጠቃሚዎች")

    users_display = pd.DataFrame([
        {
            "Username": u.get("username", ""),
            "Role": u.get("role", ""),
            "Active": u.get("active", ""),
            "Created": u.get("created_at", ""),
        }
        for u in users
    ])

    st.dataframe(
        users_display,
        use_container_width=True,
        hide_index=True
    )

    st.divider()

    st.subheader("⚙️ Role / Active ሁኔታ ማስተካከያ")

    managed_username = st.selectbox(
        "ተጠቃሚ ይምረጡ",
        [str(u.get("username", "")) for u in users],
        key="managed_user_select"
    )

    managed_user = find_user(users, managed_username)

    if managed_user:
        with st.form("manage_user_form"):
            managed_role = st.selectbox(
                "Role",
                ALL_ROLES,
                index=(
                    ALL_ROLES.index(str(managed_user.get("role", "Reports")))
                    if str(managed_user.get("role", "Reports")) in ALL_ROLES
                    else ALL_ROLES.index("Reports")
                )
            )
            managed_active = st.checkbox(
                "ተጠቃሚው Active ነው",
                value=str(managed_user.get("active", "Yes")).lower() in ("yes", "true", "1")
            )
            manage_user_submit = st.form_submit_button(
                "💾 Role / Active አስቀምጥ"
            )

        if manage_user_submit:
            # Do not allow the only Admin to be disabled or demoted.
            admin_count = sum(
                1 for u in users
                if str(u.get("role", "")) == "Admin"
                and str(u.get("active", "Yes")).lower() in ("yes", "true", "1")
            )

            if (
                str(managed_user.get("role", "")) == "Admin"
                and managed_role != "Admin"
                and admin_count <= 1
            ):
                st.error("❌ ቢያንስ አንድ Active Admin መኖር አለበት።")
            elif (
                str(managed_user.get("role", "")) == "Admin"
                and not managed_active
                and admin_count <= 1
            ):
                st.error("❌ ቢያንስ አንድ Active Admin መኖር አለበት።")
            else:
                old_role = managed_user.get("role", "")
                old_active = managed_user.get("active", "")
                managed_user["role"] = managed_role
                managed_user["active"] = "Yes" if managed_active else "No"

                log_action(
                    audit_log,
                    "UPDATE_USER",
                    f"Username={managed_username} Role={old_role}->{managed_role} Active={old_active}->{managed_user['active']}"
                )

                save_data_to_excel(members, payment_history, users, audit_log)
                st.session_state.users = users
                st.session_state.audit_log = audit_log
                st.success("የተጠቃሚው Role / Active ሁኔታ ተስተካክሏል።")
                st.rerun()

    st.divider()

    st.subheader("🗑️ ተጠቃሚ ሰርዝ")

    delete_username = st.selectbox(
        "ለመሰረዝ ተጠቃሚ",
        [str(u.get("username", "")) for u in users],
        key="delete_user_select"
    )

    delete_user = find_user(users, delete_username)

    if delete_user:
        confirm_user_delete = st.checkbox(
            f"⚠️ `{delete_username}` ተጠቃሚን መሰረዝ እፈልጋለሁ",
            key="confirm_user_delete"
        )

        if confirm_user_delete and st.button("🗑️ ተጠቃሚውን ሰርዝ"):
            admin_count = sum(
                1 for u in users
                if str(u.get("role", "")) == "Admin"
                and str(u.get("active", "Yes")).lower() in ("yes", "true", "1")
            )

            if (
                str(delete_user.get("role", "")) == "Admin"
                and str(delete_user.get("active", "Yes")).lower() in ("yes", "true", "1")
                and admin_count <= 1
            ):
                st.error("❌ የመጨረሻውን Active Admin መሰረዝ አይቻልም።")
            elif str(delete_username).strip().lower() == current_username().strip().lower():
                st.error("❌ እርስዎ እየገቡበት ያለውን account እራስዎ መሰረዝ አይችሉም።")
            else:
                users[:] = [
                    u for u in users
                    if str(u.get("username", "")).strip().lower()
                    != str(delete_username).strip().lower()
                ]

                log_action(
                    audit_log,
                    "DELETE_USER",
                    f"Username={delete_username}"
                )

                save_data_to_excel(members, payment_history, users, audit_log)
                st.session_state.users = users
                st.session_state.audit_log = audit_log
                st.success("ተጠቃሚው ተሰርዟል።")
                st.rerun()

    st.divider()

    st.subheader("🔑 Password ቀይር")

    with st.form("change_password_form"):
        target_username = st.selectbox(
            "ተጠቃሚ",
            [str(u.get("username", "")) for u in users]
        )
        old_password = st.text_input("የአሁኑ Password", type="password")
        replacement_password = st.text_input("አዲስ Password", type="password")
        change_password_submit = st.form_submit_button(
            "🔑 Password ቀይር"
        )

    if change_password_submit:

        target_user = find_user(users, target_username)

        if not target_user:
            st.error("ተጠቃሚው አልተገኘም።")

        elif len(replacement_password) < 6:
            st.error("አዲሱ Password ቢያንስ 6 ቁምፊዎች ይኑረው።")

        elif not verify_password(old_password, target_user.get("password_hash", "")):
            st.error("የአሁኑ Password ትክክል አይደለም።")

        else:
            target_user["password_hash"] = hash_password(replacement_password)

            log_action(
                audit_log,
                "CHANGE_PASSWORD",
                f"Username={target_username}"
            )

            save_data_to_excel(members, payment_history, users, audit_log)
            st.session_state.users = users
            st.session_state.audit_log = audit_log
            st.success("Password ተቀይሯል።")
            st.rerun()

    st.divider()

    st.subheader("📜 Audit Trail")

    if audit_log:
        audit_df = pd.DataFrame(audit_log)
        st.dataframe(
            audit_df.tail(200).iloc[::-1],
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("እስካሁን Audit Trail የለም።")


# ============================================================
# FOOTER
# ============================================================

st.sidebar.divider()

st.sidebar.info(
    f"👥 አባላት: {len(members)}\n\n"
    f"📅 ወራት: {len(MONTHS)}\n\n"
    f"💳 የመመዝገቢያ ክፍያ: {money(REGISTRATION_FEE)} ብር\n\n"
    f"📈 ወለድ: 2% reducing balance"
)

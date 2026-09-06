import streamlit as st
import pandas as pd
import os
import base64
import math
from io import StringIO, BytesIO
from datetime import datetime

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

SEED_VERSION = "fixed136-v1"

REGISTRATION_FEE = 500.0
LOAN_FEE_RATE = 0.10
INTEREST_RATE = 0.02

LOAN_TERMS = [3, 6, 12, 24, 36]


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


def normalize_member(member):
    """Ensure old and new records have all required fields."""

    for month in MONTHS:
        member[month] = safe_float(member.get(month, 0))

    member["የአባል ቁጥር"] = safe_int(
        member.get("የአባል ቁጥር", 0)
    )

    member["ስም"] = str(member.get("ስም", "") or "")
    member["ብሔራዊ መታወቂያ"] = str(
        member.get("ብሔራዊ መታወቂያ", "") or ""
    )

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
            "ብሔራዊ መታወቂያ": "",
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
# EXCEL SAVE
# ============================================================

def save_data_to_excel(members, payment_history=None):

    if payment_history is None:
        payment_history = []

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


# ============================================================
# EXCEL LOAD
# ============================================================

def load_database():

    if not os.path.exists(DB_FILE):

        members = create_initial_members()
        payment_history = []

        save_data_to_excel(
            members,
            payment_history
        )

        return members, payment_history

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

        return members, payment_history

    except Exception as e:

        st.error(
            f"Excel database ሲከፈት ችግር ተፈጥሯል፦ {e}"
        )

        return create_initial_members(), []


# ============================================================
# DUPLICATE NATIONAL ID
# ============================================================

def national_id_exists(
    members,
    national_id,
    exclude_member_number=None
):

    national_id = str(national_id).strip()

    if not national_id:
        return False

    for member in members:

        if (
            exclude_member_number is not None
            and safe_int(member.get("የአባል ቁጥር"))
            == safe_int(exclude_member_number)
        ):
            continue

        if str(
            member.get("ብሔራዊ መታወቂያ", "")
        ).strip() == national_id:

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

    members, payment_history = load_database()

    st.session_state.members = members
    st.session_state.payment_history = payment_history


members = st.session_state.members
payment_history = st.session_state.payment_history


# ============================================================
# HEADER
# ============================================================

st.title("🏦 ተስፋ የገንዘብ ቁጠባና ብድር ማህበር")

st.caption(
    "የቁጠባ፣ የብድር፣ የክፍያ እና የአባላት አስተዳደር ስርዓት"
)


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title("📋 ምናሌ")

page = st.sidebar.radio(
    "ገጽ ይምረጡ",
    [
        "👤 አባል መመዝገቢያ",
        "💰 የወር ቁጠባ ማስገቢያ",
        "💵 የብድር አገልግሎት",
        "📅 የብድር ክፍያ መመዝገቢያ",
        "📊 ጠቅላላ ሪፖርት",
        "✏️ የአባላት መረጃ ማስተካከያ"
    ]
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
            "ብሔራዊ መታወቂያ"
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

        elif national_id_exists(
            members,
            national_id
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
                "ብሔራዊ መታወቂያ": national_id.strip(),
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

            save_data_to_excel(
                members,
                payment_history
            )

            st.session_state.members = members

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

    member_names = [
        f'{safe_int(m.get("የአባል ቁጥር"))} - {m.get("ስም", "")}'
        for m in members
    ]

    if not member_names:

        st.warning("አባላት የሉም።")

    else:

        selected = st.selectbox(
            "አባል ይምረጡ",
            member_names
        )

        selected_number = int(
            selected.split(" - ")[0]
        )

        member = next(
            m for m in members
            if safe_int(m.get("የአባል ቁጥር"))
            == selected_number
        )

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

        if st.button(
            "💾 ቁጠባ መዝግብ",
            type="primary"
        ):

            if amount <= 0:

                st.error(
                    "የቁጠባ መጠን ከ0 በላይ ይሁን።"
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

                        save_data_to_excel(
                            members,
                            payment_history
                        )

                        st.session_state.members = members
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

                    save_data_to_excel(
                        members,
                        payment_history
                    )

                    st.session_state.members = members

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

    member_names = [
        f'{safe_int(m.get("የአባል ቁጥር"))} - {m.get("ስም", "")}'
        for m in members
    ]

    selected = st.selectbox(
        "አባል ይምረጡ",
        member_names,
        key="loan_member"
    )

    selected_number = int(
        selected.split(" - ")[0]
    )

    member = next(
        m for m in members
        if safe_int(m.get("የአባል ቁጥር"))
        == selected_number
    )

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

                save_data_to_excel(
                    members,
                    payment_history
                )

                st.session_state.members = members

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

    member_names = [
        f'{safe_int(m.get("የአባል ቁጥር"))} - {m.get("ስም", "")}'
        for m in members
    ]

    selected = st.selectbox(
        "አባል ይምረጡ",
        member_names,
        key="repayment_member"
    )

    selected_number = int(
        selected.split(" - ")[0]
    )

    member = next(
        m for m in members
        if safe_int(m.get("የአባል ቁጥር"))
        == selected_number
    )

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

            if st.button(
                "💾 ክፍያ መዝግብ",
                type="primary"
            ):

                result = apply_loan_payment(
                    member,
                    payment_amount
                )

                if result["success"]:

                    payment_history.append({
                        "የአባል ቁጥር": selected_number,
                        "ስም": member["ስም"],
                        "ቀን": datetime.now().strftime(
                            "%Y-%m-%d %H:%M:%S"
                        ),
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

                    save_data_to_excel(
                        members,
                        payment_history
                    )

                    st.session_state.members = members
                    st.session_state.payment_history = payment_history

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
    "🔎 አባል በስም ወይም በአባል ቁጥር ፈልግ",
    placeholder="ስም ወይም የአባል ቁጥር ያስገቡ..."
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

    member_names = [
        f'{safe_int(m.get("የአባል ቁጥር"))} - {m.get("ስም", "")}'
        for m in members
    ]

    if not member_names:

        st.warning("አባላት የሉም።")

    else:

        selected = st.selectbox(
            "አባል ይምረጡ",
            member_names,
            key="edit_member"
        )

        selected_number = int(
            selected.split(" - ")[0]
        )

        member_index = next(
            i
            for i, m in enumerate(members)
            if safe_int(
                m.get("የአባል ቁጥር")
            ) == selected_number
        )

        member = members[member_index]

        col1, col2 = st.columns(2)

        with col1:

            new_name = st.text_input(
                "ስም",
                value=str(
                    member.get("ስም", "")
                )
            )

            new_national_id = st.text_input(
                "ብሔራዊ መታወቂያ",
                value=str(
                    member.get(
                        "ብሔራዊ መታወቂያ",
                        ""
                    )
                )
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

            elif national_id_exists(
                members,
                new_national_id,
                exclude_member_number=selected_number
            ):

                st.error(
                    "ይህ ብሔራዊ መታወቂያ ሌላ አባል ላይ ተመዝግቧል።"
                )

            else:

                member["ስም"] = new_name.strip()

                member["ብሔራዊ መታወቂያ"] = (
                    new_national_id.strip()
                )

                member["ፎቶ"] = new_photo.strip()

                member["የመመዝገቢያ ክፍያ (ብር)"] = (
                    new_registration_fee
                )

                save_data_to_excel(
                    members,
                    payment_history
                )

                st.session_state.members = members

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
            "ይህን አባል መሰረዝ እፈልጋለሁ"
        )

        if confirm_delete:

            if st.button(
                "🗑️ አባል ሰርዝ",
                type="secondary"
            ):

                members.pop(member_index)

                save_data_to_excel(
                    members,
                    payment_history
                )

                st.session_state.members = members

                st.success(
                    "አባሉ ተሰርዟል።"
                )

                st.rerun()


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

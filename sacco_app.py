import streamlit as st
import pandas as pd
import os
import base64
import math
from io import StringIO, BytesIO
from datetime import datetime


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="SACCO Management System",
    page_icon="💰",
    layout="wide"
)


# ============================================================
# DATABASE CONFIGURATION
# ============================================================

DB_FILE = "sacco_database.xlsx"

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
# INITIAL 136 MEMBERS
# ============================================================

MEMBER_DATA = """
የአባል ቁጥር,የአባል ስም,መስከረም,ጥቅምት,ኅዳር
1,አባተ አለሙ,100,100,100
2,አበበ ጌታቸው,100,100,100
3,አበራ ተስፋዬ,100,100,100
4,አለማየሁ አሰፋ,100,100,100
5,አስራት አድማሱ,100,100,100
6,አዲሱ አሰፋ,100,100,100
7,አድማሱ ታደሰ,100,100,100
8,አለነ ገብረ,100,100,100
9,አማረ በቀለ,100,100,100
10,አሰፋ ሙሉጌታ,100,100,100
11,በላይ አለሙ,100,100,100
12,በቀለ ተስፋዬ,100,100,100
13,በዛብህ አለሙ,100,100,100
14,ብርሃኑ ጌታቸው,100,100,100
15,ብርቱካን አሰፋ,100,100,100
16,በረከት አለሙ,100,100,100
17,በድሉ ተስፋዬ,100,100,100
18,ቸርነት አለሙ,100,100,100
19,ዳዊት ጌታቸው,100,100,100
20,ደረጀ አሰፋ,100,100,100
21,ደስታ አለሙ,100,100,100
22,ደምሴ ተስፋዬ,100,100,100
23,ደምለው በቀለ,100,100,100
24,እሸቱ አለሙ,100,100,100
25,ፍቃዱ ጌታቸው,100,100,100
26,ፍሬው አሰፋ,100,100,100
27,ፍቅሩ ተስፋዬ,100,100,100
28,ገብረ እግዚአብሔር,100,100,100
29,ጌታቸው አለሙ,100,100,100
30,ግርማ ተስፋዬ,100,100,100
31,ሀብታሙ አሰፋ,100,100,100
32,ሀይሌ አለሙ,100,100,100
33,ህይወት በቀለ,100,100,100
34,ኃይሉ ጌታቸው,100,100,100
35,እንዳለ አለሙ,100,100,100
36,ኢያሱ ተስፋዬ,100,100,100
37,ካሳ አሰፋ,100,100,100
38,ከበደ አለሙ,100,100,100
39,ከፈለ ጌታቸው,100,100,100
40,ማሞ ተስፋዬ,100,100,100
41,መላኩ አለሙ,100,100,100
42,መኮንን አሰፋ,100,100,100
43,ሚካኤል ጌታቸው,100,100,100
44,ሙሉጌታ ተስፋዬ,100,100,100
45,ነጋሽ አለሙ,100,100,100
46,ንጉሴ አሰፋ,100,100,100
47,ኦስማን ጌታቸው,100,100,100
48,ፀጋዬ አለሙ,100,100,100
49,ተስፋዬ አሰፋ,100,100,100
50,ታምሩ ጌታቸው,100,100,100
51,ታደሰ አለሙ,100,100,100
52,ተፈራ ተስፋዬ,100,100,100
53,ተመስገን አሰፋ,100,100,100
54,ወልደሚካኤል ጌታቸው,100,100,100
55,ወንድሙ አለሙ,100,100,100
56,ዮሐንስ ተስፋዬ,100,100,100
57,ዘላለም አሰፋ,100,100,100
58,ዘውዱ ጌታቸው,100,100,100
59,እዮብ አለሙ,100,100,100
60,እድሉ ተስፋዬ,100,100,100
61,ሰለሞን አሰፋ,100,100,100
62,ሳሙኤል ጌታቸው,100,100,100
63,ሲሳይ አለሙ,100,100,100
64,ሽመልስ ተስፋዬ,100,100,100
65,ሹመት አሰፋ,100,100,100
66,ሽፈራው ጌታቸው,100,100,100
67,ሰንደቁ አለሙ,100,100,100
68,ሰይፈ ተስፋዬ,100,100,100
69,ሰውነት አሰፋ,100,100,100
70,ሰላም ጌታቸው,100,100,100
71,አስቴር አለሙ,100,100,100
72,ብርቱካን ተስፋዬ,100,100,100
73,ህዳሴ አሰፋ,100,100,100
74,ሙሉ ጌታቸው,100,100,100
75,ማርታ አለሙ,100,100,100
76,ሚስራቅ ተስፋዬ,100,100,100
77,ራሔል አሰፋ,100,100,100
78,ሳራ ጌታቸው,100,100,100
79,ሰላም አለሙ,100,100,100
80,ህይወት ተስፋዬ,100,100,100
81,ወርቅነሽ አሰፋ,100,100,100
82,ዘሪቱ ጌታቸው,100,100,100
83,እመቤት አለሙ,100,100,100
84,እህተ ተስፋዬ,100,100,100
85,እቴነሽ አሰፋ,100,100,100
86,ማህሌት ጌታቸው,100,100,100
87,ምህረት አለሙ,100,100,100
88,አልማዝ ተስፋዬ,100,100,100
89,ምስጋና አሰፋ,100,100,100
90,ፍሬህይወት ጌታቸው,100,100,100
91,እልፍነሽ አለሙ,100,100,100
92,ምህረት ተስፋዬ,100,100,100
93,ህይወት አሰፋ,100,100,100
94,ሰናይት ጌታቸው,100,100,100
95,ሰላማዊት አለሙ,100,100,100
96,ሳምራዊት ተስፋዬ,100,100,100
97,ሰናይት አሰፋ,100,100,100
98,ህዳሴ ጌታቸው,100,100,100
99,ሚሚ አለሙ,100,100,100
100,ባሳዝነው በላይ,100,100,100
101,አማኑኤል ተስፋዬ,100,100,100
102,ሙሉነህ አሰፋ,100,100,100
103,በረከት ጌታቸው,100,100,100
104,ታሪኩ አለሙ,100,100,100
105,ተስፋሁን ተስፋዬ,100,100,100
106,አብርሃም አሰፋ,100,100,100
107,አብይ ጌታቸው,100,100,100
108,ዮሴፍ አለሙ,100,100,100
109,ዳንኤል ተስፋዬ,100,100,100
110,አስማረ አሰፋ,100,100,100
111,አምሳሉ ጌታቸው,100,100,100
112,ዮናታን አለሙ,100,100,100
113,ተናኜ አንገሎ,100,100,100
114,ብሩክ ተስፋዬ,100,100,100
115,አሸናፊ አሰፋ,100,100,100
116,ሄኖክ ጌታቸው,100,100,100
117,ማርቆስ አለሙ,100,100,100
118,ናትናኤል ተስፋዬ,100,100,100
119,አርአያ አሰፋ,100,100,100
120,ሚካኤል ጌታቸው,100,100,100
121,ሙሉቀን አለሙ,100,100,100
122,በላይነህ ተስፋዬ,100,100,100
123,አለምነህ አሰፋ,100,100,100
124,አስራት ጌታቸው,100,100,100
125,አማረ አለሙ,100,100,100
126,ተሾመ ተስፋዬ,100,100,100
127,አለነ አሰፋ,100,100,100
128,ደረጀ ጌታቸው,100,100,100
129,ሙሉጌታ አለሙ,100,100,100
130,ተስፋዬ ተስፋዬ,100,100,100
131,አበበ አሰፋ,100,100,100
132,ብርሃኑ ጌታቸው,100,100,100
133,አዲሱ አለሙ,100,100,100
134,ፍጹም ፍስሃ,100,100,100
135,አስማረ ተስፋዬ,100,100,100
136,አብርሃም ጌታቸው,100,100,100
"""


# ============================================================
# CREATE INITIAL MEMBERS
# ============================================================

def create_initial_members():

    df = pd.read_csv(StringIO(MEMBER_DATA))

    members = []

    for _, row in df.iterrows():

        member_number = int(row["የአባል ቁጥር"])

        member = {
            "የአባል ቁጥር": member_number,
            "የአባል ስም": str(row["የአባል ስም"]).strip(),

            # Temporary National ID = Member Number
            "ብሔራዊ መታወቂያ": str(member_number),

            "የመመዝገቢያ ክፍያ": REGISTRATION_FEE,

            "የፎቶ ፋይል": "",

            "የብድር መጠን": 0.0,
            "የቀረ ብድር": 0.0,
            "የተከፈለ ወለድ": 0.0,
            "የተከፈለ ዋና ገንዘብ": 0.0
        }

        for month in MONTHS:

            if month in df.columns:
                member[month] = safe_float(row[month])
            else:
                member[month] = 0.0

        members.append(member)

    return members


# ============================================================
# NORMALIZE MEMBER
# ============================================================

def normalize_member(member):

    normalized = dict(member)

    required_fields = [
        "የአባል ቁጥር",
        "የአባል ስም",
        "ብሔራዊ መታወቂያ",
        "የመመዝገቢያ ክፍያ",
        "የፎቶ ፋይል",
        "የብድር መጠን",
        "የቀረ ብድር",
        "የተከፈለ ወለድ",
        "የተከፈለ ዋና ገንዘብ"
    ]

    for field in required_fields:

        if field not in normalized:
            normalized[field] = ""

    for month in MONTHS:

        if month not in normalized:
            normalized[month] = 0.0

        normalized[month] = safe_float(normalized[month])

    normalized["የአባል ቁጥር"] = safe_int(
        normalized["የአባል ቁጥር"]
    )

    normalized["የአባል ስም"] = str(
        normalized["የአባል ስም"]
    ).strip()

    normalized["ብሔራዊ መታወቂያ"] = str(
        normalized["ብሔራዊ መታወቂያ"]
    ).strip()

    normalized["የመመዝገቢያ ክፍያ"] = safe_float(
        normalized["የመመዝገቢያ ክፍያ"]
    )

    normalized["የብድር መጠን"] = safe_float(
        normalized["የብድር መጠን"]
    )

    normalized["የቀረ ብድር"] = safe_float(
        normalized["የቀረ ብድር"]
    )

    normalized["የተከፈለ ወለድ"] = safe_float(
        normalized["የተከፈለ ወለድ"]
    )

    normalized["የተከፈለ ዋና ገንዘብ"] = safe_float(
        normalized["የተከፈለ ዋና ገንዘብ"]
    )

    return normalized


# ============================================================
# SAFE CONVERSIONS
# ============================================================

def safe_float(value):

    try:

        if pd.isna(value):
            return 0.0

        if isinstance(value, str):

            value = value.replace(",", "").strip()

            if value == "":
                return 0.0

        return float(value)

    except Exception:

        return 0.0


def safe_int(value):

    try:

        if pd.isna(value):
            return 0

        return int(float(value))

    except Exception:

        return 0


def money(value):

    return f"{safe_float(value):,.2f}"


# ============================================================
# LOAN CALCULATIONS
# ============================================================

def calculate_monthly_payment(principal, months):

    principal = safe_float(principal)

    months = safe_int(months)

    if principal <= 0 or months <= 0:
        return 0.0

    monthly_rate = INTEREST_RATE

    if monthly_rate == 0:

        return principal / months

    payment = (
        principal
        * monthly_rate
        * (1 + monthly_rate) ** months
        /
        ((1 + monthly_rate) ** months - 1)
    )

    return payment


def calculate_current_interest(member):

    remaining = safe_float(
        member.get("የቀረ ብድር", 0)
    )

    if remaining <= 0:
        return 0.0

    return remaining * INTEREST_RATE


# ============================================================
# APPLY LOAN PAYMENT
# ============================================================

def apply_loan_payment(member, payment_amount):

    payment_amount = safe_float(payment_amount)

    remaining = safe_float(
        member.get("የቀረ ብድር", 0)
    )

    if payment_amount <= 0:
        return 0.0, 0.0

    if remaining <= 0:
        return 0.0, 0.0

    current_interest = calculate_current_interest(member)

    interest_paid = min(
        payment_amount,
        current_interest
    )

    remaining_payment = (
        payment_amount - interest_paid
    )

    principal_paid = min(
        remaining_payment,
        remaining
    )

    new_remaining = (
        remaining - principal_paid
    )

    member["የቀረ ብድር"] = max(
        0.0,
        new_remaining
    )

    member["የተከፈለ ወለድ"] = (
        safe_float(
            member.get("የተከፈለ ወለድ", 0)
        )
        + interest_paid
    )

    member["የተከፈለ ዋና ገንዘብ"] = (
        safe_float(
            member.get("የተከፈለ ዋና ገንዘብ", 0)
        )
        + principal_paid
    )

    return interest_paid, principal_paid


# ============================================================
# SAVE DATABASE
# ============================================================

def save_data_to_excel(
    members,
    payment_history=None
):

    members_df = pd.DataFrame(members)

    if payment_history is None:
        payment_history = []

    payment_history_df = pd.DataFrame(
        payment_history
    )

    with pd.ExcelWriter(
        DB_FILE,
        engine="openpyxl"
    ) as writer:

        members_df.to_excel(
            writer,
            sheet_name="አባላት",
            index=False
        )

        payment_history_df.to_excel(
            writer,
            sheet_name="የብድር_ክፍያ_ታሪክ",
            index=False
        )


# ============================================================
# DATABASE LOAD
# ============================================================

def load_database():

    # --------------------------------------------------------
    # No database yet
    # --------------------------------------------------------

    if not os.path.exists(DB_FILE):

        members = create_initial_members()

        payment_history = []

        save_data_to_excel(
            members,
            payment_history
        )

        return members, payment_history


    # --------------------------------------------------------
    # Existing database
    # --------------------------------------------------------

    try:

        excel_file = pd.ExcelFile(DB_FILE)

        # ------------------------------
        # Members
        # ------------------------------

        if "አባላት" in excel_file.sheet_names:

            members_df = pd.read_excel(
                DB_FILE,
                sheet_name="አባላት"
            )

            members = (
                members_df
                .fillna("")
                .to_dict("records")
            )

        else:

            members = create_initial_members()


        members = [
            normalize_member(member)
            for member in members
        ]


        # ------------------------------
        # Payment History
        # ------------------------------

        if "የብድር_ክፍያ_ታሪክ" in excel_file.sheet_names:

            payment_df = pd.read_excel(
                DB_FILE,
                sheet_name="የብድር_ክፍያ_ታሪክ"
            )

            payment_history = (
                payment_df
                .fillna("")
                .to_dict("records")
            )

        else:

            payment_history = []


        # ----------------------------------------------------
        # IMPORTANT:
        # Fill National ID for the original 136 members only
        # when their National ID is blank.
        #
        # Existing real National IDs are NOT overwritten.
        # ----------------------------------------------------

        changed = False

        seed_members = {}

        try:

            seed_df = pd.read_csv(
                StringIO(MEMBER_DATA)
            )

            for _, row in seed_df.iterrows():

                number = int(
                    row["የአባል ቁጥር"]
                )

                name = str(
                    row["የአባል ስም"]
                ).strip()

                seed_members[number] = name

        except Exception:

            seed_members = {}


        for member in members:

            member_number = safe_int(
                member.get(
                    "የአባል ቁጥር",
                    0
                )
            )

            current_id = str(
                member.get(
                    "ብሔራዊ መታወቂያ",
                    ""
                )
            ).strip()

            member_name = str(
                member.get(
                    "የአባል ስም",
                    ""
                )
            ).strip()


            # Only original seeded 136 members
            if (
                member_number in seed_members
                and member_name == seed_members[member_number]
                and current_id == ""
            ):

                member[
                    "ብሔራዊ መታወቂያ"
                ] = str(member_number)

                changed = True


        # ----------------------------------------------------
        # Save automatic National ID migration
        # ----------------------------------------------------

        if changed:

            save_data_to_excel(
                members,
                payment_history
            )


        return members, payment_history


    except Exception as e:

        st.error(
            f"የዳታቤዝ ፋይሉን ማንበብ አልተቻለም፦ {e}"
        )

        return [], []


# ============================================================
# NATIONAL ID DUPLICATION CHECK
# ============================================================

def national_id_exists(
    members,
    national_id,
    exclude_member_number=None
):

    national_id = str(
        national_id
    ).strip()

    if national_id == "":
        return False


    for member in members:

        member_number = safe_int(
            member.get(
                "የአባል ቁጥር",
                0
            )
        )

        if (
            exclude_member_number is not None
            and member_number == safe_int(
                exclude_member_number
            )
        ):
            continue


        existing_id = str(
            member.get(
                "ብሔራዊ መታወቂያ",
                ""
            )
        ).strip()


        if (
            existing_id != ""
            and existing_id == national_id
        ):

            return True


    return False


# ============================================================
# TOTAL SAVINGS
# ============================================================

def total_savings(member):

    total = 0.0

    for month in MONTHS:

        total += safe_float(
            member.get(month, 0)
        )

    return total


# ============================================================
# EXCEL DOWNLOAD
# ============================================================

def create_excel_download(
    members,
    payment_history
):

    output = BytesIO()

    members_df = pd.DataFrame(
        members
    )

    payment_history_df = pd.DataFrame(
        payment_history
    )

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        members_df.to_excel(
            writer,
            sheet_name="አባላት",
            index=False
        )

        payment_history_df.to_excel(
            writer,
            sheet_name="የብድር_ክፍያ_ታሪክ",
            index=False
        )

    return output.getvalue()


# ============================================================
# LOAD SESSION DATA
# ============================================================

if "members" not in st.session_state:

    (
        st.session_state.members,
        st.session_state.payment_history
    ) = load_database()


members = st.session_state.members

payment_history = (
    st.session_state.payment_history
)


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title("📋 SACCO ማውጫ")

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

    st.title("👤 አዲስ አባል መመዝገቢያ")

    st.info(
        "⚠️ የብሔራዊ መታወቂያ ቁጥር ማስገባት ግዴታ ነው። "
        "አንድ የብሔራዊ መታወቂያ ቁጥር ለሁለት አባላት መጠቀም አይቻልም።"
    )


    # --------------------------------------------------------
    # Next member number
    # --------------------------------------------------------

    existing_numbers = [
        safe_int(
            m.get(
                "የአባል ቁጥር",
                0
            )
        )
        for m in members
    ]

    if existing_numbers:

        next_member_number = (
            max(existing_numbers) + 1
        )

    else:

        next_member_number = 1


    st.write(
        f"**የሚሰጠው የአባል ቁጥር:** "
        f"{next_member_number}"
    )


    # --------------------------------------------------------
    # Registration Form
    # --------------------------------------------------------

    with st.form(
        "member_registration_form"
    ):

        name = st.text_input(
            "የአባል ስም *"
        )

        national_id = st.text_input(
            "ብሔራዊ መታወቂያ *",
            placeholder="የእውነተኛውን መታወቂያ ቁጥር ያስገቡ"
        )

        registration_fee = st.number_input(
            "የመመዝገቢያ ክፍያ",
            min_value=0.0,
            value=REGISTRATION_FEE,
            step=50.0
        )

        photo = st.file_uploader(
            "የአባል ፎቶ",
            type=[
                "jpg",
                "jpeg",
                "png"
            ]
        )

        submit = st.form_submit_button(
            "➕ አባል መዝግብ"
        )


    if submit:

        name = name.strip()
        national_id = national_id.strip()


        # ----------------------------------------------------
        # Validation
        # ----------------------------------------------------

        if name == "":

            st.error(
                "እባክዎ የአባል ስም ያስገቡ።"
            )

        elif national_id == "":

            st.error(
                "❌ የብሔራዊ መታወቂያ ቁጥር "
                "ማስገባት ግዴታ ነው።"
            )

        elif national_id_exists(
            members,
            national_id
        ):

            st.error(
                "❌ ይህ የብሔራዊ መታወቂያ "
                "ቁጥር ቀድሞ ለሌላ አባል ተመዝግቧል።"
            )

        else:

            # ------------------------------------------------
            # Photo
            # ------------------------------------------------

            photo_data = ""

            if photo is not None:

                photo_bytes = photo.read()

                photo_data = base64.b64encode(
                    photo_bytes
                ).decode("utf-8")


            # ------------------------------------------------
            # New Member
            # ------------------------------------------------

            new_member = {

                "የአባል ቁጥር":
                    next_member_number,

                "የአባል ስም":
                    name,

                "ብሔራዊ መታወቂያ":
                    national_id,

                "የመመዝገቢያ ክፍያ":
                    registration_fee,

                "የፎቶ ፋይል":
                    photo_data,

                "የብድር መጠን":
                    0.0,

                "የቀረ ብድር":
                    0.0,

                "የተከፈለ ወለድ":
                    0.0,

                "የተከፈለ ዋና ገንዘብ":
                    0.0
            }


            for month in MONTHS:

                new_member[month] = 0.0


            members.append(
                new_member
            )


            save_data_to_excel(
                members,
                payment_history
            )


            st.session_state.members = members

            st.success(
                f"✅ {name} በአባል ቁጥር "
                f"{next_member_number} ተመዝግቧል።"
            )

            st.rerun()


# ============================================================
# 2. MONTHLY SAVINGS
# ============================================================

elif page == "💰 የወር ቁጠባ ማስገቢያ":

    st.title("💰 የወር ቁጠባ ማስገቢያ")


    if not members:

        st.warning(
            "ምንም አባል አልተመዘገበም።"
        )

    else:

        member_options = {

            f"{m['የአባል ቁጥር']} - {m['የአባል ስም']}":
            m["የአባል ቁጥር"]

            for m in members
        }


        selected = st.selectbox(
            "አባል ይምረጡ",
            list(member_options.keys())
        )


        member_number = member_options[
            selected
        ]


        member = next(
            (
                m for m in members
                if safe_int(
                    m["የአባል ቁጥር"]
                ) == member_number
            ),
            None
        )


        if member:

            st.write(
                f"### {member['የአባል ስም']}"
            )

            st.write(
                f"**የአባል ቁጥር:** "
                f"{member_number}"
            )

            st.write(
                f"**ብሔራዊ መታወቂያ:** "
                f"{member.get('ብሔራዊ መታወቂያ', '')}"
            )


            month = st.selectbox(
                "ወር ይምረጡ",
                MONTHS
            )


            current_amount = safe_float(
                member.get(
                    month,
                    0
                )
            )


            amount = st.number_input(
                "የወሩ ቁጠባ",
                min_value=0.0,
                value=current_amount,
                step=10.0
            )


            if st.button(
                "💾 ቁጠባ አስቀምጥ"
            ):

                old_amount = current_amount

                new_amount = safe_float(
                    amount
                )

                difference = (
                    new_amount - old_amount
                )


                # ------------------------------------------------
                # If member has outstanding loan,
                # new payment first reduces loan.
                # ------------------------------------------------

                outstanding = safe_float(
                    member.get(
                        "የቀረ ብድር",
                        0
                    )
                )


                if (
                    difference > 0
                    and outstanding > 0
                ):

                    interest_paid, principal_paid = (
                        apply_loan_payment(
                            member,
                            difference
                        )
                    )


                    if (
                        interest_paid
                        + principal_paid
                        < difference
                    ):

                        extra = (
                            difference
                            -
                            interest_paid
                            -
                            principal_paid
                        )

                        member[month] = (
                            old_amount + extra
                        )

                    else:

                        member[month] = old_amount


                else:

                    member[month] = new_amount


                save_data_to_excel(
                    members,
                    payment_history
                )

                st.session_state.members = members

                st.success(
                    "✅ የወሩ ቁጠባ ተመዝግቧል።"
                )

                st.rerun()


# ============================================================
# 3. LOAN SERVICE
# ============================================================

elif page == "💵 የብድር አገልግሎት":

    st.title("💵 የብድር አገልግሎት")


    if not members:

        st.warning(
            "ምንም አባል የለም።"
        )

    else:

        member_options = {

            f"{m['የአባል ቁጥር']} - {m['የአባል ስም']}":
            m["የአባል ቁጥር"]

            for m in members
        }


        selected = st.selectbox(
            "አባል ይምረጡ",
            list(member_options.keys())
        )


        member_number = member_options[
            selected
        ]


        member = next(
            (
                m for m in members
                if safe_int(
                    m["የአባል ቁጥር"]
                ) == member_number
            ),
            None
        )


        if member:

            savings = total_savings(
                member
            )

            existing_debt = safe_float(
                member.get(
                    "የቀረ ብድር",
                    0
                )
            )


            max_loan = (
                savings * 4
            )


            st.metric(
                "ጠቅላላ ቁጠባ",
                f"{money(savings)} ETB"
            )

            st.metric(
                "ከፍተኛ የብድር መጠን",
                f"{money(max_loan)} ETB"
            )

            st.metric(
                "ያለው ብድር",
                f"{money(existing_debt)} ETB"
            )


            if existing_debt > 0:

                st.warning(
                    "⚠️ አባሉ አሁን ያልተከፈለ "
                    "ብድር ስላለበት አዲስ ብድር "
                    "መውሰድ አይችልም።"
                )

            else:

                loan_amount = st.number_input(
                    "የብድር መጠን",
                    min_value=0.0,
                    max_value=float(max_loan),
                    step=100.0
                )


                term = st.selectbox(
                    "የብድር ጊዜ",
                    LOAN_TERMS
                )


                loan_fee = (
                    loan_amount
                    * LOAN_FEE_RATE
                )

                cash_given = (
                    loan_amount
                    - loan_fee
                )


                monthly_payment = (
                    calculate_monthly_payment(
                        loan_amount,
                        term
                    )
                )


                st.write(
                    f"**የብድር አገልግሎት ክፍያ (10%):** "
                    f"{money(loan_fee)} ETB"
                )

                st.write(
                    f"**ለአባሉ የሚሰጠው ገንዘብ:** "
                    f"{money(cash_given)} ETB"
                )

                st.write(
                    f"**የወር ክፍያ:** "
                    f"{money(monthly_payment)} ETB"
                )


                if st.button(
                    "💵 ብድር መዝግብ"
                ):

                    if loan_amount <= 0:

                        st.error(
                            "የብድር መጠን ከዜሮ በላይ መሆን አለበት።"
                        )

                    elif loan_amount > max_loan:

                        st.error(
                            "የብድር መጠኑ ከተፈቀደው "
                            "ከፍተኛ መጠን በላይ ነው።"
                        )

                    else:

                        member["የብድር መጠን"] = (
                            loan_amount
                        )

                        member["የቀረ ብድር"] = (
                            loan_amount
                        )

                        member["የተከፈለ ወለድ"] = 0.0

                        member["የተከፈለ ዋና ገንዘብ"] = 0.0


                        save_data_to_excel(
                            members,
                            payment_history
                        )

                        st.session_state.members = members

                        st.success(
                            f"✅ {money(loan_amount)} ETB "
                            "ብድር ተመዝግቧል።"
                        )

                        st.rerun()


# ============================================================
# 4. LOAN REPAYMENT
# ============================================================

elif page == "📅 የብድር ክፍያ መመዝገቢያ":

    st.title("📅 የብድር ክፍያ መመዝገቢያ")


    if not members:

        st.warning(
            "ምንም አባል የለም።"
        )

    else:

        members_with_debt = [

            m for m in members

            if safe_float(
                m.get(
                    "የቀረ ብድር",
                    0
                )
            ) > 0
        ]


        if not members_with_debt:

            st.info(
                "አሁን ያልተከፈለ ብድር ያለበት አባል የለም።"
            )

        else:

            member_options = {

                f"{m['የአባል ቁጥር']} - {m['የአባል ስም']}":
                m["የአባል ቁጥር"]

                for m in members_with_debt
            }


            selected = st.selectbox(
                "አባል ይምረጡ",
                list(member_options.keys())
            )


            member_number = member_options[
                selected
            ]


            member = next(
                (
                    m for m in members
                    if safe_int(
                        m["የአባል ቁጥር"]
                    ) == member_number
                ),
                None
            )


            if member:

                remaining = safe_float(
                    member.get(
                        "የቀረ ብድር",
                        0
                    )
                )


                current_interest = (
                    calculate_current_interest(
                        member
                    )
                )


                st.metric(
                    "የቀረ ብድር",
                    f"{money(remaining)} ETB"
                )

                st.metric(
                    "የዚህ ወር ወለድ",
                    f"{money(current_interest)} ETB"
                )


                payment_amount = st.number_input(
                    "የክፍያ መጠን",
                    min_value=0.0,
                    max_value=float(
                        remaining
                        + current_interest
                    ),
                    step=50.0
                )


                payment_date = st.date_input(
                    "የክፍያ ቀን",
                    value=datetime.today()
                )


                if st.button(
                    "💾 ክፍያ መዝግብ"
                ):

                    if payment_amount <= 0:

                        st.error(
                            "የክፍያ መጠን ከዜሮ በላይ መሆን አለበት።"
                        )

                    else:

                        interest_paid, principal_paid = (
                            apply_loan_payment(
                                member,
                                payment_amount
                            )
                        )


                        history_record = {

                            "ቀን":
                                str(payment_date),

                            "የአባል ቁጥር":
                                member_number,

                            "የአባል ስም":
                                member["የአባል ስም"],

                            "ብሔራዊ መታወቂያ":
                                member.get(
                                    "ብሔራዊ መታወቂያ",
                                    ""
                                ),

                            "የተከፈለ ጠቅላላ":
                                payment_amount,

                            "የተከፈለ ወለድ":
                                interest_paid,

                            "የተከፈለ ዋና ገንዘብ":
                                principal_paid,

                            "የቀረ ብድር":
                                member["የቀረ ብድር"]
                        }


                        payment_history.append(
                            history_record
                        )


                        save_data_to_excel(
                            members,
                            payment_history
                        )


                        st.session_state.members = members

                        st.session_state.payment_history = (
                            payment_history
                        )


                        st.success(
                            "✅ የብድር ክፍያ ተመዝግቧል።"
                        )


                        st.write(
                            f"**የተከፈለ ወለድ:** "
                            f"{money(interest_paid)} ETB"
                        )

                        st.write(
                            f"**የተከፈለ ዋና ገንዘብ:** "
                            f"{money(principal_paid)} ETB"
                        )

                        st.write(
                            f"**የቀረ ብድር:** "
                            f"{money(member['የቀረ ብድር'])} ETB"
                        )


                        st.rerun()


# ============================================================
# 5. REPORT
# ============================================================

elif page == "📊 ጠቅላላ ሪፖርት":

    st.title("📊 ጠቅላላ ሪፖርት")


    total_members = len(
        members
    )


    total_savings_all = sum(
        total_savings(m)
        for m in members
    )


    total_loans = sum(
        safe_float(
            m.get(
                "የብድር መጠን",
                0
            )
        )
        for m in members
    )


    total_remaining_loans = sum(
        safe_float(
            m.get(
                "የቀረ ብድር",
                0
            )
        )
        for m in members
    )


    total_interest = sum(
        safe_float(
            m.get(
                "የተከፈለ ወለድ",
                0
            )
        )
        for m in members
    )


    total_registration_fees = sum(
        safe_float(
            m.get(
                "የመመዝገቢያ ክፍያ",
                0
            )
        )
        for m in members
    )


    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "👥 ጠቅላላ አባላት",
            total_members
        )

    with col2:

        st.metric(
            "💰 ጠቅላላ ቁጠባ",
            f"{money(total_savings_all)} ETB"
        )

    with col3:

        st.metric(
            "💵 ጠቅላላ የተሰጠ ብድር",
            f"{money(total_loans)} ETB"
        )


    col4, col5, col6 = st.columns(3)

    with col4:

        st.metric(
            "📉 የቀረ ብድር",
            f"{money(total_remaining_loans)} ETB"
        )

    with col5:

        st.metric(
            "📈 የተከፈለ ወለድ",
            f"{money(total_interest)} ETB"
        )

    with col6:

        st.metric(
            "🧾 የመመዝገቢያ ክፍያ",
            f"{money(total_registration_fees)} ETB"
        )


    st.divider()

    st.subheader(
        "👥 የአባላት ዝርዝር"
    )


    report_rows = []

    for member in members:

        report_rows.append({

            "የአባል ቁጥር":
                member["የአባል ቁጥር"],

            "የአባል ስም":
                member["የአባል ስም"],

            "ብሔራዊ መታወቂያ":
                member.get(
                    "ብሔራዊ መታወቂያ",
                    ""
                ),

            "ጠቅላላ ቁጠባ":
                total_savings(member),

            "የብድር መጠን":
                member.get(
                    "የብድር መጠን",
                    0
                ),

            "የቀረ ብድር":
                member.get(
                    "የቀረ ብድር",
                    0
                ),

            "የተከፈለ ወለድ":
                member.get(
                    "የተከፈለ ወለድ",
                    0
                )
        })


    report_df = pd.DataFrame(
        report_rows
    )


    st.dataframe(
        report_df,
        use_container_width=True,
        hide_index=True
    )


    st.divider()


    excel_data = create_excel_download(
        members,
        payment_history
    )


    st.download_button(
        label="📥 Excel ሪፖርት አውርድ",
        data=excel_data,
        file_name="SACCO_Report.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


# ============================================================
# 6. EDIT MEMBER INFORMATION
# ============================================================

elif page == "✏️ የአባላት መረጃ ማስተካከያ":

    st.title("✏️ የአባላት መረጃ ማስተካከያ")


    if not members:

        st.warning(
            "ምንም አባል አልተመዘገበም።"
        )

    else:

        member_options = {

            f"{m['የአባል ቁጥር']} - {m['የአባል ስም']}":
            m["የአባል ቁጥር"]

            for m in members
        }


        selected = st.selectbox(
            "አባል ይምረጡ",
            list(member_options.keys())
        )


        member_number = member_options[
            selected
        ]


        member_index = next(
            (
                i for i, m in enumerate(members)
                if safe_int(
                    m["የአባል ቁጥር"]
                ) == member_number
            ),
            None
        )


        if member_index is not None:

            member = members[
                member_index
            ]


            current_name = str(
                member.get(
                    "የአባል ስም",
                    ""
                )
            )

            current_id = str(
                member.get(
                    "ብሔራዊ መታወቂያ",
                    ""
                )
            )


            # ------------------------------------------------
            # Temporary ID Notice
            # ------------------------------------------------

            if (
                current_id
                == str(member_number)
            ):

                st.warning(
                    "⚠️ የሚታየው የብሔራዊ "
                    "መታወቂያ የአባል ቁጥር ነው። "
                    "ይህ ጊዜያዊ መረጃ ስለሆነ "
                    "እውነተኛውን ብሔራዊ መታወቂያ "
                    "ቁጥር በመተካት ያስቀምጡ።"
                )


            # ------------------------------------------------
            # Edit form
            # ------------------------------------------------

            with st.form(
                f"edit_member_{member_number}"
            ):

                new_name = st.text_input(
                    "የአባል ስም *",
                    value=current_name
                )


                new_national_id = st.text_input(
                    "ብሔራዊ መታወቂያ *",
                    value=current_id
                )


                new_registration_fee = st.number_input(
                    "የመመዝገቢያ ክፍያ",
                    min_value=0.0,
                    value=safe_float(
                        member.get(
                            "የመመዝገቢያ ክፍያ",
                            0
                        )
                    ),
                    step=50.0
                )


                new_photo = st.file_uploader(
                    "አዲስ ፎቶ ካለ",
                    type=[
                        "jpg",
                        "jpeg",
                        "png"
                    ]
                )


                save_changes = st.form_submit_button(
                    "💾 ለውጥ አስቀምጥ"
                )


            if save_changes:

                new_name = new_name.strip()

                new_national_id = (
                    new_national_id.strip()
                )


                if new_name == "":

                    st.error(
                        "የአባል ስም ባዶ መሆን አይችልም።"
                    )

                elif new_national_id == "":

                    st.error(
                        "❌ የብሔራዊ መታወቂያ "
                        "ቁጥር ማስገባት ግዴታ ነው።"
                    )

                elif national_id_exists(
                    members,
                    new_national_id,
                    exclude_member_number=member_number
                ):

                    st.error(
                        "❌ ይህ የብሔራዊ መታወቂያ "
                        "ቁጥር ለሌላ አባል ተመዝግቧል።"
                    )

                else:

                    member[
                        "የአባል ስም"
                    ] = new_name


                    member[
                        "ብሔራዊ መታወቂያ"
                    ] = new_national_id


                    member[
                        "የመመዝገቢያ ክፍያ"
                    ] = new_registration_fee


                    if new_photo is not None:

                        photo_bytes = (
                            new_photo.read()
                        )

                        member[
                            "የፎቶ ፋይል"
                        ] = base64.b64encode(
                            photo_bytes
                        ).decode("utf-8")


                    save_data_to_excel(
                        members,
                        payment_history
                    )


                    st.session_state.members = members

                    st.success(
                        "✅ የአባሉ መረጃ ተስተካክሏል።"
                    )

                    st.rerun()


            # ------------------------------------------------
            # Delete Member
            # ------------------------------------------------

            st.divider()

            st.subheader(
                "🗑️ አባል ሰርዝ"
            )

            st.warning(
                "⚠️ አባልን መሰረዝ የአባሉን "
                "መረጃ ከአባላት ዝርዝር ያስወግዳል። "
                "የብድር ክፍያ ታሪክ ግን አይሰረዝም።"
            )


            confirm_delete = st.checkbox(
                "ይህን አባል ለመሰረዝ አረጋግጣለሁ።"
            )


            if st.button(
                "🗑️ አባል ሰርዝ"
            ):

                if not confirm_delete:

                    st.error(
                        "እባክዎ የመሰረዝ ማረጋገጫን ይምረጡ።"
                    )

                else:

                    deleted_member = members.pop(
                        member_index
                    )


                    save_data_to_excel(
                        members,
                        payment_history
                    )


                    st.session_state.members = members

                    st.success(
                        f"✅ {deleted_member['የአባል ስም']} "
                        "ተሰርዟል።"
                    )

                    st.rerun()

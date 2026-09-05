import streamlit as st
import pandas as pd
import os
import base64
import math
from io import StringIO, BytesIO
from openpyxl import load_workbook

# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="የቁጠባ እና ብድር ሲስተም",
    layout="wide"
)

st.title("🏦 ተስፋ የገንዘብ ቁጠባና ብድር ማህበር")


# ============================================================
# DATABASE
# ============================================================

DB_FILE = "sacco_database.xlsx"

SEED_VERSION = "fixed136-v2"

MONTHS = [
    "ግንቦት 2018",
    "ሰኔ 2018",
    "ሐምሌ 2018"
]

LOAN_INTEREST_RATE = 0.02
LOAN_FEE_RATE = 0.10
REGISTRATION_FEE = 500.0

LOAN_TERMS = [3, 6, 12, 24, 36]

HISTORY_SHEET = "የብድር_ክፍያ_ታሪክ"


# ============================================================
# INITIAL 136 MEMBERS
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
# CREATE INITIAL MEMBERS
# ============================================================

def create_initial_members():

    df = pd.read_csv(
        StringIO(MEMBER_DATA.strip()),
        sep="\t",
        header=None,
        names=[
            "ID",
            "የአባል ስም",
            "ግንቦት 2018",
            "ሰኔ 2018",
            "ሐምሌ 2018"
        ]
    )

    members = {}

    for _, row in df.iterrows():

        member_id = str(int(row["ID"]))

        may = float(row["ግንቦት 2018"])
        june = float(row["ሰኔ 2018"])
        july = float(row["ሐምሌ 2018"])

        total = may + june + july

        members[member_id] = {
            "የአባል ስም": str(row["የአባል ስም"]).strip(),
            "National ID": "",
            "የአባል ፎቶ": "",

            "ግንቦት 2018": may,
            "ሰኔ 2018": june,
            "ሐምሌ 2018": july,

            "ጠቅላላ ቁጠባ (ብር)": total,

            "የመመዝገቢያ ክፍያ (ብር)": 0.0,

            "ብድር ሁኔታ": "የለም",

            "የተበደረው ጠቅላላ (ብር)": 0.0,

            "10% የብድር ክፍያ (ብር)": 0.0,

            "በእጅ የተሰጠ 90% (ብር)": 0.0,

            "የብድር ጊዜ (ወር)": 0,

            "የብድር ወርሃዊ ክፍያ (ብር)": 0.0,

            "የቀረው ዋና ብድር (ብር)": 0.0,

            "የቀረው ዕዳ (ብር)": 0.0,

            "የተከፈለ ወለድ (ብር)": 0.0,

            "የተከፈለ ዋና ብድር (ብር)": 0.0,

            "_seed_version": SEED_VERSION
        }

    return members


# ============================================================
# NORMALIZE MEMBER
# ============================================================

def normalize_member(member):

    defaults = {
        "የአባል ስም": "",
        "National ID": "",
        "የአባል ፎቶ": "",

        "ግንቦት 2018": 0.0,
        "ሰኔ 2018": 0.0,
        "ሐምሌ 2018": 0.0,

        "ጠቅላላ ቁጠባ (ብር)": 0.0,

        "የመመዝገቢያ ክፍያ (ብር)": 0.0,

        "ብድር ሁኔታ": "የለም",

        "የተበደረው ጠቅላላ (ብር)": 0.0,

        "10% የብድር ክፍያ (ብር)": 0.0,

        "በእጅ የተሰጠ 90% (ብር)": 0.0,

        "የብድር ጊዜ (ወር)": 0,

        "የብድር ወርሃዊ ክፍያ (ብር)": 0.0,

        "የቀረው ዋና ብድር (ብር)": 0.0,

        "የቀረው ዕዳ (ብር)": 0.0,

        "የተከፈለ ወለድ (ብር)": 0.0,

        "የተከፈለ ዋና ብድር (ብር)": 0.0
    }

    for key, value in defaults.items():

        if key not in member or pd.isna(member[key]):
            member[key] = value

    text_fields = [
        "የአባል ስም",
        "National ID",
        "የአባል ፎቶ",
        "ብድር ሁኔታ"
    ]

    for key in text_fields:

        if pd.isna(member[key]):
            member[key] = ""
        else:
            member[key] = str(member[key]).strip()

    numeric_fields = [
        "ግንቦት 2018",
        "ሰኔ 2018",
        "ሐምሌ 2018",
        "ጠቅላላ ቁጠባ (ብር)",
        "የመመዝገቢያ ክፍያ (ብር)",
        "የተበደረው ጠቅላላ (ብር)",
        "10% የብድር ክፍያ (ብር)",
        "በእጅ የተሰጠ 90% (ብር)",
        "የብድር ጊዜ (ወር)",
        "የብድር ወርሃዊ ክፍያ (ብር)",
        "የቀረው ዋና ብድር (ብር)",
        "የቀረው ዕዳ (ብር)",
        "የተከፈለ ወለድ (ብር)",
        "የተከፈለ ዋና ብድር (ብር)"
    ]

    for key in numeric_fields:

        try:
            member[key] = float(member[key])
        except Exception:
            member[key] = 0.0

    # Backward compatibility:
    # Old database has only "የቀረው ዕዳ".
    if (
        member["የቀረው ዋና ብድር (ብር)"] == 0
        and member["የቀረው ዕዳ (ብር)"] > 0
    ):
        member["የቀረው ዋና ብድር (ብር)"] = (
            member["የቀረው ዕዳ (ብር)"]
        )

    return member


# ============================================================
# LOAD PAYMENT HISTORY
# ============================================================

def load_payment_history():

    if not os.path.exists(DB_FILE):
        return []

    try:

        history_df = pd.read_excel(
            DB_FILE,
            sheet_name=HISTORY_SHEET
        )

        if history_df.empty:
            return []

        return history_df.to_dict(
            orient="records"
        )

    except Exception:
        return []


# ============================================================
# SAVE DATABASE + HISTORY
# ============================================================

def save_data_to_excel(
    db,
    payment_history=None
):

    if not db:
        return

    rows = []

    for member_id, member in db.items():

        row = dict(member)

        row["መታወቂያ ቁጥር (ID)"] = str(
            member_id
        )

        rows.append(row)

    df = pd.DataFrame(rows)

    first_columns = [
        "መታወቂያ ቁጥር (ID)",
        "የአባል ስም",
        "National ID",
        "የአባል ፎቶ",

        "ግንቦት 2018",
        "ሰኔ 2018",
        "ሐምሌ 2018",

        "ጠቅላላ ቁጠባ (ብር)",

        "የመመዝገቢያ ክፍያ (ብር)",

        "ብድር ሁኔታ",

        "የተበደረው ጠቅላላ (ብር)",

        "10% የብድር ክፍያ (ብር)",

        "በእጅ የተሰጠ 90% (ብር)",

        "የብድር ጊዜ (ወር)",

        "የብድር ወርሃዊ ክፍያ (ብር)",

        "የቀረው ዋና ብድር (ብር)",

        "የቀረው ዕዳ (ብር)",

        "የተከፈለ ወለድ (ብር)",

        "የተከፈለ ዋና ብድር (ብር)",

        "_seed_version"
    ]

    existing_columns = [
        col
        for col in first_columns
        if col in df.columns
    ]

    other_columns = [
        col
        for col in df.columns
        if col not in existing_columns
    ]

    df = df[
        existing_columns + other_columns
    ]

    if payment_history is None:

        payment_history = load_payment_history()

    history_df = pd.DataFrame(
        payment_history
    )

    # --------------------------------------------------------
    # Write both sheets
    # --------------------------------------------------------

    with pd.ExcelWriter(
        DB_FILE,
        engine="openpyxl",
        mode="w"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="አባላት",
            index=False
        )

        if history_df.empty:

            history_df = pd.DataFrame(
                columns=[
                    "ተ.ቁ",
                    "የአባል ID",
                    "የአባል ስም",
                    "የክፍያ ቀን",
                    "የክፍያ መጠን (ብር)",
                    "የተከፈለ ወለድ (ብር)",
                    "የተከፈለ ዋና ብድር (ብር)",
                    "የቀረው ዋና ብድር (ብር)",
                    "የቀረው ዕዳ (ብር)"
                ]
            )

        history_df.to_excel(
            writer,
            sheet_name=HISTORY_SHEET,
            index=False
        )


# ============================================================
# LOAD DATABASE
# ============================================================

def load_database():

    fixed_members = create_initial_members()

    if not os.path.exists(DB_FILE):

        save_data_to_excel(
            fixed_members,
            []
        )

        return fixed_members

    try:

        df = pd.read_excel(
            DB_FILE,
            sheet_name="አባላት",
            dtype={
                "መታወቂያ ቁጥር (ID)": str
            }
        )

        if (
            df.empty
            or
            "መታወቂያ ቁጥር (ID)"
            not in df.columns
        ):

            save_data_to_excel(
                fixed_members,
                load_payment_history()
            )

            return fixed_members

        df["መታወቂያ ቁጥር (ID)"] = (
            df["መታወቂያ ቁጥር (ID)"]
            .astype(str)
            .str.strip()
            .str.replace(
                r"\.0$",
                "",
                regex=True
            )
        )

        old_data = (
            df
            .set_index(
                "መታወቂያ ቁጥር (ID)"
            )
            .to_dict(
                orient="index"
            )
        )

    except Exception:

        save_data_to_excel(
            fixed_members,
            load_payment_history()
        )

        return fixed_members

    # --------------------------------------------------------
    # Check whether fixed 136 members have already been loaded
    # --------------------------------------------------------

    already_seeded = False

    if old_data:

        versions = []

        for member in old_data.values():

            version = member.get(
                "_seed_version",
                ""
            )

            if not pd.isna(version):

                versions.append(
                    str(version)
                )

        if (
            SEED_VERSION in versions
            or
            "fixed136-v1" in versions
        ):

            already_seeded = True

    # --------------------------------------------------------
    # FIRST RUN OF NEW VERSION
    # --------------------------------------------------------

    if not already_seeded:

        # Preserve old payment history.
        payment_history = load_payment_history()

        new_data = fixed_members

        save_data_to_excel(
            new_data,
            payment_history
        )

        return new_data

    # --------------------------------------------------------
    # Existing data
    # --------------------------------------------------------

    data = {}

    for member_id, member in old_data.items():

        member_id = str(
            member_id
        ).strip()

        member = normalize_member(
            member
        )

        # Upgrade seed version.
        member["_seed_version"] = (
            SEED_VERSION
        )

        data[member_id] = member

    return data


# ============================================================
# INITIALIZE DATABASE
# ============================================================

data = load_database()

payment_history = load_payment_history()


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def national_id_exists(
    national_id,
    exclude_member_id=None
):

    national_id = str(
        national_id or ""
    ).strip()

    if not national_id:
        return False

    for member_id, member in data.items():

        if (
            exclude_member_id is not None
            and
            str(member_id)
            ==
            str(exclude_member_id)
        ):
            continue

        existing_id = str(
            member.get(
                "National ID",
                ""
            )
            or ""
        ).strip()

        if (
            existing_id
            and
            existing_id == national_id
        ):
            return True

    return False


def calculate_monthly_payment(
    principal,
    annual_rate,
    months
):
    """
    Amortized payment formula.

    P = principal
    r = monthly interest rate
    n = number of months
    """

    principal = float(principal)
    r = float(annual_rate)
    n = int(months)

    if principal <= 0 or n <= 0:
        return 0.0

    if r == 0:
        return principal / n

    payment = (
        principal
        * r
        * ((1 + r) ** n)
        /
        (((1 + r) ** n) - 1)
    )

    return payment


def calculate_total_scheduled_interest(
    principal,
    months
):

    monthly_payment = calculate_monthly_payment(
        principal,
        LOAN_INTEREST_RATE,
        months
    )

    total_payment = (
        monthly_payment * months
    )

    return max(
        total_payment - principal,
        0
    )


def apply_loan_payment(
    member,
    payment_amount
):

    payment_amount = float(
        payment_amount
    )

    remaining_principal = float(
        member.get(
            "የቀረው ዋና ብድር (ብር)",
            member.get(
                "የቀረው ዕዳ (ብር)",
                0
            )
        )
    )

    if remaining_principal <= 0:

        member["ብድር ሁኔታ"] = "የለም"

        member[
            "የቀረው ዋና ብድር (ብር)"
        ] = 0.0

        member[
            "የቀረው ዕዳ (ብር)"
        ] = 0.0

        return {
            "interest_paid": 0.0,
            "principal_paid": 0.0,
            "used": 0.0,
            "excess": payment_amount,
            "old_principal": 0.0,
            "new_principal": 0.0
        }

    # --------------------------------------------------------
    # REDUCING BALANCE INTEREST
    # --------------------------------------------------------

    interest_due = (
        remaining_principal
        * LOAN_INTEREST_RATE
    )

    interest_paid = min(
        payment_amount,
        interest_due
    )

    after_interest = (
        payment_amount
        - interest_paid
    )

    principal_paid = min(
        after_interest,
        remaining_principal
    )

    new_principal = (
        remaining_principal
        - principal_paid
    )

    used = (
        interest_paid
        + principal_paid
    )

    excess = max(
        payment_amount - used,
        0
    )

    # --------------------------------------------------------
    # Update member
    # --------------------------------------------------------

    member[
        "የቀረው ዋና ብድር (ብር)"
    ] = round(
        max(new_principal, 0),
        2
    )

    member[
        "የቀረው ዕዳ (ብር)"
    ] = member[
        "የቀረው ዋና ብድር (ብር)"
    ]

    member[
        "የተከፈለ ወለድ (ብር)"
    ] = round(
        float(
            member.get(
                "የተከፈለ ወለድ (ብር)",
                0
            )
        )
        + interest_paid,
        2
    )

    member[
        "የተከፈለ ዋና ብድር (ብር)"
    ] = round(
        float(
            member.get(
                "የተከፈለ ዋና ብድር (ብር)",
                0
            )
        )
        + principal_paid,
        2
    )

    if new_principal <= 0.01:

        member[
            "የቀረው ዋና ብድር (ብር)"
        ] = 0.0

        member[
            "የቀረው ዕዳ (ብር)"
        ] = 0.0

        member[
            "ብድር ሁኔታ"
        ] = "የለም"

    else:

        member[
            "ብድር ሁኔታ"
        ] = "ያለበት"

    return {
        "interest_paid": interest_paid,
        "principal_paid": principal_paid,
        "used": used,
        "excess": excess,
        "old_principal": remaining_principal,
        "new_principal": max(
            new_principal,
            0
        )
    }


def add_payment_history(
    member_id,
    member,
    payment_amount,
    result
):

    payment_history.append({
        "ተ.ቁ": len(payment_history) + 1,

        "የአባል ID": str(
            member_id
        ),

        "የአባል ስም": member.get(
            "የአባል ስም",
            ""
        ),

        "የክፍያ ቀን": pd.Timestamp.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        ),

        "የክፍያ መጠን (ብር)": round(
            float(payment_amount),
            2
        ),

        "የተከፈለ ወለድ (ብር)": round(
            float(
                result["interest_paid"]
            ),
            2
        ),

        "የተከፈለ ዋና ብድር (ብር)": round(
            float(
                result["principal_paid"]
            ),
            2
        ),

        "የቀረው ዋና ብድር (ብር)": round(
            float(
                result["new_principal"]
            ),
            2
        ),

        "የቀረው ዕዳ (ብር)": round(
            float(
                result["new_principal"]
            ),
            2
        )
    })


# ============================================================
# SIDEBAR MENU
# ============================================================

menu = st.sidebar.selectbox(
    "ያሉ አማራጮች",
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

if menu == "👤 አባል መመዝገቢያ":

    st.header(
        "👤 አዲስ አባል መመዝገቢያ"
    )

    st.info(
        "📌 አዲስ አባል ሲመዘገብ "
        "የመመዝገቢያ ክፍያ 500 ብር ነው። "
        "ይህ ከብድር 10% ክፍያ የተለየ ነው።"
    )

    with st.form(
        "member_registration_form"
    ):

        m_id = st.text_input(
            "የአባል ID:",
            placeholder="ለምሳሌ 137"
        )

        m_name = st.text_input(
            "የአባል ስም:"
        )

        m_national_id = st.text_input(
            "National ID (አማራጭ):"
        )

        m_photo = st.file_uploader(
            "የአባል ፎቶ (አማራጭ):",
            type=[
                "jpg",
                "jpeg",
                "png"
            ]
        )

        register = st.form_submit_button(
            "✅ አባል መዝግብ"
        )

    if register:

        m_id = str(m_id).strip()

        m_name = str(
            m_name
        ).strip()

        m_national_id = str(
            m_national_id or ""
        ).strip()

        if not m_id:

            st.error(
                "❌ ID ያስገቡ።"
            )

        elif not m_name:

            st.error(
                "❌ የአባል ስም ያስገቡ።"
            )

        elif m_id in data:

            st.error(
                f"❌ ID {m_id} ቀድሞ "
                f"ተመዝግቧል።"
            )

        elif (
            m_national_id
            and
            national_id_exists(
                m_national_id
            )
        ):

            st.error(
                "❌ ይህ National ID "
                "ቀድሞ ለሌላ አባል "
                "ተመዝግቧል።"
            )

        else:

            photo_data = ""

            if m_photo is not None:

                photo_bytes = m_photo.read()

                photo_data = (
                    base64.b64encode(
                        photo_bytes
                    )
                    .decode("utf-8")
                )

            data[m_id] = {

                "የአባል ስም": m_name,

                "National ID": m_national_id,

                "የአባል ፎቶ": photo_data,

                "ግንቦት 2018": 0.0,
                "ሰኔ 2018": 0.0,
                "ሐምሌ 2018": 0.0,

                "ጠቅላላ ቁጠባ (ብር)": 0.0,

                "የመመዝገቢያ ክፍያ (ብር)": (
                    REGISTRATION_FEE
                ),

                "ብድር ሁኔታ": "የለም",

                "የተበደረው ጠቅላላ (ብር)": 0.0,

                "10% የብድር ክፍያ (ብር)": 0.0,

                "በእጅ የተሰጠ 90% (ብር)": 0.0,

                "የብድር ጊዜ (ወር)": 0,

                "የብድር ወርሃዊ ክፍያ (ብር)": 0.0,

                "የቀረው ዋና ብድር (ብር)": 0.0,

                "የቀረው ዕዳ (ብር)": 0.0,

                "የተከፈለ ወለድ (ብር)": 0.0,

                "የተከፈለ ዋና ብድር (ብር)": 0.0,

                "_seed_version": SEED_VERSION
            }

            save_data_to_excel(
                data,
                payment_history
            )

            st.success(
                f"✅ {m_name} "
                f"በተሳካ ሁኔታ ተመዝግቧል።"
            )

            st.info(
                "🧾 የመመዝገቢያ ክፍያ: "
                f"{REGISTRATION_FEE:,.2f} ብር"
            )

            st.rerun()


# ============================================================
# 2. MONTHLY SAVINGS
# ============================================================

elif menu == "💰 የወር ቁጠባ ማስገቢያ":

    st.header(
        "💰 የወርሃዊ ቁጠባ መመዝገቢያ"
    )

    member_ids = list(
        data.keys()
    )

    member_id = st.selectbox(
        "አባል ይምረጡ:",
        member_ids,
        format_func=lambda x:
            f"{x} - {data[x]['የአባል ስም']}"
    )

    member = data[member_id]

    st.info(
        f"👤 {member['የአባል ስም']}\n\n"
        f"💰 አሁን ያለ ቁጠባ: "
        f"{member['ጠቅላላ ቁጠባ (ብር)']:,.2f} ብር"
    )

    month = st.selectbox(
        "ወር ይምረጡ:",
        [
            "ግንቦት 2018",
            "ሰኔ 2018",
            "ሐምሌ 2018",
            "ነሐሴ 2018",
            "መስከረም 2019",
            "ጥቅምት 2019",
            "ሌላ ወር"
        ]
    )

    if month == "ሌላ ወር":

        month = st.text_input(
            "የወሩን ስም ያስገቡ:"
        )

    existing_amount = 0.0

    if month:

        existing_amount = float(
            member.get(
                month,
                0
            )
            or 0
        )

    # --------------------------------------------------------
    # EXISTING MONTH WARNING
    # --------------------------------------------------------

    if (
        month
        and
        existing_amount > 0
    ):

        st.warning(
            f"⚠️ {month} ላይ "
            f"{existing_amount:,.2f} ብር "
            f"ቀድሞ ተመዝግቧል።\n\n"
            "እንደገና ከሚያስገቡ ይጠንቀቁ። "
            "አሁኑን መጠን መተካት "
            "ወይም ተጨማሪ ቁጠባ መጨመር ይችላሉ።"
        )

        operation = st.radio(
            "ምን ማድረግ ይፈልጋሉ?",
            [
                "➕ ተጨማሪ ቁጠባ ጨምር",
                "✏️ የነበረውን መጠን አስተካክል"
            ],
            horizontal=True
        )

    else:

        operation = (
            "➕ ተጨማሪ ቁጠባ ጨምር"
        )

    if operation == "✏️ የነበረውን መጠን አስተካክል":

        amount = st.number_input(
            f"አዲሱ {month} ቁጠባ (ብር):",
            min_value=0.0,
            value=float(existing_amount),
            step=100.0
        )

    else:

        amount = st.number_input(
            "የቁጠባ / ክፍያ መጠን (ብር):",
            min_value=0.0,
            step=100.0
        )

    if st.button(
        "💾 ቁጠባ መዝግብ"
    ):

        if amount <= 0:

            st.error(
                "❌ የገንዘብ መጠን ያስገቡ።"
            )

        elif not month:

            st.error(
                "❌ ወር ያስገቡ።"
            )

        # ----------------------------------------------------
        # EDIT EXISTING SAVINGS
        # ----------------------------------------------------

        elif operation == "✏️ የነበረውን መጠን አስተካክል":

            old_value = float(
                member.get(
                    month,
                    0
                )
                or 0
            )

            new_value = float(
                amount
            )

            difference = (
                new_value - old_value
            )

            member[month] = new_value

            member[
                "ጠቅላላ ቁጠባ (ብር)"
            ] = max(
                float(
                    member.get(
                        "ጠቅላላ ቁጠባ (ብር)",
                        0
                    )
                )
                + difference,
                0
            )

            save_data_to_excel(
                data,
                payment_history
            )

            st.success(
                f"✅ {month} ቁጠባ "
                f"ከ{old_value:,.2f} ወደ "
                f"{new_value:,.2f} ብር "
                "ተስተካክሏል።"
            )

            st.rerun()

        # ----------------------------------------------------
        # ADD SAVINGS
        # ----------------------------------------------------

        else:

            # ------------------------------------------------
            # If member has outstanding loan,
            # money goes to loan first.
            # ------------------------------------------------

            if (
                float(
                    member.get(
                        "የቀረው ዕዳ (ብር)",
                        0
                    )
                ) > 0
            ):

                result = apply_loan_payment(
                    member,
                    amount
                )

                used_for_loan = result[
                    "used"
                ]

                excess = result[
                    "excess"
                ]

                # Only excess becomes savings.
                if excess > 0:

                    member[month] = (
                        float(
                            member.get(
                                month,
                                0
                            )
                            or 0
                        )
                        + excess
                    )

                    member[
                        "ጠቅላላ ቁጠባ (ብር)"
                    ] += excess

                # Add loan payment history.
                add_payment_history(
                    member_id,
                    member,
                    used_for_loan,
                    result
                )

                save_data_to_excel(
                    data,
                    payment_history
                )

                st.success(
                    f"✅ {used_for_loan:,.2f} ብር "
                    "ለብድር ክፍያ ተጠቅሟል።"
                )

                if excess > 0:

                    st.info(
                        f"💰 {excess:,.2f} ብር "
                        f"እንደ ቁጠባ "
                        f"በ{month} ተጨምሯል።"
                    )

                st.info(
                    f"💳 የቀረው ዕዳ: "
                    f"{member['የቀረው ዕዳ (ብር)']:,.2f} ብር"
                )

            else:

                old_value = float(
                    member.get(
                        month,
                        0
                    )
                    or 0
                )

                member[month] = (
                    old_value
                    + float(amount)
                )

                member[
                    "ጠቅላላ ቁጠባ (ብር)"
                ] += float(amount)

                save_data_to_excel(
                    data,
                    payment_history
                )

                st.success(
                    f"✅ {amount:,.2f} ብር "
                    f"በ{month} ተመዝግቧል።"
                )

            st.rerun()


# ============================================================
# 3. LOAN SERVICE
# ============================================================

elif menu == "💵 የብድር አገልግሎት":

    st.header(
        "💵 የብድር አገልግሎት"
    )

    member_ids = list(
        data.keys()
    )

    member_id = st.selectbox(
        "አባል ይምረጡ:",
        member_ids,
        format_func=lambda x:
            f"{x} - {data[x]['የአባል ስም']}"
    )

    member = data[member_id]

    savings = float(
        member.get(
            "ጠቅላላ ቁጠባ (ብር)",
            0
        )
    )

    max_loan = savings * 4

    st.write(
        f"👤 **አባል:** "
        f"{member['የአባል ስም']}"
    )

    st.write(
        f"💰 **ጠቅላላ ቁጠባ:** "
        f"{savings:,.2f} ብር"
    )

    st.write(
        f"🏦 **ከፍተኛ የብድር መጠን (4×):** "
        f"{max_loan:,.2f} ብር"
    )

    if (
        member.get(
            "የቀረው ዕዳ (ብር)",
            0
        )
        > 0
    ):

        st.warning(
            f"⚠️ ይህ አባል ቀድሞ "
            f"{member['የቀረው ዕዳ (ብር)']:,.2f} "
            "ብር ዕዳ አለበት። "
            "አዲስ ብድር መውሰድ አይችልም።"
        )

    else:

        loan_amount = st.number_input(
            "የሚፈልጉት የብድር መጠን (ብር):",
            min_value=0.0,
            max_value=(
                max_loan
                if max_loan > 0
                else 0.0
            ),
            step=100.0
        )

        loan_term = st.selectbox(
            "የብድር መክፈያ ጊዜ:",
            LOAN_TERMS,
            format_func=lambda x:
                f"{x} ወራት"
        )

        if loan_amount > 0:

            # ------------------------------------------------
            # 10% LOAN FEE
            # ------------------------------------------------

            loan_fee = (
                loan_amount
                * LOAN_FEE_RATE
            )

            net_payout = (
                loan_amount
                - loan_fee
            )

            # ------------------------------------------------
            # REDUCING BALANCE AMORTIZATION
            # ------------------------------------------------

            monthly_payment = (
                calculate_monthly_payment(
                    loan_amount,
                    LOAN_INTEREST_RATE,
                    loan_term
                )
            )

            total_scheduled_interest = (
                calculate_total_scheduled_interest(
                    loan_amount,
                    loan_term
                )
            )

            total_scheduled_payment = (
                loan_amount
                + total_scheduled_interest
            )

            first_month_interest = (
                loan_amount
                * LOAN_INTEREST_RATE
            )

            first_month_principal = max(
                monthly_payment
                - first_month_interest,
                0
            )

            st.divider()

            st.subheader(
                "📋 የብድር ማጠቃለያ"
            )

            col1, col2, col3, col4 = (
                st.columns(4)
            )

            with col1:

                st.metric(
                    "የብድር መጠን",
                    f"{loan_amount:,.2f} ብር"
                )

            with col2:

                st.metric(
                    "10% የብድር ክፍያ",
                    f"{loan_fee:,.2f} ብር"
                )

            with col3:

                st.metric(
                    "በእጅ የሚሰጥ 90%",
                    f"{net_payout:,.2f} ብር"
                )

            with col4:

                st.metric(
                    "የብድር ጊዜ",
                    f"{loan_term} ወር"
                )

            st.divider()

            c1, c2, c3 = st.columns(3)

            with c1:

                st.write(
                    f"📈 **ወርሃዊ ወለድ:** "
                    f"2% በቀረው ዋና ብድር ላይ"
                )

            with c2:

                st.write(
                    f"📊 **የመጀመሪያ ወር ወለድ:** "
                    f"{first_month_interest:,.2f} ብር"
                )

            with c3:

                st.write(
                    f"💳 **ወርሃዊ ክፍያ:** "
                    f"**{monthly_payment:,.2f} ብር**"
                )

            st.write(
                f"📌 የመጀመሪያ ወር ዋና ብድር "
                f"ክፍያ: "
                f"{first_month_principal:,.2f} ብር"
            )

            st.write(
                f"📈 በተመረጠው {loan_term} ወር "
                f"ሙሉ በሙሉ ከተከፈለ "
                f"የሚጠበቀው ጠቅላላ ወለድ: "
                f"**{total_scheduled_interest:,.2f} ብር**"
            )

            st.write(
                f"💰 የተጠበቀ ጠቅላላ ክፍያ: "
                f"**{total_scheduled_payment:,.2f} ብር**"
            )

            st.info(
                "ℹ️ ይህ የ2% ወለድ በቀረው "
                "ዋና ብድር ላይ የሚሰላ "
                "reducing-balance ስርዓት ነው። "
                "አባሉ ቀድሞ ከከፈለ የወደፊት "
                "ወለድ ይቀንሳል።"
            )

            st.warning(
                f"⚠️ የ10% የብድር ክፍያ "
                f"({loan_fee:,.2f} ብር) "
                "ከ500 ብር የመመዝገቢያ "
                "ክፍያ የተለየ ነው።"
            )

            if st.button(
                "💵 ብድር ስጥ"
            ):

                if loan_amount > max_loan:

                    st.error(
                        "❌ ከተፈቀደው "
                        "የብድር መጠን በላይ ነው።"
                    )

                elif savings <= 0:

                    st.error(
                        "❌ ቁጠባ ሳይኖር "
                        "ብድር መውሰድ አይቻልም።"
                    )

                else:

                    member[
                        "ብድር ሁኔታ"
                    ] = "ያለበት"

                    member[
                        "የተበደረው ጠቅላላ (ብር)"
                    ] = loan_amount

                    member[
                        "10% የብድር ክፍያ (ብር)"
                    ] = loan_fee

                    member[
                        "በእጅ የተሰጠ 90% (ብር)"
                    ] = net_payout

                    member[
                        "የብድር ጊዜ (ወር)"
                    ] = loan_term

                    member[
                        "የብድር ወርሃዊ ክፍያ (ብር)"
                    ] = monthly_payment

                    member[
                        "የቀረው ዋና ብድር (ብር)"
                    ] = loan_amount

                    member[
                        "የቀረው ዕዳ (ብር)"
                    ] = loan_amount

                    member[
                        "የተከፈለ ወለድ (ብር)"
                    ] = 0.0

                    member[
                        "የተከፈለ ዋና ብድር (ብር)"
                    ] = 0.0

                    save_data_to_excel(
                        data,
                        payment_history
                    )

                    st.success(
                        f"🎉 ለ{member['የአባል ስም']} "
                        f"{loan_amount:,.2f} ብር "
                        "ብድር ተፈቅዷል!"
                    )

                    st.info(
                        f"💵 በእጅ የሚሰጠው 90%: "
                        f"{net_payout:,.2f} ብር\n\n"
                        f"🧾 10% የብድር ክፍያ: "
                        f"{loan_fee:,.2f} ብር\n\n"
                        f"📅 የብድር ጊዜ: "
                        f"{loan_term} ወራት\n\n"
                        f"💳 ወርሃዊ ክፍያ: "
                        f"{monthly_payment:,.2f} ብር"
                    )

                    st.rerun()


# ============================================================
# 4. LOAN REPAYMENT + HISTORY
# ============================================================

elif menu == "📅 የብድር ክፍያ መመዝገቢያ":

    st.header(
        "📅 የብድር ክፍያ መመዝገቢያ"
    )

    member_ids_with_loan = [
        member_id
        for member_id, member in data.items()
        if float(
            member.get(
                "የቀረው ዕዳ (ብር)",
                0
            )
        ) > 0
    ]

    if not member_ids_with_loan:

        st.info(
            "📌 በአሁኑ ጊዜ "
            "ያለ ብድር አባል የለም።"
        )

    else:

        member_id = st.selectbox(
            "የሚከፍለውን አባል ይምረጡ:",
            member_ids_with_loan,
            format_func=lambda x:
                f"{x} - {data[x]['የአባል ስም']}"
        )

        member = data[member_id]

        remaining_principal = float(
            member.get(
                "የቀረው ዋና ብድር (ብር)",
                member.get(
                    "የቀረው ዕዳ (ብር)",
                    0
                )
            )
        )

        current_interest = (
            remaining_principal
            * LOAN_INTEREST_RATE
        )

        monthly_payment = float(
            member.get(
                "የብድር ወርሃዊ ክፍያ (ብር)",
                0
            )
        )

        st.warning(
            f"👤 {member['የአባል ስም']}\n\n"
            f"💳 የቀረው ዋና ብድር: "
            f"{remaining_principal:,.2f} ብር\n\n"
            f"📈 የአሁኑ 2% ወለድ: "
            f"{current_interest:,.2f} ብር"
        )

        if monthly_payment > 0:

            st.info(
                f"📅 የተወሰነው ወርሃዊ ክፍያ: "
                f"{monthly_payment:,.2f} ብር"
            )

        payment = st.number_input(
            "የክፍያ መጠን (ብር):",
            min_value=0.0,
            step=100.0
        )

        if st.button(
            "💾 ክፍያ መዝግብ"
        ):

            if payment <= 0:

                st.error(
                    "❌ የክፍያ መጠን ያስገቡ።"
                )

            else:

                result = apply_loan_payment(
                    member,
                    payment
                )

                # Record history.
                add_payment_history(
                    member_id,
                    member,
                    payment - result["excess"],
                    result
                )

                save_data_to_excel(
                    data,
                    payment_history
                )

                st.success(
                    f"✅ ክፍያው ተመዝግቧል።"
                )

                c1, c2, c3 = st.columns(3)

                with c1:

                    st.metric(
                        "የተከፈለ ወለድ",
                        f"{result['interest_paid']:,.2f} ብር"
                    )

                with c2:

                    st.metric(
                        "የተከፈለ ዋና",
                        f"{result['principal_paid']:,.2f} ብር"
                    )

                with c3:

                    st.metric(
                        "የቀረ ዋና",
                        f"{result['new_principal']:,.2f} ብር"
                    )

                if result["excess"] > 0:

                    st.info(
                        f"💰 {result['excess']:,.2f} ብር "
                        "ከዕዳው በላይ ተከፍሏል። "
                        "ወደ ቁጠባ አልጨመርነውም፤ "
                        "በተለየ እንዲመዘገብ ተወስኗል።"
                    )

                if (
                    result["new_principal"]
                    <= 0.01
                ):

                    st.success(
                        "🎉 የአባሉ ብድር "
                        "ሙሉ በሙሉ ተከፍሏል።"
                    )

                st.rerun()

        # ----------------------------------------------------
        # MEMBER PAYMENT HISTORY
        # ----------------------------------------------------

        st.divider()

        st.subheader(
            "📜 የዚህ አባል የብድር ክፍያ ታሪክ"
        )

        member_history = [
            row
            for row in payment_history
            if str(
                row.get(
                    "የአባል ID",
                    ""
                )
            )
            ==
            str(member_id)
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
                "📌 ለዚህ አባል "
                "የብድር ክፍያ ታሪክ የለም።"
            )


# ============================================================
# 5. REPORT
# ============================================================

elif menu == "📊 ጠቅላላ ሪፖርት":

    st.header(
        "📊 ጠቅላላ የአባላት፣ "
        "የቁጠባ እና የብድር ሪፖርት"
    )

    if data:

        rows = []

        for member_id, member in data.items():

            row = {

                "ID": str(member_id),

                "የአባል ስም": member.get(
                    "የአባል ስም",
                    ""
                ),

                "National ID": member.get(
                    "National ID",
                    ""
                ),

                "ግንቦት 2018": float(
                    member.get(
                        "ግንቦት 2018",
                        0
                    )
                ),

                "ሰኔ 2018": float(
                    member.get(
                        "ሰኔ 2018",
                        0
                    )
                ),

                "ሐምሌ 2018": float(
                    member.get(
                        "ሐምሌ 2018",
                        0
                    )
                ),

                "ጠቅላላ ቁጠባ (ብር)": float(
                    member.get(
                        "ጠቅላላ ቁጠባ (ብር)",
                        0
                    )
                ),

                "የመመዝገቢያ ክፍያ (ብር)": float(
                    member.get(
                        "የመመዝገቢያ ክፍያ (ብር)",
                        0
                    )
                ),

                "ብድር ሁኔታ": member.get(
                    "ብድር ሁኔታ",
                    "የለም"
                ),

                "የተበደረው ጠቅላላ (ብር)": float(
                    member.get(
                        "የተበደረው ጠቅላላ (ብር)",
                        0
                    )
                ),

                "10% የብድር ክፍያ (ብር)": float(
                    member.get(
                        "10% የብድር ክፍያ (ብር)",
                        0
                    )
                ),

                "በእጅ የተሰጠ 90% (ብር)": float(
                    member.get(
                        "በእጅ የተሰጠ 90% (ብር)",
                        0
                    )
                ),

                "የብድር ጊዜ (ወር)": int(
                    float(
                        member.get(
                            "የብድር ጊዜ (ወር)",
                            0
                        )
                    )
                ),

                "የብድር ወርሃዊ ክፍያ (ብር)": float(
                    member.get(
                        "የብድር ወርሃዊ ክፍያ (ብር)",
                        0
                    )
                ),

                "የቀረው ዋና ብድር (ብር)": float(
                    member.get(
                        "የቀረው ዋና ብድር (ብር)",
                        member.get(
                            "የቀረው ዕዳ (ብር)",
                            0
                        )
                    )
                ),

                "የቀረው ዕዳ (ብር)": float(
                    member.get(
                        "የቀረው ዕዳ (ብር)",
                        0
                    )
                ),

                "የተከፈለ ወለድ (ብር)": float(
                    member.get(
                        "የተከፈለ ወለድ (ብር)",
                        0
                    )
                ),

                "የተከፈለ ዋና ብድር (ብር)": float(
                    member.get(
                        "የተከፈለ ዋና ብድር (ብር)",
                        0
                    )
                )
            }

            rows.append(row)

        report_df = pd.DataFrame(
            rows
        )

        report_df["_sort_id"] = pd.to_numeric(
            report_df["ID"],
            errors="coerce"
        )

        report_df = (
            report_df
            .sort_values("_sort_id")
            .drop(
                columns=["_sort_id"]
            )
        )

        # ----------------------------------------------------
        # SUMMARY
        # ----------------------------------------------------

        total_members = len(data)

        total_savings = report_df[
            "ጠቅላላ ቁጠባ (ብር)"
        ].sum()

        total_loan = report_df[
            "የተበደረው ጠቅላላ (ብር)"
        ].sum()

        total_debt = report_df[
            "የቀረው ዕዳ (ብር)"
        ].sum()

        total_loan_fee = report_df[
            "10% የብድር ክፍያ (ብር)"
        ].sum()

        total_interest = report_df[
            "የተከፈለ ወለድ (ብር)"
        ].sum()

        total_registration_fee = report_df[
            "የመመዝገቢያ ክፍያ (ብር)"
        ].sum()

        cooperative_income = (
            total_loan_fee
            + total_interest
            + total_registration_fee
        )

        # ----------------------------------------------------
        # MAIN METRICS
        # ----------------------------------------------------

        c1, c2, c3, c4 = st.columns(4)

        with c1:

            st.metric(
                "👥 አባላት",
                f"{total_members:,}"
            )

        with c2:

            st.metric(
                "💰 ጠቅላላ ቁጠባ",
                f"{total_savings:,.2f} ብር"
            )

        with c3:

            st.metric(
                "💵 ጠቅላላ ብድር",
                f"{total_loan:,.2f} ብር"
            )

        with c4:

            st.metric(
                "💳 የቀረ ዕዳ",
                f"{total_debt:,.2f} ብር"
            )

        st.divider()

        # ----------------------------------------------------
        # INCOME
        # ----------------------------------------------------

        st.subheader(
            "🏦 የማህበሩ ገቢ ማጠቃለያ"
        )

        i1, i2, i3, i4 = st.columns(4)

        with i1:

            st.metric(
                "🧾 10% የብድር ክፍያ",
                f"{total_loan_fee:,.2f} ብር"
            )

        with i2:

            st.metric(
                "📈 የተከፈለ ወለድ",
                f"{total_interest:,.2f} ብር"
            )

        with i3:

            st.metric(
                "📝 የመመዝገቢያ ክፍያ",
                f"{total_registration_fee:,.2f} ብር"
            )

        with i4:

            st.metric(
                "🏦 ጠቅላላ የማህበሩ ገቢ",
                f"{cooperative_income:,.2f} ብር"
            )

        st.info(
            "📌 የማህበሩ ገቢ = "
            "10% የብድር ክፍያ + "
            "የተከፈለ ወለድ + "
            "የመመዝገቢያ ክፍያ"
        )

        st.divider()

        # ----------------------------------------------------
        # REPORT TABLE
        # ----------------------------------------------------

        st.subheader(
            "📋 የአባላት ሙሉ ሪፖርት"
        )

        st.dataframe(
            report_df,
            use_container_width=True,
            hide_index=True
        )

        # ----------------------------------------------------
        # LOAN PAYMENT HISTORY
        # ----------------------------------------------------

        st.divider()

        st.subheader(
            "📜 የሁሉም አባላት የብድር ክፍያ ታሪክ"
        )

        if payment_history:

            all_history_df = pd.DataFrame(
                payment_history
            )

            st.dataframe(
                all_history_df,
                use_container_width=True,
                hide_index=True
            )

        else:

            st.info(
                "📌 እስካሁን "
                "የብድር ክፍያ ታሪክ የለም።"
            )

        # ----------------------------------------------------
        # DOWNLOAD EXCEL
        # ----------------------------------------------------

        if os.path.exists(DB_FILE):

            with open(
                DB_FILE,
                "rb"
            ) as file:

                st.download_button(
                    label="📥 የሙሉ ሲስተም Excel አውርድ",
                    data=file,
                    file_name="sacco_database.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    ),
                    use_container_width=True
                )

    else:

        st.info(
            "📌 እስካሁን የተመዘገበ "
            "መረጃ የለም።"
        )


# ============================================================
# 6. MEMBER EDIT / DELETE
# ============================================================

elif menu == "✏️ የአባላት መረጃ ማስተካከያ":

    st.header(
        "✏️ የአባላት መረጃ ማስተካከያ"
    )

    if not data:

        st.info(
            "ምንም አባል የለም።"
        )

    else:

        member_id = st.selectbox(
            "አባል ይምረጡ:",
            list(data.keys()),
            format_func=lambda x:
                f"{x} - {data[x]['የአባል ስም']}"
        )

        member = data[member_id]

        # ----------------------------------------------------
        # PHOTO
        # ----------------------------------------------------

        if member.get(
            "የአባል ፎቶ"
        ):

            try:

                photo_bytes = (
                    base64.b64decode(
                        member[
                            "የአባል ፎቶ"
                        ]
                    )
                )

                st.image(
                    photo_bytes,
                    width=180
                )

            except Exception:

                pass

        st.write(
            f"**ID:** {member_id}"
        )

        with st.form(
            f"edit_form_{member_id}"
        ):

            new_name = st.text_input(
                "የአባል ስም:",
                value=member.get(
                    "የአባል ስም",
                    ""
                )
            )

            new_national_id = st.text_input(
                "National ID (አማራጭ):",
                value=member.get(
                    "National ID",
                    ""
                )
            )

            new_photo = st.file_uploader(
                "አዲስ ፎቶ ካለ ይጫኑ:",
                type=[
                    "jpg",
                    "jpeg",
                    "png"
                ]
            )

            update_button = (
                st.form_submit_button(
                    "💾 መረጃውን አዘምን"
                )
            )

        if update_button:

            new_name = str(
                new_name
            ).strip()

            new_national_id = str(
                new_national_id or ""
            ).strip()

            if not new_name:

                st.error(
                    "❌ የአባል ስም "
                    "ባዶ ሊሆን አይችልም።"
                )

            elif national_id_exists(
                new_national_id,
                exclude_member_id=member_id
            ):

                st.error(
                    "❌ ይህ National ID "
                    "ቀድሞ ለሌላ አባል "
                    "ተመዝግቧል።"
                )

            else:

                member[
                    "የአባል ስም"
                ] = new_name

                member[
                    "National ID"
                ] = new_national_id

                if new_photo is not None:

                    photo_bytes = (
                        new_photo.read()
                    )

                    member[
                        "የአባል ፎቶ"
                    ] = (
                        base64.b64encode(
                            photo_bytes
                        )
                        .decode("utf-8")
                    )

                save_data_to_excel(
                    data,
                    payment_history
                )

                st.success(
                    "✅ የአባሉ መረጃ "
                    "በትክክል ተዘምኗል።"
                )

                st.rerun()

        # ----------------------------------------------------
        # DELETE MEMBER
        # ----------------------------------------------------

        st.divider()

        st.subheader(
            "🗑️ አባል ሰርዝ"
        )

        st.warning(
            "⚠️ አባልን መሰረዝ ይቻላል። "
            "ከተሰረዘ በኋላ በRegistration "
            "ገጽ እንደ አዲስ ማስገባት ይችላሉ።"
        )

        confirm_delete = st.checkbox(
            "ይህን አባል ለመሰረዝ አረጋግጣለሁ"
        )

        if st.button(
            "🗑️ አባሉን ሰርዝ",
            disabled=not confirm_delete
        ):

            del data[member_id]

            save_data_to_excel(
                data,
                payment_history
            )

            st.success(
                "✅ አባሉ ተሰርዟል።"
            )

            st.rerun()

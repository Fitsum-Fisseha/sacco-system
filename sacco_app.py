import streamlit as st
import pandas as pd
import os
import base64
from io import StringIO, BytesIO

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

# This version identifies the initial 136-member data.
# It prevents the 136 members from being loaded repeatedly.
SEED_VERSION = "fixed136-v1"

MONTHS = [
    "ግንቦት 2018",
    "ሰኔ 2018",
    "ሐምሌ 2018"
]


# ============================================================
# INITIAL 136 MEMBERS
# ============================================================
#
# Format:
# ID    Name    May    June    July
#
# Empty savings = 0
#
# IMPORTANT:
# ID 113 is ተናኜ አንገሎ
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
# CREATE FIXED MEMBER DATA
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

            "ብድር ሁኔታ": "የለም",
            "የተበደረው ጠቅላላ (ብር)": 0.0,
            "በእጅ የተሰጠ 90% (ብር)": 0.0,
            "የቀረው ዕዳ (ብር)": 0.0,

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

        "ብድር ሁኔታ": "የለም",
        "የተበደረው ጠቅላላ (ብር)": 0.0,
        "በእጅ የተሰጠ 90% (ብር)": 0.0,
        "የቀረው ዕዳ (ብር)": 0.0
    }

    for key, value in defaults.items():

        if key not in member or pd.isna(member[key]):
            member[key] = value

    # Text fields
    for key in [
        "የአባል ስም",
        "National ID",
        "የአባል ፎቶ",
        "ብድር ሁኔታ"
    ]:
        if pd.isna(member[key]):
            member[key] = ""
        else:
            member[key] = str(member[key]).strip()

    # Numeric fields
    numeric_fields = [
        "ግንቦት 2018",
        "ሰኔ 2018",
        "ሐምሌ 2018",
        "ጠቅላላ ቁጠባ (ብር)",
        "የተበደረው ጠቅላላ (ብር)",
        "በእጅ የተሰጠ 90% (ብር)",
        "የቀረው ዕዳ (ብር)"
    ]

    for key in numeric_fields:
        try:
            member[key] = float(member[key])
        except Exception:
            member[key] = 0.0

    return member


# ============================================================
# SAVE DATABASE
# ============================================================

def save_data_to_excel(db):

    if not db:
        return

    rows = []

    for member_id, member in db.items():

        row = dict(member)

        row["መታወቂያ ቁጥር (ID)"] = str(member_id)

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
        "ብድር ሁኔታ",
        "የተበደረው ጠቅላላ (ብር)",
        "በእጅ የተሰጠ 90% (ብር)",
        "የቀረው ዕዳ (ብር)",
        "_seed_version"
    ]

    existing_columns = [
        col for col in first_columns
        if col in df.columns
    ]

    other_columns = [
        col for col in df.columns
        if col not in existing_columns
    ]

    df = df[existing_columns + other_columns]

    df.to_excel(
        DB_FILE,
        index=False
    )


# ============================================================
# LOAD DATABASE
# ============================================================

def load_database():

    fixed_members = create_initial_members()

    # --------------------------------------------------------
    # If database doesn't exist, create the 136 members
    # --------------------------------------------------------

    if not os.path.exists(DB_FILE):

        save_data_to_excel(fixed_members)

        return fixed_members

    # --------------------------------------------------------
    # Read existing database
    # --------------------------------------------------------

    try:

        df = pd.read_excel(
            DB_FILE,
            dtype={
                "መታወቂያ ቁጥር (ID)": str
            }
        )

        if (
            df.empty
            or "መታወቂያ ቁጥር (ID)" not in df.columns
        ):

            save_data_to_excel(fixed_members)

            return fixed_members

        df["መታወቂያ ቁጥር (ID)"] = (
            df["መታወቂያ ቁጥር (ID)"]
            .astype(str)
            .str.strip()
            .str.replace(r"\.0$", "", regex=True)
        )

        old_data = (
            df
            .set_index("መታወቂያ ቁጥር (ID)")
            .to_dict(orient="index")
        )

    except Exception:

        save_data_to_excel(fixed_members)

        return fixed_members


    # --------------------------------------------------------
    # IMPORTANT:
    #
    # The first time this new version runs, the old member
    # records are replaced by the 136 supplied members.
    #
    # Future new members are NOT blocked.
    # --------------------------------------------------------

    already_seeded = False

    if old_data:

        versions = []

        for member in old_data.values():

            version = member.get("_seed_version", "")

            if not pd.isna(version):
                versions.append(str(version))

        if SEED_VERSION in versions:
            already_seeded = True


    # --------------------------------------------------------
    # FIRST RUN OF FIXED 136
    # --------------------------------------------------------

    if not already_seeded:

        # Completely replace old member records.
        #
        # No old member records are preserved.
        #
        # This creates exactly the supplied 136 members.

        new_data = fixed_members

        save_data_to_excel(new_data)

        return new_data


    # --------------------------------------------------------
    # AFTER INITIAL 136 MEMBERS ARE LOADED
    #
    # Keep all existing members, including future members.
    # --------------------------------------------------------

    data = {}

    for member_id, member in old_data.items():

        member_id = str(member_id).strip()

        member = normalize_member(member)

        data[member_id] = member


    return data


# ============================================================
# INITIALIZE DATABASE
# ============================================================

data = load_database()


# ============================================================
# NATIONAL ID DUPLICATE CHECK
# ============================================================

def national_id_exists(
    national_id,
    exclude_member_id=None
):

    national_id = str(
        national_id or ""
    ).strip()

    # National ID is optional.
    if not national_id:
        return False

    for member_id, member in data.items():

        if (
            exclude_member_id is not None
            and str(member_id) == str(exclude_member_id)
        ):
            continue

        existing_id = str(
            member.get("National ID", "") or ""
        ).strip()

        if (
            existing_id
            and existing_id == national_id
        ):
            return True

    return False


# ============================================================
# LOAN PAYMENT FUNCTION
# ============================================================

def apply_loan_payment(
    member,
    payment_amount
):

    payment_amount = float(payment_amount)

    remaining_debt = float(
        member.get(
            "የቀረው ዕዳ (ብር)",
            0
        )
    )

    original_loan = float(
        member.get(
            "የተበደረው ጠቅላላ (ብር)",
            0
        )
    )

    if remaining_debt <= 0:

        member["ብድር ሁኔታ"] = "የለም"
        member["የቀረው ዕዳ (ብር)"] = 0.0

        return payment_amount


    # 2% monthly interest based on original loan
    monthly_interest = original_loan * 0.02

    # Interest is paid first
    interest_paid = min(
        payment_amount,
        monthly_interest
    )

    remaining_payment = (
        payment_amount - interest_paid
    )

    # Remaining payment reduces principal
    principal_paid = min(
        remaining_payment,
        remaining_debt
    )

    new_debt = (
        remaining_debt - principal_paid
    )

    member["የቀረው ዕዳ (ብር)"] = max(
        new_debt,
        0
    )

    # If loan is completely paid
    if member["የቀረው ዕዳ (ብር)"] <= 0.01:

        member["የቀረው ዕዳ (ብር)"] = 0.0
        member["ብድር ሁኔታ"] = "የለም"

    else:

        member["ብድር ሁኔታ"] = "ያለበት"


    # Return only amount that was beyond the debt
    total_used = interest_paid + principal_paid

    excess = payment_amount - total_used

    return max(excess, 0)


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

    st.header("👤 አዲስ አባል መመዝገቢያ")

    st.info(
        "📌 ይህ ገጽ ለወደፊት አዲስ አባላትን "
        "ለመመዝገብ ይጠቀሙበታል።"
    )

    with st.form("member_registration_form"):

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
        m_name = str(m_name).strip()
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
                f"❌ ID {m_id} ቀድሞ ተመዝግቧል።"
            )

        elif (
            m_national_id
            and national_id_exists(m_national_id)
        ):

            st.error(
                "❌ ይህ National ID ቀድሞ "
                "ለሌላ አባል ተመዝግቧል።"
            )

        else:

            photo_data = ""

            if m_photo is not None:

                photo_bytes = m_photo.read()

                photo_data = base64.b64encode(
                    photo_bytes
                ).decode("utf-8")


            data[m_id] = {

                "የአባል ስም": m_name,

                "National ID": m_national_id,

                "የአባል ፎቶ": photo_data,

                "ግንቦት 2018": 0.0,
                "ሰኔ 2018": 0.0,
                "ሐምሌ 2018": 0.0,

                "ጠቅላላ ቁጠባ (ብር)": 0.0,

                "ብድር ሁኔታ": "የለም",

                "የተበደረው ጠቅላላ (ብር)": 0.0,

                "በእጅ የተሰጠ 90% (ብር)": 0.0,

                "የቀረው ዕዳ (ብር)": 0.0,

                "_seed_version": SEED_VERSION
            }

            save_data_to_excel(data)

            st.success(
                f"✅ {m_name} በተሳካ ሁኔታ ተመዝግቧል።"
            )

            st.rerun()


# ============================================================
# 2. MONTHLY SAVINGS
# ============================================================

elif menu == "💰 የወር ቁጠባ ማስገቢያ":

    st.header("💰 የወርሃዊ ቁጠባ መመዝገቢያ")

    member_ids = list(data.keys())

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


    amount = st.number_input(
        "የቁጠባ / ክፍያ መጠን (ብር):",
        min_value=0.0,
        step=100.0
    )


    if st.button("💾 ቁጠባ መዝግብ"):

        if amount <= 0:

            st.error(
                "❌ የገንዘብ መጠን ያስገቡ።"
            )

        elif not month:

            st.error(
                "❌ ወር ያስገቡ።"
            )

        else:

            # ------------------------------------------------
            # If member has outstanding loan,
            # payment goes to loan first.
            # ------------------------------------------------

            if (
                member["የቀረው ዕዳ (ብር)"]
                > 0
            ):

                old_debt = member[
                    "የቀረው ዕዳ (ብር)"
                ]

                excess = apply_loan_payment(
                    member,
                    amount
                )

                used_for_loan = (
                    amount - excess
                )

                # Only excess money becomes savings
                if excess > 0:

                    member[month] = (
                        float(
                            member.get(month, 0)
                            or 0
                        )
                        + excess
                    )

                    member[
                        "ጠቅላላ ቁጠባ (ብር)"
                    ] += excess

                save_data_to_excel(data)

                st.success(
                    f"✅ {used_for_loan:,.2f} ብር "
                    f"ለብድር ክፍያ ተጠቅሟል።"
                )

                if excess > 0:

                    st.info(
                        f"💰 {excess:,.2f} ብር "
                        f"እንደ ቁጠባ ተጨምሯል።"
                    )

                st.info(
                    f"💳 የቀረው ዕዳ: "
                    f"{member['የቀረው ዕዳ (ብር)']:,.2f} ብር"
                )

            else:

                member[month] = (
                    float(
                        member.get(month, 0)
                        or 0
                    )
                    + amount
                )

                member[
                    "ጠቅላላ ቁጠባ (ብር)"
                ] += amount

                save_data_to_excel(data)

                st.success(
                    f"✅ {amount:,.2f} ብር "
                    f"በ{month} ተመዝግቧል።"
                )

            st.rerun()


# ============================================================
# 3. LOAN SERVICE
# ============================================================

elif menu == "💵 የብድር አገልግሎት":

    st.header("💵 የብድር አገልግሎት")

    member_ids = list(data.keys())

    member_id = st.selectbox(
        "አባል ይምረጡ:",
        member_ids,
        format_func=lambda x:
            f"{x} - {data[x]['የአባል ስም']}"
    )

    member = data[member_id]

    savings = float(
        member[
            "ጠቅላላ ቁጠባ (ብር)"
        ]
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

    if member[
        "የቀረው ዕዳ (ብር)"
    ] > 0:

        st.warning(
            f"⚠️ ይህ አባል ቀድሞ "
            f"{member['የቀረው ዕዳ (ብር)']:,.2f} "
            f"ብር ዕዳ አለበት። "
            f"አዲስ ብድር መውሰድ አይችልም።"
        )

    else:

        loan_amount = st.number_input(
            "የሚፈልጉት የብድር መጠን (ብር):",
            min_value=0.0,
            max_value=max_loan if max_loan > 0 else 0.0,
            step=100.0
        )

        if loan_amount > 0:

            upfront_fee = loan_amount * 0.10

            net_payout = (
                loan_amount - upfront_fee
            )

            monthly_principal = (
                loan_amount / 36
            )

            monthly_interest = (
                loan_amount * 0.02
            )

            monthly_payment = (
                monthly_principal
                + monthly_interest
            )

            st.divider()

            col1, col2, col3 = st.columns(3)

            with col1:

                st.metric(
                    "የብድር መጠን",
                    f"{loan_amount:,.2f} ብር"
                )

            with col2:

                st.metric(
                    "10% ክፍያ",
                    f"{upfront_fee:,.2f} ብር"
                )

            with col3:

                st.metric(
                    "በእጅ የሚሰጥ 90%",
                    f"{net_payout:,.2f} ብር"
                )

            st.write(
                f"📅 ወርሃዊ ዋና ብድር: "
                f"{monthly_principal:,.2f} ብር"
            )

            st.write(
                f"📈 ወርሃዊ ወለድ (2%): "
                f"{monthly_interest:,.2f} ብር"
            )

            st.write(
                f"💳 ወርሃዊ ክፍያ: "
                f"**{monthly_payment:,.2f} ብር**"
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
                        "በእጅ የተሰጠ 90% (ብር)"
                    ] = net_payout

                    member[
                        "የቀረው ዕዳ (ብር)"
                    ] = loan_amount

                    save_data_to_excel(data)

                    st.success(
                        f"🎉 ለ{member['የአባል ስም']} "
                        f"{loan_amount:,.2f} ብር "
                        f"ብድር ተፈቅዷል!"
                    )

                    st.info(
                        f"💵 በእጅ የሚሰጠው 90%: "
                        f"{net_payout:,.2f} ብር\n\n"
                        f"📅 የወርሃዊ ክፍያ "
                        f"(36 ወራት): "
                        f"{monthly_payment:,.2f} ብር"
                    )

                    st.rerun()


# ============================================================
# 4. LOAN REPAYMENT
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
            "📌 በአሁኑ ጊዜ ያለ ብድር አባል የለም።"
        )

    else:

        member_id = st.selectbox(
            "የሚከፍለውን አባል ይምረጡ:",
            member_ids_with_loan,
            format_func=lambda x:
                f"{x} - {data[x]['የአባል ስም']}"
        )

        member = data[member_id]

        st.warning(
            f"👤 {member['የአባል ስም']}\n\n"
            f"💳 የቀረ ዕዳ: "
            f"{member['የቀረው ዕዳ (ብር)']:,.2f} ብር"
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

                old_debt = member[
                    "የቀረው ዕዳ (ብር)"
                ]

                excess = apply_loan_payment(
                    member,
                    payment
                )

                save_data_to_excel(data)

                used_for_loan = (
                    payment - excess
                )

                st.success(
                    f"✅ {used_for_loan:,.2f} ብር "
                    f"ከዕዳው ላይ ተከፍሏል።"
                )

                if excess > 0:

                    st.info(
                        f"💰 {excess:,.2f} ብር "
                        f"ከዕዳው በላይ ስለሆነ "
                        f"በቁጠባ ላይ ሊጨመር ይችላል።"
                    )

                    member[
                        "ጠቅላላ ቁጠባ (ብር)"
                    ] += excess

                    save_data_to_excel(data)

                st.info(
                    f"💳 አሁን የቀረው ዕዳ: "
                    f"{member['የቀረው ዕዳ (ብር)']:,.2f} ብር"
                )

                st.rerun()


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
                "በእጅ የተሰጠ 90% (ብር)": float(
                    member.get(
                        "በእጅ የተሰጠ 90% (ብር)",
                        0
                    )
                ),
                "የቀረው ዕዳ (ብር)": float(
                    member.get(
                        "የቀረው ዕዳ (ብር)",
                        0
                    )
                )
            }

            rows.append(row)


        report_df = pd.DataFrame(rows)

        # Sort by ID numerically
        report_df["_sort_id"] = pd.to_numeric(
            report_df["ID"],
            errors="coerce"
        )

        report_df = (
            report_df
            .sort_values("_sort_id")
            .drop(columns=["_sort_id"])
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
        # REPORT TABLE
        # ----------------------------------------------------

        st.dataframe(
            report_df,
            use_container_width=True,
            hide_index=True
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
                    label="📥 የአባላት መረጃ Excel አውርድ",
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

        if member.get("የአባል ፎቶ"):

            try:

                photo_bytes = base64.b64decode(
                    member["የአባል ፎቶ"]
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

            update_button = st.form_submit_button(
                "💾 መረጃውን አዘምን"
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
                    "❌ የአባል ስም ባዶ ሊሆን አይችልም።"
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

                    photo_bytes = new_photo.read()

                    member[
                        "የአባል ፎቶ"
                    ] = base64.b64encode(
                        photo_bytes
                    ).decode("utf-8")


                save_data_to_excel(data)

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

            save_data_to_excel(data)

            st.success(
                "✅ አባሉ ተሰርዟል።"
            )

            st.rerun()

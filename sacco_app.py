import streamlit as st
import pandas as pd
import os
import base64
import math
from io import BytesIO

# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="የቁጠባ እና ብድር ሲስተም",
    layout="wide"
)

st.title("🏦 ተስፋ የገንዘብ ቁጠባና ብድር ማህበር")

DB_FILE = "sacco_database.xlsx"

# ============================================================
# CONSTANTS
# ============================================================

REGISTRATION_FEE = 500.0
LOAN_FEE_RATE = 0.10
MONTHLY_INTEREST_RATE = 0.02

LOAN_TERMS = [3, 6, 12, 24, 36]

MONTHS = [
    "ግንቦት 2018",
    "ሰኔ 2018",
    "ሐምሌ 2018",
    "ነሐሴ 2018",
    "መስከረም 2019",
    "ጥቅምት 2019",
    "ሌላ ወር"
]

# ============================================================
# MEMBER DEFAULT DATA
# ============================================================

def empty_member():
    return {
        "የአባል ስም": "",
        "National ID": "",
        "የአባል ፎቶ": "",

        "ጠቅላላ ቁጠባ (ብር)": 0.0,

        "ብድር ሁኔታ": "የለም",

        "የተበደረው ጠቅላላ (ብር)": 0.0,

        # 10% loan income
        "10% የብድር ክፍያ (ብር)": 0.0,

        # Amount actually given to member
        "በእጅ የተሰጠ 90% (ብር)": 0.0,

        # Loan principal
        "የቀረው ዋና ብድር (ብር)": 0.0,

        # Total debt including current calculated interest
        "የቀረው ዕዳ (ብር)": 0.0,

        # Loan term
        "የብድር ጊዜ (ወር)": 36,

        # Cumulative interest
        "የተከፈለ ወለድ (ብር)": 0.0,

        # Cumulative principal
        "የተከፈለ ዋና ብድር (ብር)": 0.0,

        # Registration fee is separate from loan income
        "የመመዝገቢያ ክፍያ (ብር)": REGISTRATION_FEE
    }


# ============================================================
# NORMALIZE MEMBER
# ============================================================

def normalize_member(member):

    defaults = empty_member()

    for key, value in defaults.items():
        if key not in member:
            member[key] = value

    # -----------------------------------------
    # Text fields
    # -----------------------------------------

    for field in [
        "የአባል ስም",
        "National ID"
    ]:
        if pd.isna(member.get(field)):
            member[field] = ""
        else:
            member[field] = str(member.get(field, "")).strip()

    if pd.isna(member.get("የአባል ፎቶ")):
        member["የአባል ፎቶ"] = ""

    # -----------------------------------------
    # Numeric fields
    # -----------------------------------------

    numeric_fields = [
        "ጠቅላላ ቁጠባ (ብር)",
        "የተበደረው ጠቅላላ (ብር)",
        "10% የብድር ክፍያ (ብር)",
        "በእጅ የተሰጠ 90% (ብር)",
        "የቀረው ዋና ብድር (ብር)",
        "የቀረው ዕዳ (ብር)",
        "የተከፈለ ወለድ (ብር)",
        "የተከፈለ ዋና ብድር (ብር)",
        "የመመዝገቢያ ክፍያ (ብር)"
    ]

    for field in numeric_fields:
        try:
            value = member.get(field, 0)

            if pd.isna(value):
                member[field] = 0.0
            else:
                member[field] = float(value)

        except Exception:
            member[field] = 0.0

    # -----------------------------------------
    # Loan term
    # -----------------------------------------

    try:
        member["የብድር ጊዜ (ወር)"] = int(
            float(member.get("የብድር ጊዜ (ወር)", 36))
        )
    except Exception:
        member["የብድር ጊዜ (ወር)"] = 36

    # -----------------------------------------
    # BACKWARD COMPATIBILITY
    #
    # Old database had only "remaining debt".
    # Use it as remaining principal when necessary.
    # -----------------------------------------

    remaining_principal = float(
        member.get("የቀረው ዋና ብድር (ብር)", 0)
    )

    remaining_debt = float(
        member.get("የቀረው ዕዳ (ብር)", 0)
    )

    original_loan = float(
        member.get("የተበደረው ጠቅላላ (ብር)", 0)
    )

    if remaining_principal <= 0 and remaining_debt > 0:
        member["የቀረው ዋና ብድር (ብር)"] = remaining_debt

    # Old records may not have 10% fee
    if (
        member.get("10% የብድር ክፍያ (ብር)", 0) == 0
        and original_loan > 0
    ):
        member["10% የብድር ክፍያ (ብር)"] = (
            original_loan * LOAN_FEE_RATE
        )

    # Old records may not have registration fee
    if (
        member.get("የመመዝገቢያ ክፍያ (ብር)", 0) == 0
    ):
        member["የመመዝገቢያ ክፍያ (ብር)"] = REGISTRATION_FEE

    return member


# ============================================================
# LOAD DATABASE
# ============================================================

def load_data():

    if not os.path.exists(DB_FILE):
        return {}

    try:

        df = pd.read_excel(
            DB_FILE,
            dtype={"መታወቂያ ቁጥር (ID)": str}
        )

        if df.empty:
            return {}

        if "መታወቂያ ቁጥር (ID)" not in df.columns:
            return {}

        df["መታወቂያ ቁጥር (ID)"] = (
            df["መታወቂያ ቁጥር (ID)"]
            .astype(str)
            .str.strip()
        )

        df = df.set_index(
            "መታወቂያ ቁጥር (ID)"
        )

        data = df.to_dict(
            orient="index"
        )

        for member_id in data:
            data[member_id] = normalize_member(
                data[member_id]
            )

        return data

    except Exception as e:

        st.error(
            f"Excel database ማንበብ አልተቻለም፦ {e}"
        )

        return {}


# ============================================================
# SAVE DATABASE
# ============================================================

def save_data_to_excel(data):

    rows = []

    for member_id, member in data.items():

        row = normalize_member(member.copy())

        row["መታወቂያ ቁጥር (ID)"] = str(
            member_id
        )

        rows.append(row)

    if rows:

        df = pd.DataFrame(rows)

        first_column = (
            "መታወቂያ ቁጥር (ID)"
        )

        other_columns = [
            col
            for col in df.columns
            if col != first_column
        ]

        df = df[
            [first_column] + other_columns
        ]

    else:

        df = pd.DataFrame(
            columns=[
                "መታወቂያ ቁጥር (ID)",
                "የአባል ስም",
                "National ID",
                "የአባል ፎቶ",
                "ጠቅላላ ቁጠባ (ብር)",
                "ብድር ሁኔታ",
                "የተበደረው ጠቅላላ (ብር)",
                "10% የብድር ክፍያ (ብር)",
                "በእጅ የተሰጠ 90% (ብር)",
                "የቀረው ዋና ብድር (ብር)",
                "የቀረው ዕዳ (ብር)",
                "የብድር ጊዜ (ወር)",
                "የተከፈለ ወለድ (ብር)",
                "የተከፈለ ዋና ብድር (ብር)",
                "የመመዝገቢያ ክፍያ (ብር)"
            ]
        )

    df.to_excel(
        DB_FILE,
        index=False
    )


# ============================================================
# INITIALIZE
# ============================================================

data = load_data()


# ============================================================
# NATIONAL ID DUPLICATE CHECK
# ============================================================

def national_id_exists(
    national_id,
    exclude_member_id=None
):

    national_id = str(
        national_id
    ).strip()

    if not national_id:
        return False

    for member_id, member in data.items():

        if exclude_member_id is not None:

            if str(member_id) == str(
                exclude_member_id
            ):
                continue

        existing_id = str(
            member.get(
                "National ID",
                ""
            )
        ).strip()

        if (
            existing_id
            and existing_id == national_id
        ):
            return True

    return False


# ============================================================
# LOAN CALCULATION
# ============================================================

def calculate_monthly_payment(
    principal,
    annual_or_monthly_rate,
    months
):

    principal = float(principal)
    rate = float(annual_or_monthly_rate)
    months = int(months)

    if principal <= 0 or months <= 0:
        return 0.0

    # Monthly reducing-balance amortization
    if rate == 0:
        return principal / months

    payment = (
        principal
        * rate
        * ((1 + rate) ** months)
        / (
            ((1 + rate) ** months) - 1
        )
    )

    return payment


# ============================================================
# LOAN PAYMENT FUNCTION
# ============================================================

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
            0
        )
    )

    if remaining_principal <= 0:

        member[
            "የቀረው ዋና ብድር (ብር)"
        ] = 0.0

        member[
            "የቀረው ዕዳ (ብር)"
        ] = 0.0

        member[
            "ብድር ሁኔታ"
        ] = "የለም"

        return {
            "interest_paid": 0.0,
            "principal_paid": 0.0,
            "excess": payment_amount
        }

    # --------------------------------------------------------
    # REDUCING BALANCE INTEREST
    # --------------------------------------------------------

    monthly_interest = (
        remaining_principal
        * MONTHLY_INTEREST_RATE
    )

    # Interest first
    interest_paid = min(
        payment_amount,
        monthly_interest
    )

    remaining_payment = (
        payment_amount
        - interest_paid
    )

    # Then principal
    principal_paid = min(
        remaining_payment,
        remaining_principal
    )

    new_principal = (
        remaining_principal
        - principal_paid
    )

    excess = (
        payment_amount
        - interest_paid
        - principal_paid
    )

    # --------------------------------------------------------
    # SAVE PAYMENT HISTORY
    # --------------------------------------------------------

    member[
        "የተከፈለ ወለድ (ብር)"
    ] = float(
        member.get(
            "የተከፈለ ወለድ (ብር)",
            0
        )
    ) + interest_paid

    member[
        "የተከፈለ ዋና ብድር (ብር)"
    ] = float(
        member.get(
            "የተከፈለ ዋና ብድር (ብር)",
            0
        )
    ) + principal_paid

    member[
        "የቀረው ዋና ብድር (ብር)"
    ] = max(
        new_principal,
        0
    )

    # --------------------------------------------------------
    # CURRENT DEBT
    # --------------------------------------------------------

    if member[
        "የቀረው ዋና ብድር (ብር)"
    ] > 0:

        member[
            "የቀረው ዕዳ (ብር)"
        ] = (
            member[
                "የቀረው ዋና ብድር (ብር)"
            ]
            + (
                member[
                    "የቀረው ዋና ብድር (ብር)"
            ]
                * MONTHLY_INTEREST_RATE
            )
        )

        member[
            "ብድር ሁኔታ"
        ] = "ያለበት"

    else:

        member[
            "የቀረው ዕዳ (ብር)"
        ] = 0.0

        member[
            "ብድር ሁኔታ"
        ] = "የለም"

    return {
        "interest_paid": interest_paid,
        "principal_paid": principal_paid,
        "excess": max(excess, 0)
    }


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.header("📋 ምናሌ")

menu = st.sidebar.radio(
    "ገጽ ይምረጡ",
    [
        "👤 አባል መመዝገቢያ",
        "💰 የወር ቁጠባ ማስገቢያ",
        "💵 የብድር አገልግሎት",
        "📅 የብድር ክፍያ መመዝገቢያ",
        "📊 ጠቅላላ ሪፖርት",
        "✏️ የአባላት መረጃ ማስተካከያ",
        "📤 Excel ፋይል አስገባ"
    ]
)


# ============================================================
# 1. MEMBER REGISTRATION
# ============================================================

if menu == "👤 አባል መመዝገቢያ":

    st.header("👤 አዲስ አባል መመዝገቢያ")

    with st.form(
        "member_registration_form"
    ):

        col1, col2 = st.columns(2)

        with col1:

            member_id = st.text_input(
                "ID *",
                placeholder="ለምሳሌ 137"
            ).strip()

            member_name = st.text_input(
                "የአባል ስም *"
            ).strip()

            national_id = st.text_input(
                "National ID (አማራጭ)"
            ).strip()

        with col2:

            st.info(
                f"የመመዝገቢያ ክፍያ፦ "
                f"{REGISTRATION_FEE:,.2f} ብር"
            )

            photo = st.file_uploader(
                "የአባል ፎቶ (አማራጭ)",
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

        if not member_id:

            st.error(
                "ID ያስገቡ።"
            )

        elif not member_name:

            st.error(
                "የአባል ስም ያስገቡ።"
            )

        elif member_id in data:

            st.error(
                "ይህ ID ቀድሞ ተመዝግቧል።"
            )

        elif (
            national_id
            and national_id_exists(
                national_id
            )
        ):

            st.error(
                "ይህ National ID "
                "ቀድሞ ተመዝግቧል።"
            )

        else:

            member = empty_member()

            member[
                "የአባል ስም"
            ] = member_name

            member[
                "National ID"
            ] = national_id

            if photo is not None:

                photo_bytes = photo.read()

                member[
                    "የአባል ፎቶ"
                ] = base64.b64encode(
                    photo_bytes
                ).decode("utf-8")

            data[member_id] = member

            save_data_to_excel(data)

            st.success(
                f"✅ {member_name} "
                "በትክክል ተመዝግቧል።"
            )

            st.rerun()


# ============================================================
# 2. MONTHLY SAVINGS
# ============================================================

elif menu == "💰 የወር ቁጠባ ማስገቢያ":

    st.header(
        "💰 የወር ቁጠባ ማስገቢያ"
    )

    if not data:

        st.warning(
            "ምንም አባል አልተመዘገበም።"
        )

    else:

        member_id = st.selectbox(
            "አባል ይምረጡ",
            list(data.keys()),
            format_func=lambda x:
                f"{x} - "
                f"{data[x]['የአባል ስም']}"
        )

        member = data[member_id]

        st.info(
            f"👤 አባል፦ "
            f"{member['የአባል ስም']}"
        )

        st.write(
            f"💰 ጠቅላላ ቁጠባ፦ "
            f"**{member['ጠቅላላ ቁጠባ (ብር)']:,.2f} ብር**"
        )

        month = st.selectbox(
            "ወር ይምረጡ",
            MONTHS
        )

        current_month_amount = float(
            member.get(
                month,
                0
            ) or 0
        )

        # ----------------------------------------------------
        # EXISTING MONTH WARNING
        # ----------------------------------------------------

        if current_month_amount > 0:

            st.warning(
                f"⚠️ {month} ላይ "
                f"{current_month_amount:,.2f} ብር "
                "ቀድሞ ተመዝግቧል።\n\n"
                "በዚህ ወር ሌላ ብር ከመጨመር "
                "ይልቅ ካለው መጠን ላይ "
                "✏️ Edit ማድረግ ይችላሉ።"
            )

            edit_month = st.checkbox(
                f"✏️ {month} የቁጠባ መጠን አስተካክል"
            )

            if edit_month:

                new_amount = st.number_input(
                    f"አዲሱ የ{month} ቁጠባ",
                    min_value=0.0,
                    value=current_month_amount,
                    step=100.0
                )

                if st.button(
                    "💾 የወሩን መጠን አዘምን"
                ):

                    difference = (
                        new_amount
                        - current_month_amount
                    )

                    member[
                        "ጠቅላላ ቁጠባ (ብር)"
                    ] += difference

                    member[month] = new_amount

                    save_data_to_excel(data)

                    st.success(
                        f"✅ {month} "
                        f"ከ{current_month_amount:,.2f} "
                        f"ወደ {new_amount:,.2f} "
                        "ብር ተስተካክሏል።"
                    )

                    st.rerun()

        else:

            amount = st.number_input(
                "የቁጠባ መጠን (ብር)",
                min_value=0.0,
                step=100.0
            )

            if st.button(
                "💾 ቁጠባ መዝግብ"
            ):

                if amount <= 0:

                    st.error(
                        "የገንዘብ መጠን ያስገቡ።"
                    )

                else:

                    # ----------------------------------------
                    # NORMAL SAVINGS
                    # ----------------------------------------

                    member[
                        "ጠቅላላ ቁጠባ (ብር)"
                    ] += amount

                    member[month] = amount

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

    st.header(
        "💵 የብድር ማመልከቻ"
    )

    if not data:

        st.warning(
            "ምንም አባል አልተመዘገበም።"
        )

    else:

        member_id = st.selectbox(
            "አባል ይምረጡ",
            list(data.keys()),
            format_func=lambda x:
                f"{x} - "
                f"{data[x]['የአባል ስም']}"
        )

        member = data[member_id]

        savings = float(
            member.get(
                "ጠቅላላ ቁጠባ (ብር)",
                0
            )
        )

        max_loan = savings * 4

        current_debt = float(
            member.get(
                "የቀረው ዕዳ (ብር)",
                0
            )
        )

        st.write(
            f"👤 አባል፦ "
            f"**{member['የአባል ስም']}**"
        )

        col1, col2 = st.columns(2)

        with col1:

            st.metric(
                "💰 ጠቅላላ ቁጠባ",
                f"{savings:,.2f} ብር"
            )

        with col2:

            st.metric(
                "💵 ከፍተኛ ብድር",
                f"{max_loan:,.2f} ብር"
            )

        if current_debt > 0:

            st.error(
                f"❌ ይህ አባል "
                f"{current_debt:,.2f} ብር "
                "የቀረ ዕዳ አለበት። "
                "አዲስ ብድር መውሰድ አይችልም።"
            )

        elif savings <= 0:

            st.error(
                "ቁጠባ ሳይኖር ብድር መውሰድ አይቻልም።"
            )

        else:

            loan_amount = st.number_input(
                "የሚፈለገው ብድር (ብር)",
                min_value=0.0,
                max_value=float(max_loan),
                step=100.0
            )

            loan_term = st.selectbox(
                "የብድር ጊዜ",
                LOAN_TERMS,
                format_func=lambda x:
                    f"{x} ወር"
            )

            if loan_amount > 0:

                # --------------------------------------------
                # 10% LOAN FEE
                # --------------------------------------------

                loan_fee = (
                    loan_amount
                    * LOAN_FEE_RATE
                )

                net_payment = (
                    loan_amount
                    - loan_fee
                )

                # --------------------------------------------
                # REDUCING BALANCE PAYMENT
                # --------------------------------------------

                monthly_payment = (
                    calculate_monthly_payment(
                        loan_amount,
                        MONTHLY_INTEREST_RATE,
                        loan_term
                    )
                )

                first_month_interest = (
                    loan_amount
                    * MONTHLY_INTEREST_RATE
                )

                first_month_principal = (
                    monthly_payment
                    - first_month_interest
                )

                total_expected_payment = (
                    monthly_payment
                    * loan_term
                )

                total_expected_interest = (
                    total_expected_payment
                    - loan_amount
                )

                st.subheader(
                    "📋 የብድር ማጠቃለያ"
                )

                col1, col2, col3 = st.columns(3)

                with col1:

                    st.metric(
                        "የተበደረው 100%",
                        f"{loan_amount:,.2f} ብር"
                    )

                with col2:

                    st.metric(
                        "10% የማህበሩ ገቢ",
                        f"{loan_fee:,.2f} ብር"
                    )

                with col3:

                    st.metric(
                        "በእጅ የሚሰጥ 90%",
                        f"{net_payment:,.2f} ብር"
                    )

                st.divider()

                col1, col2, col3 = st.columns(3)

                with col1:

                    st.metric(
                        f"ወርሃዊ ክፍያ ({loan_term} ወር)",
                        f"{monthly_payment:,.2f} ብር"
                    )

                with col2:

                    st.metric(
                        "የመጀመሪያ ወር ወለድ",
                        f"{first_month_interest:,.2f} ብር"
                    )

                with col3:

                    st.metric(
                        "የመጀመሪያ ወር ዋና",
                        f"{first_month_principal:,.2f} ብር"
                    )

                st.info(
                    f"📈 የወለድ መጠን፦ "
                    f"{MONTHLY_INTEREST_RATE * 100:.0f}% "
                    "በወር — በቀሪ ዋና ብድር ላይ\n\n"
                    f"💰 ለ{loan_term} ወራት "
                    f"ጠቅላላ የሚከፈለው፦ "
                    f"{total_expected_payment:,.2f} ብር\n\n"
                    f"📊 የሚገመተው ጠቅላላ ወለድ፦ "
                    f"{total_expected_interest:,.2f} ብር\n\n"
                    f"🏦 የማህበሩ ገቢ፦ "
                    f"10% ({loan_fee:,.2f}) + "
                    f"ወለድ ({total_expected_interest:,.2f})"
                )

                if st.button(
                    "💵 ብድር ስጥ"
                ):

                    if loan_amount > max_loan:

                        st.error(
                            "ከተፈቀደው የብድር መጠን በላይ ነው።"
                        )

                    else:

                        # ------------------------------------
                        # RECORD LOAN
                        # ------------------------------------

                        member[
                            "የተበደረው ጠቅላላ (ብር)"
                        ] = loan_amount

                        member[
                            "10% የብድር ክፍያ (ብር)"
                        ] = loan_fee

                        member[
                            "በእጅ የተሰጠ 90% (ብር)"
                        ] = net_payment

                        member[
                            "የቀረው ዋና ብድር (ብር)"
                        ] = loan_amount

                        member[
                            "የቀረው ዕዳ (ብር)"
                        ] = loan_amount

                        member[
                            "የብድር ጊዜ (ወር)"
                        ] = loan_term

                        member[
                            "የተከፈለ ወለድ (ብር)"
                        ] = 0.0

                        member[
                            "የተከፈለ ዋና ብድር (ብር)"
                        ] = 0.0

                        member[
                            "ብድር ሁኔታ"
                        ] = "ያለበት"

                        save_data_to_excel(
                            data
                        )

                        st.success(
                            f"🎉 {member['የአባል ስም']} "
                            "ብድሩ ተመዝግቧል።"
                        )

                        st.info(
                            f"💵 የብድሩ መጠን፦ "
                            f"{loan_amount:,.2f} ብር\n\n"
                            f"🏦 10% የማህበሩ ገቢ፦ "
                            f"{loan_fee:,.2f} ብር\n\n"
                            f"👤 ለአባሉ በእጅ የተሰጠ፦ "
                            f"{net_payment:,.2f} ብር\n\n"
                            f"🔴 የአባሉ ዋና ዕዳ፦ "
                            f"{loan_amount:,.2f} ብር"
                        )

                        st.rerun()


# ============================================================
# 4. LOAN REPAYMENT
# ============================================================

elif menu == "📅 የብድር ክፍያ መመዝገቢያ":

    st.header(
        "📅 የብድር ክፍያ መመዝገቢያ"
    )

    members_with_debt = [
        member_id
        for member_id, member in data.items()
        if float(
            member.get(
                "የቀረው ዋና ብድር (ብር)",
                0
            )
        ) > 0
    ]

    if not members_with_debt:

        st.info(
            "አሁን የቀረ ብድር ያለበት "
            "አባል የለም።"
        )

    else:

        member_id = st.selectbox(
            "አባል ይምረጡ",
            members_with_debt,
            format_func=lambda x:
                f"{x} - "
                f"{data[x]['የአባል ስም']}"
        )

        member = data[member_id]

        original_loan = float(
            member.get(
                "የተበደረው ጠቅላላ (ብር)",
                0
            )
        )

        remaining_principal = float(
            member.get(
                "የቀረው ዋና ብድር (ብር)",
                0
            )
        )

        loan_term = int(
            member.get(
                "የብድር ጊዜ (ወር)",
                36
            )
        )

        current_interest = (
            remaining_principal
            * MONTHLY_INTEREST_RATE
        )

        normal_monthly_payment = (
            calculate_monthly_payment(
                original_loan,
                MONTHLY_INTEREST_RATE,
                loan_term
            )
        )

        st.write(
            f"👤 አባል፦ "
            f"**{member['የአባል ስም']}**"
        )

        col1, col2, col3 = st.columns(3)

        with col1:

            st.metric(
                "የመጀመሪያ ብድር",
                f"{original_loan:,.2f} ብር"
            )

        with col2:

            st.metric(
                "የቀረ ዋና",
                f"{remaining_principal:,.2f} ብር"
            )

        with col3:

            st.metric(
                "የዚህ ወር 2% ወለድ",
                f"{current_interest:,.2f} ብር"
            )

        st.info(
            f"📅 የብድር ጊዜ፦ "
            f"{loan_term} ወር\n\n"
            f"💳 መደበኛ ወርሃዊ ክፍያ፦ "
            f"{normal_monthly_payment:,.2f} ብር\n\n"
            f"📈 ይህ ወር ወለድ፦ "
            f"{current_interest:,.2f} ብር\n\n"
            f"📉 ከክፍያው ቀሪው ወደ ዋና ብድር ይሄዳል።"
        )

        payment = st.number_input(
            "የክፍያ መጠን (ብር)",
            min_value=0.0,
            step=100.0
        )

        if st.button(
            "💾 ክፍያ መዝግብ"
        ):

            if payment <= 0:

                st.error(
                    "የክፍያ መጠን ያስገቡ።"
                )

            else:

                result = apply_loan_payment(
                    member,
                    payment
                )

                interest_paid = result[
                    "interest_paid"
                ]

                principal_paid = result[
                    "principal_paid"
                ]

                excess = result[
                    "excess"
                ]

                save_data_to_excel(
                    data
                )

                st.success(
                    f"✅ {payment:,.2f} ብር "
                    "ክፍያ ተመዝግቧል።"
                )

                col1, col2, col3 = st.columns(3)

                with col1:

                    st.metric(
                        "የተከፈለ ወለድ",
                        f"{interest_paid:,.2f} ብር"
                    )

                with col2:

                    st.metric(
                        "የተከፈለ ዋና",
                        f"{principal_paid:,.2f} ብር"
                    )

                with col3:

                    st.metric(
                        "የቀረ ዋና",
                        f"{member['የቀረው ዋና ብድር (ብር)']:,.2f} ብር"
                    )

                if excess > 0:

                    member[
                        "ጠቅላላ ቁጠባ (ብር)"
                    ] += excess

                    save_data_to_excel(
                        data
                    )

                    st.info(
                        f"ℹ️ {excess:,.2f} ብር "
                        "ተረፍቶ ወደ ቁጠባ ተጨምሯል።"
                    )

                if (
                    member[
                        "የቀረው ዋና ብድር (ብር)"
                    ] <= 0
                ):

                    st.success(
                        "🎉 ይህ አባል ብድሩን "
                        "ሙሉ በሙሉ ከፍሎ ጨርሷል!"
                    )

                st.rerun()


# ============================================================
# 5. GENERAL REPORT
# ============================================================

elif menu == "📊 ጠቅላላ ሪፖርት":

    st.header(
        "📊 ጠቅላላ የአባላት፣ "
        "የቁጠባ እና የብድር ሪፖርት"
    )

    if not data:

        st.warning(
            "ምንም የአባላት መረጃ የለም።"
        )

    else:

        report_rows = []

        for member_id, member in data.items():

            report_rows.append({

                "ID": member_id,

                "የአባል ስም":
                    member.get(
                        "የአባል ስም",
                        ""
                    ),

                "National ID":
                    member.get(
                        "National ID",
                        ""
                    ),

                "ጠቅላላ ቁጠባ (ብር)":
                    member.get(
                        "ጠቅላላ ቁጠባ (ብር)",
                        0
                    ),

                "ብድር ሁኔታ":
                    member.get(
                        "ብድር ሁኔታ",
                        "የለም"
                    ),

                "የተበደረው ጠቅላላ (ብር)":
                    member.get(
                        "የተበደረው ጠቅላላ (ብር)",
                        0
                    ),

                "10% የብድር ገቢ (ብር)":
                    member.get(
                        "10% የብድር ክፍያ (ብር)",
                        0
                    ),

                "በእጅ የተሰጠ 90% (ብር)":
                    member.get(
                        "በእጅ የተሰጠ 90% (ብር)",
                        0
                    ),

                "የብድር ጊዜ (ወር)":
                    member.get(
                        "የብድር ጊዜ (ወር)",
                        36
                    ),

                "የቀረው ዋና ብድር (ብር)":
                    member.get(
                        "የቀረው ዋና ብድር (ብር)",
                        0
                    ),

                "የቀረው ዕዳ (ብር)":
                    member.get(
                        "የቀረው ዕዳ (ብር)",
                        0
                    ),

                "የተከፈለ ወለድ (ብር)":
                    member.get(
                        "የተከፈለ ወለድ (ብር)",
                        0
                    ),

                "የተከፈለ ዋና ብድር (ብር)":
                    member.get(
                        "የተከፈለ ዋና ብድር (ብር)",
                        0
                    ),

                "የመመዝገቢያ ክፍያ (ብር)":
                    member.get(
                        "የመመዝገቢያ ክፍያ (ብር)",
                        REGISTRATION_FEE
                    )
            })

        report_df = pd.DataFrame(
            report_rows
        )

        # ----------------------------------------------------
        # SEARCH
        # ----------------------------------------------------

        search = st.text_input(
            "🔎 አባል ስም / ID / National ID ፈልግ"
        ).strip()

        if search:

            mask = (
                report_df["ID"]
                .astype(str)
                .str.contains(
                    search,
                    case=False,
                    na=False
                )
                |
                report_df["የአባል ስም"]
                .astype(str)
                .str.contains(
                    search,
                    case=False,
                    na=False
                )
                |
                report_df["National ID"]
                .astype(str)
                .str.contains(
                    search,
                    case=False,
                    na=False
                )
            )

            report_df = report_df[
                mask
            ]

        # ----------------------------------------------------
        # TOTALS
        # ----------------------------------------------------

        total_members = len(
            report_df
        )

        total_savings = report_df[
            "ጠቅላላ ቁጠባ (ብር)"
        ].sum()

        total_loans = report_df[
            "የተበደረው ጠቅላላ (ብር)"
        ].sum()

        total_loan_fees = report_df[
            "10% የብድር ገቢ (ብር)"
        ].sum()

        total_interest = report_df[
            "የተከፈለ ወለድ (ብር)"
        ].sum()

        total_remaining = report_df[
            "የቀረው ዋና ብድር (ብር)"
        ].sum()

        # ----------------------------------------------------
        # METRICS
        # ----------------------------------------------------

        col1, col2, col3, col4 = st.columns(4)

        with col1:

            st.metric(
                "👥 አባላት",
                total_members
            )

        with col2:

            st.metric(
                "💰 ጠቅላላ ቁጠባ",
                f"{total_savings:,.2f} ብር"
            )

        with col3:

            st.metric(
                "💵 ጠቅላላ ብድር",
                f"{total_loans:,.2f} ብር"
            )

        with col4:

            st.metric(
                "🔴 የቀረ ዋና",
                f"{total_remaining:,.2f} ብር"
            )

        st.divider()

        col1, col2, col3 = st.columns(3)

        with col1:

            st.metric(
                "🏦 10% የብድር ገቢ",
                f"{total_loan_fees:,.2f} ብር"
            )

        with col2:

            st.metric(
                "📈 የተሰበሰበ ወለድ",
                f"{total_interest:,.2f} ብር"
            )

        with col3:

            st.metric(
                "🏦 የብድር ገቢ ጠቅላላ",
                f"{total_loan_fees + total_interest:,.2f} ብር"
            )

        st.divider()

        st.dataframe(
            report_df,
            use_container_width=True,
            hide_index=True
        )

        # ----------------------------------------------------
        # DOWNLOAD REPORT
        # ----------------------------------------------------

        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            report_df.to_excel(
                writer,
                index=False,
                sheet_name="SACCO Report"
            )

        output.seek(0)

        st.download_button(
            label="📥 Excel ሪፖርት አውርድ",
            data=output,
            file_name="sacco_report.xlsx",
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            )
        )


# ============================================================
# 6. EDIT / DELETE MEMBER
# ============================================================

elif menu == "✏️ የአባላት መረጃ ማስተካከያ":

    st.header(
        "✏️ የአባላት መረጃ ማስተካከያ"
    )

    if not data:

        st.warning(
            "ምንም አባል አልተመዘገበም።"
        )

    else:

        member_id = st.selectbox(
            "አባል ይምረጡ",
            list(data.keys()),
            format_func=lambda x:
                f"{x} - "
                f"{data[x]['የአባል ስም']}"
        )

        member = data[member_id]

        st.subheader(
            f"👤 {member['የአባል ስም']}"
        )

        with st.form(
            "edit_member_form"
        ):

            new_name = st.text_input(
                "የአባል ስም",
                value=str(
                    member.get(
                        "የአባል ስም",
                        ""
                    )
                )
            ).strip()

            new_national_id = st.text_input(
                "National ID",
                value=str(
                    member.get(
                        "National ID",
                        ""
                    )
                )
            ).strip()

            new_photo = st.file_uploader(
                "አዲስ ፎቶ",
                type=[
                    "jpg",
                    "jpeg",
                    "png"
                ]
            )

            update = st.form_submit_button(
                "💾 መረጃ አዘምን"
            )

        if update:

            if not new_name:

                st.error(
                    "የአባል ስም ባዶ ሊሆን አይችልም።"
                )

            elif (
                new_national_id
                and national_id_exists(
                    new_national_id,
                    exclude_member_id=member_id
                )
            ):

                st.error(
                    "ይህ National ID "
                    "ለሌላ አባል ተመዝግቧል።"
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
                    ] = base64.b64encode(
                        photo_bytes
                    ).decode("utf-8")

                save_data_to_excel(
                    data
                )

                st.success(
                    "የአባሉ መረጃ "
                    "በትክክል ተዘምኗል።"
                )

                st.rerun()

        # ----------------------------------------------------
        # DELETE
        # ----------------------------------------------------

        st.divider()

        st.subheader(
            "🗑️ አባል ሰርዝ"
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
                data
            )

            st.success(
                "አባሉ ተሰርዟል።"
            )

            st.rerun()


# ============================================================
# 7. IMPORT EXCEL
# ============================================================

elif menu == "📤 Excel ፋይል አስገባ":

    st.header(
        "📤 Excel ፋይል አስገባ"
    )

    st.info(
        "Excel ፋይሉ ቢያንስ "
        "ID እና የአባል ስም "
        "እንዲኖረው ያስፈልጋል።"
    )

    uploaded_file = st.file_uploader(
        "Excel ፋይል ይምረጡ",
        type=["xlsx"]
    )

    if uploaded_file is not None:

        try:

            import_df = pd.read_excel(
                uploaded_file,
                dtype=str
            )

            import_df = import_df.fillna("")

            st.subheader(
                "👀 የExcel መረጃ"
            )

            st.dataframe(
                import_df,
                use_container_width=True,
                hide_index=True
            )

            # ----------------------------------------------
            # FIND ID
            # ----------------------------------------------

            id_column = None

            possible_id_columns = [
                "መታወቂያ ቁጥር (ID)",
                "ID",
                "ተ.ቁ"
            ]

            for col in possible_id_columns:

                if col in import_df.columns:

                    id_column = col
                    break

            # ----------------------------------------------
            # FIND NAME
            # ----------------------------------------------

            name_column = None

            possible_name_columns = [
                "የአባል ስም",
                "የሠራተኛው ስም",
                "Name",
                "name"
            ]

            for col in possible_name_columns:

                if col in import_df.columns:

                    name_column = col
                    break

            if id_column is None:

                st.error(
                    "Excel ፋይሉ "
                    "የID አምድ የለውም።"
                )

            elif name_column is None:

                st.error(
                    "Excel ፋይሉ "
                    "የአባል ስም አምድ የለውም።"
                )

            else:

                if st.button(
                    "📥 ወደ ሲስተሙ አስገባ"
                ):

                    imported_count = 0
                    skipped_count = 0
                    errors = []

                    ids = (
                        import_df[
                            id_column
                        ]
                        .astype(str)
                        .str.strip()
                    )

                    duplicate_ids = ids[
                        ids.duplicated()
                        & (ids != "")
                    ].unique().tolist()

                    if duplicate_ids:

                        st.error(
                            "በExcel ውስጥ "
                            "የተደጋገሙ IDዎች አሉ፦ "
                            + ", ".join(
                                duplicate_ids
                            )
                        )

                    else:

                        for _, row in import_df.iterrows():

                            new_id = str(
                                row[id_column]
                            ).strip()

                            new_name = str(
                                row[name_column]
                            ).strip()

                            if not new_id:

                                skipped_count += 1
                                continue

                            if not new_name:

                                skipped_count += 1
                                continue

                            if new_id in data:

                                skipped_count += 1
                                continue

                            imported_national_id = ""

                            if (
                                "National ID"
                                in import_df.columns
                            ):

                                imported_national_id = str(
                                    row[
                                        "National ID"
                                    ]
                                ).strip()

                            if (
                                imported_national_id
                                and national_id_exists(
                                    imported_national_id
                                )
                            ):

                                errors.append(
                                    f"{new_name} - "
                                    "National ID ይደገማል"
                                )

                                continue

                            member = empty_member()

                            member[
                                "የአባል ስም"
                            ] = new_name

                            member[
                                "National ID"
                            ] = imported_national_id

                            # ----------------------------------
                            # PRESERVE EXTRA EXCEL COLUMNS
                            # ----------------------------------

                            for column in import_df.columns:

                                if column in [
                                    id_column,
                                    name_column,
                                    "National ID"
                                ]:
                                    continue

                                value = row[
                                    column
                                ]

                                if value == "":
                                    continue

                                try:

                                    member[
                                        column
                                    ] = float(value)

                                except Exception:

                                    member[
                                        column
                                    ] = value

                            data[
                                new_id
                            ] = normalize_member(
                                member
                            )

                            imported_count += 1

                        save_data_to_excel(
                            data
                        )

                        st.success(
                            f"✅ {imported_count} "
                            "አባላት በትክክል ገብተዋል።"
                        )

                        if skipped_count > 0:

                            st.warning(
                                f"⚠️ {skipped_count} "
                                "ረድፎች ተዘልለዋል።"
                            )

                        if errors:

                            st.error(
                                "አንዳንድ ረድፎች "
                                "አልገቡም።"
                            )

                            for error in errors:

                                st.write(
                                    f"• {error}"
                                )

                        st.rerun()

        except Exception as e:

            st.error(
                f"Excel ፋይሉን ማንበብ "
                f"አልተቻለም፦ {e}"
            )


# ============================================================
# FOOTER
# ============================================================

st.sidebar.divider()

st.sidebar.caption(
    "🏦 ተስፋ የገንዘብ ቁጠባና ብድር ማህበር"
)

st.sidebar.caption(
    f"👥 አባላት፦ {len(data)}"
)

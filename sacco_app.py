import streamlit as st
import pandas as pd
import os
import base64

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

        df = df.set_index("መታወቂያ ቁጥር (ID)")

        data = df.to_dict(orient="index")

        # Normalize every member
        for member_id in data:
            data[member_id] = normalize_member(data[member_id])

        return data

    except Exception:
        return {}


# ============================================================
# NORMALIZE MEMBER DATA
# ============================================================

def normalize_member(member):
    defaults = {
        "የአባል ስም": "",
        "National ID": "",
        "የአባል ፎቶ": "",
        "ጠቅላላ ቁጠባ (ብር)": 0.0,
        "ብድር ሁኔታ": "የለም",
        "የተበደረው ጠቅላላ (ብር)": 0.0,
        "በእጅ የተሰጠ 90% (ብር)": 0.0,
        "የቀረው ዕዳ (ብር)": 0.0
    }

    for key, value in defaults.items():
        if key not in member:
            member[key] = value

    # Clean National ID
    if pd.isna(member["National ID"]):
        member["National ID"] = ""
    else:
        member["National ID"] = str(member["National ID"]).strip()

    # Clean name
    if pd.isna(member["የአባል ስም"]):
        member["የአባል ስም"] = ""
    else:
        member["የአባል ስም"] = str(member["የአባል ስም"]).strip()

    # Numeric fields
    numeric_fields = [
        "ጠቅላላ ቁጠባ (ብር)",
        "የተበደረው ጠቅላላ (ብር)",
        "በእጅ የተሰጠ 90% (ብር)",
        "የቀረው ዕዳ (ብር)"
    ]

    for field in numeric_fields:
        try:
            if pd.isna(member[field]):
                member[field] = 0.0
            else:
                member[field] = float(member[field])
        except:
            member[field] = 0.0

    if pd.isna(member["የአባል ፎቶ"]):
        member["የአባል ፎቶ"] = ""

    return member


# ============================================================
# SAVE DATABASE
# ============================================================

def save_data_to_excel(data):
    rows = []

    for member_id, member in data.items():
        row = member.copy()
        row["መታወቂያ ቁጥር (ID)"] = str(member_id)
        rows.append(row)

    if rows:
        df = pd.DataFrame(rows)

        # Keep ID as first column
        first_column = "መታወቂያ ቁጥር (ID)"

        other_columns = [
            col for col in df.columns
            if col != first_column
        ]

        df = df[[first_column] + other_columns]

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
                "በእጅ የተሰጠ 90% (ብር)",
                "የቀረው ዕዳ (ብር)"
            ]
        )

    df.to_excel(DB_FILE, index=False)


# ============================================================
# INITIALIZE DATABASE
# ============================================================

data = load_data()


# ============================================================
# NATIONAL ID DUPLICATE CHECK
# ============================================================

def national_id_exists(national_id, exclude_member_id=None):

    national_id = str(national_id).strip()

    # Empty National ID is allowed
    if not national_id:
        return False

    for member_id, member in data.items():

        if exclude_member_id is not None:
            if str(member_id) == str(exclude_member_id):
                continue

        existing_id = str(
            member.get("National ID", "")
        ).strip()

        if existing_id and existing_id == national_id:
            return True

    return False


# ============================================================
# LOAN PAYMENT FUNCTION
# ============================================================

def apply_loan_payment(member, payment_amount):

    payment_amount = float(payment_amount)

    remaining_debt = float(
        member.get("የቀረው ዕዳ (ብር)", 0)
    )

    original_loan = float(
        member.get("የተበደረው ጠቅላላ (ብር)", 0)
    )

    if remaining_debt <= 0:
        member["ብድር ሁኔታ"] = "የለም"
        member["የቀረው ዕዳ (ብር)"] = 0.0
        return payment_amount

    # 2% monthly interest based on original loan
    monthly_interest = original_loan * 0.02

    # Interest is paid first
    interest_paid = min(payment_amount, monthly_interest)

    remaining_payment = payment_amount - interest_paid

    # Remaining amount reduces principal
    principal_paid = min(
        remaining_payment,
        remaining_debt
    )

    new_debt = remaining_debt - principal_paid

    member["የቀረው ዕዳ (ብር)"] = max(new_debt, 0)

    if member["የቀረው ዕዳ (ብር)"] <= 0:
        member["የቀረው ዕዳ (ብር)"] = 0.0
        member["ብድር ሁኔታ"] = "የለም"
    else:
        member["ብድር ሁኔታ"] = "ያለ ብድር"

    # Return excess money after debt is cleared
    total_used = interest_paid + principal_paid
    excess = payment_amount - total_used

    return max(excess, 0)


# ============================================================
# SIDEBAR MENU
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

    with st.form("member_registration_form"):

        col1, col2 = st.columns(2)

        with col1:

            member_id = st.text_input(
                "ID *",
                placeholder="ለምሳሌ 1"
            ).strip()

            member_name = st.text_input(
                "የአባል ስም *"
            ).strip()

            national_id = st.text_input(
                "National ID (አማራጭ)"
            ).strip()

        with col2:

            photo = st.file_uploader(
                "የአባል ፎቶ (አማራጭ)",
                type=["jpg", "jpeg", "png"]
            )

        submit = st.form_submit_button(
            "➕ አባል መዝግብ"
        )

    if submit:

        if not member_id:
            st.error("ID ያስገቡ።")

        elif not member_name:
            st.error("የአባል ስም ያስገቡ።")

        elif member_id in data:
            st.error("ይህ ID ቀድሞ ተመዝግቧል።")

        elif national_id and national_id_exists(national_id):
            st.error(
                "ይህ National ID ቀድሞ ተመዝግቧል። "
                "ሌላ National ID ያስገቡ።"
            )

        else:

            photo_data = ""

            if photo is not None:
                photo_bytes = photo.read()
                photo_data = base64.b64encode(
                    photo_bytes
                ).decode("utf-8")

            data[member_id] = {
                "የአባል ስም": member_name,
                "National ID": national_id,
                "የአባል ፎቶ": photo_data,
                "ጠቅላላ ቁጠባ (ብር)": 0.0,
                "ብድር ሁኔታ": "የለም",
                "የተበደረው ጠቅላላ (ብር)": 0.0,
                "በእጅ የተሰጠ 90% (ብር)": 0.0,
                "የቀረው ዕዳ (ብር)": 0.0
            }

            save_data_to_excel(data)

            st.success(
                f"አባል {member_name} በትክክል ተመዝግቧል።"
            )

            st.rerun()


# ============================================================
# 2. MONTHLY SAVINGS
# ============================================================

elif menu == "💰 የወር ቁጠባ ማስገቢያ":

    st.header("💰 የወር ቁጠባ ማስገቢያ")

    if not data:
        st.warning("ምንም አባል አልተመዘገበም።")

    else:

        member_id = st.selectbox(
            "አባል ይምረጡ",
            list(data.keys()),
            format_func=lambda x:
                f"{x} - {data[x]['የአባል ስም']}"
        )

        member = data[member_id]

        st.info(
            f"👤 አባል: {member['የአባል ስም']}"
        )

        st.write(
            f"💰 አሁን ያለ ቁጠባ: "
            f"{member['ጠቅላላ ቁጠባ (ብር)']:,.2f} ብር"
        )

        if member["የቀረው ዕዳ (ብር)"] > 0:

            st.warning(
                f"የቀረ ዕዳ: "
                f"{member['የቀረው ዕዳ (ብር)']:,.2f} ብር"
            )

        month = st.selectbox(
            "ወር ይምረጡ",
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

        amount = st.number_input(
            "የቁጠባ / ክፍያ መጠን (ብር)",
            min_value=0.0,
            step=100.0
        )

        if st.button("💾 ቁጠባ መዝግብ"):

            if amount <= 0:
                st.error("የገንዘብ መጠን ያስገቡ።")

            else:

                # If there is outstanding debt,
                # payment goes toward the loan first.
                if member["የቀረው ዕዳ (ብር)"] > 0:

                    excess = apply_loan_payment(
                        member,
                        amount
                    )

                    if excess > 0:
                        member[
                            "ጠቅላላ ቁጠባ (ብር)"
                        ] += excess

                    st.success(
                        f"{amount:,.2f} ብር ተመዝግቧል። "
                        f"የቀረ ዕዳ: "
                        f"{member['የቀረው ዕዳ (ብር)']:,.2f} ብር"
                    )

                else:

                    member[
                        "ጠቅላላ ቁጠባ (ብር)"
                    ] += amount

                    st.success(
                        f"{amount:,.2f} ብር "
                        f"በ{month} ተመዝግቧል።"
                    )

                # Store monthly amount
                member[month] = (
                    float(member.get(month, 0) or 0)
                    + amount
                )

                save_data_to_excel(data)

                st.rerun()


# ============================================================
# 3. LOAN SERVICE
# ============================================================

elif menu == "💵 የብድር አገልግሎት":

    st.header("💵 የብድር አገልግሎት")

    if not data:
        st.warning("ምንም አባል አልተመዘገበም።")

    else:

        member_id = st.selectbox(
            "አባል ይምረጡ",
            list(data.keys()),
            format_func=lambda x:
                f"{x} - {data[x]['የአባል ስም']}"
        )

        member = data[member_id]

        savings = float(
            member["ጠቅላላ ቁጠባ (ብር)"]
        )

        max_loan = savings * 4

        st.write(
            f"👤 አባል: **{member['የአባል ስም']}**"
        )

        st.write(
            f"💰 ጠቅላላ ቁጠባ: "
            f"**{savings:,.2f} ብር**"
        )

        st.write(
            f"💵 ከፍተኛ የሚበደር መጠን: "
            f"**{max_loan:,.2f} ብር**"
        )

        current_debt = float(
            member["የቀረው ዕዳ (ብር)"]
        )

        if current_debt > 0:

            st.error(
                f"ይህ አባል አሁን "
                f"{current_debt:,.2f} ብር ዕዳ አለበት። "
                f"አዲስ ብድር መውሰድ አይችልም።"
            )

        else:

            loan_amount = st.number_input(
                "የሚፈለገው ብድር (ብር)",
                min_value=0.0,
                max_value=max_loan,
                step=100.0
            )

            if loan_amount > 0:

                fee = loan_amount * 0.10
                net_payment = loan_amount * 0.90

                monthly_principal = loan_amount / 36
                monthly_interest = loan_amount * 0.02
                monthly_payment = (
                    monthly_principal +
                    monthly_interest
                )

                st.subheader("📋 የብድር ማጠቃለያ")

                col1, col2, col3 = st.columns(3)

                with col1:
                    st.metric(
                        "የተበደረው",
                        f"{loan_amount:,.2f} ብር"
                    )

                with col2:
                    st.metric(
                        "10% ክፍያ",
                        f"{fee:,.2f} ብር"
                    )

                with col3:
                    st.metric(
                        "በእጅ የሚሰጥ 90%",
                        f"{net_payment:,.2f} ብር"
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

                if st.button("💵 ብድር ስጥ"):

                    if loan_amount > max_loan:
                        st.error(
                            "ከተፈቀደው የብድር መጠን በላይ ነው።"
                        )

                    elif savings <= 0:
                        st.error(
                            "ቁጠባ ሳይኖር ብድር መውሰድ አይቻልም።"
                        )

                    else:

                        member[
                            "የተበደረው ጠቅላላ (ብር)"
                        ] = loan_amount

                        member[
                            "በእጅ የተሰጠ 90% (ብር)"
                        ] = net_payment

                        member[
                            "የቀረው ዕዳ (ብር)"
                        ] = loan_amount

                        member[
                            "ብድር ሁኔታ"
                        ] = "ያለ ብድር"

                        save_data_to_excel(data)

                        st.success(
                            f"{loan_amount:,.2f} ብር ብድር "
                            f"ተመዝግቧል።"
                        )

                        st.rerun()


# ============================================================
# 4. LOAN REPAYMENT
# ============================================================

elif menu == "📅 የብድር ክፍያ መመዝገቢያ":

    st.header("📅 የብድር ክፍያ መመዝገቢያ")

    members_with_debt = [
        member_id
        for member_id, member in data.items()
        if float(member.get(
            "የቀረው ዕዳ (ብር)", 0
        )) > 0
    ]

    if not members_with_debt:

        st.info("አሁን የቀረ ዕዳ ያለበት አባል የለም።")

    else:

        member_id = st.selectbox(
            "አባል ይምረጡ",
            members_with_debt,
            format_func=lambda x:
                f"{x} - {data[x]['የአባል ስም']}"
        )

        member = data[member_id]

        st.write(
            f"👤 አባል: **{member['የአባል ስም']}**"
        )

        st.write(
            f"💵 የመጀመሪያ ብድር: "
            f"{member['የተበደረው ጠቅላላ (ብር)']:,.2f} ብር"
        )

        st.write(
            f"🔴 የቀረ ዕዳ: "
            f"**{member['የቀረው ዕዳ (ብር)']:,.2f} ብር**"
        )

        payment = st.number_input(
            "የክፍያ መጠን (ብር)",
            min_value=0.0,
            step=100.0
        )

        if st.button("💾 ክፍያ መዝግብ"):

            if payment <= 0:
                st.error(
                    "የክፍያ መጠን ያስገቡ።"
                )

            else:

                old_debt = member[
                    "የቀረው ዕዳ (ብር)"
                ]

                excess = apply_loan_payment(
                    member,
                    payment
                )

                if excess > 0:

                    member[
                        "ጠቅላላ ቁጠባ (ብር)"
                    ] += excess

                save_data_to_excel(data)

                st.success(
                    f"{payment:,.2f} ብር ክፍያ ተመዝግቧል።"
                )

                st.write(
                    f"ቀድሞ ዕዳ: {old_debt:,.2f} ብር"
                )

                st.write(
                    f"አሁን የቀረ ዕዳ: "
                    f"{member['የቀረው ዕዳ (ብር)']:,.2f} ብር"
                )

                if excess > 0:
                    st.info(
                        f"{excess:,.2f} ብር ተረፍቶ "
                        f"ወደ ቁጠባ ተጨምሯል።"
                    )

                st.rerun()


# ============================================================
# 5. GENERAL REPORT
# ============================================================

elif menu == "📊 ጠቅላላ ሪፖርት":

    st.header("📊 ጠቅላላ ሪፖርት")

    if not data:

        st.warning("ምንም የአባላት መረጃ የለም።")

    else:

        report_rows = []

        for member_id, member in data.items():

            report_rows.append({
                "ID": member_id,
                "የአባል ስም": member.get(
                    "የአባል ስም", ""
                ),
                "National ID": member.get(
                    "National ID", ""
                ),
                "ጠቅላላ ቁጠባ (ብር)": member.get(
                    "ጠቅላላ ቁጠባ (ብር)", 0
                ),
                "ብድር ሁኔታ": member.get(
                    "ብድር ሁኔታ", "የለም"
                ),
                "የተበደረው ጠቅላላ (ብር)": member.get(
                    "የተበደረው ጠቅላላ (ብር)", 0
                ),
                "በእጅ የተሰጠ 90% (ብር)": member.get(
                    "በእጅ የተሰጠ 90% (ብር)", 0
                ),
                "የቀረው ዕዳ (ብር)": member.get(
                    "የቀረው ዕዳ (ብር)", 0
                )
            })

        report_df = pd.DataFrame(report_rows)

        # Search
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

            report_df = report_df[mask]

        # Summary
        total_members = len(report_df)

        total_savings = report_df[
            "ጠቅላላ ቁጠባ (ብር)"
        ].sum()

        total_loans = report_df[
            "የተበደረው ጠቅላላ (ብር)"
        ].sum()

        total_debt = report_df[
            "የቀረው ዕዳ (ብር)"
        ].sum()

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
                "🔴 የቀረ ዕዳ",
                f"{total_debt:,.2f} ብር"
            )

        st.divider()

        st.dataframe(
            report_df,
            use_container_width=True,
            hide_index=True
        )

        # ====================================================
        # DOWNLOAD EXCEL
        # ====================================================

        download_df = report_df.copy()

        excel_file = "sacco_report.xlsx"

        download_df.to_excel(
            excel_file,
            index=False
        )

        with open(
            excel_file,
            "rb"
        ) as file:

            st.download_button(
                label="📥 Excel ሪፖርት አውርድ",
                data=file,
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

    st.header("✏️ የአባላት መረጃ ማስተካከያ")

    if not data:

        st.warning("ምንም አባል አልተመዘገበም።")

    else:

        member_id = st.selectbox(
            "አባል ይምረጡ",
            list(data.keys()),
            format_func=lambda x:
                f"{x} - {data[x]['የአባል ስም']}"
        )

        member = data[member_id]

        st.subheader(
            f"👤 {member['የአባል ስም']}"
        )

        with st.form("edit_member_form"):

            new_name = st.text_input(
                "የአባል ስም",
                value=str(
                    member.get(
                        "የአባል ስም", ""
                    )
                )
            ).strip()

            new_national_id = st.text_input(
                "National ID (አማራጭ)",
                value=str(
                    member.get(
                        "National ID", ""
                    )
                )
            ).strip()

            new_photo = st.file_uploader(
                "አዲስ ፎቶ (አማራጭ)",
                type=["jpg", "jpeg", "png"]
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
                    "ይህ National ID ቀድሞ "
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

                    photo_bytes = new_photo.read()

                    member[
                        "የአባል ፎቶ"
                    ] = base64.b64encode(
                        photo_bytes
                    ).decode("utf-8")

                save_data_to_excel(data)

                st.success(
                    "የአባሉ መረጃ በትክክል ተዘምኗል።"
                )

                st.rerun()

        st.divider()

        st.subheader("🗑️ አባል ሰርዝ")

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
                "አባሉ ተሰርዟል።"
            )

            st.rerun()


# ============================================================
# 7. IMPORT EXCEL
# ============================================================

elif menu == "📤 Excel ፋይል አስገባ":

    st.header("📤 Excel ፋይል አስገባ")

    st.info(
        "Excel ፋይሉ ቢያንስ "
        "ID እና የአባል ስም እንዲኖረው ያስፈልጋል። "
        "National ID አማራጭ ነው።"
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

            st.subheader("👀 የExcel መረጃ")

            st.dataframe(
                import_df,
                use_container_width=True,
                hide_index=True
            )

            # Find ID column
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

            # Find name column
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
                    "Excel ፋይሉ የID አምድ የለውም።"
                )

            elif name_column is None:

                st.error(
                    "Excel ፋይሉ የአባል ስም አምድ የለውም።"
                )

            else:

                if st.button(
                    "📥 ወደ ሲስተሙ አስገባ"
                ):

                    imported_count = 0
                    skipped_count = 0
                    errors = []

                    # Check duplicate IDs inside Excel
                    ids = (
                        import_df[id_column]
                        .astype(str)
                        .str.strip()
                    )

                    duplicate_ids = ids[
                        ids.duplicated()
                        & (ids != "")
                    ].unique().tolist()

                    if duplicate_ids:

                        st.error(
                            "በExcel ውስጥ የተደጋገሙ IDዎች አሉ፦ "
                            + ", ".join(duplicate_ids)
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

                            if "National ID" in import_df.columns:
                                imported_national_id = str(
                                    row["National ID"]
                                ).strip()

                            # National ID is optional,
                            # but if supplied it must be unique.
                            if (
                                imported_national_id
                                and national_id_exists(
                                    imported_national_id
                                )
                            ):
                                errors.append(
                                    f"{new_name} - "
                                    f"National ID ይደገማል"
                                )
                                continue

                            data[new_id] = {
                                "የአባል ስም": new_name,
                                "National ID":
                                    imported_national_id,
                                "የአባል ፎቶ": "",
                                "ጠቅላላ ቁጠባ (ብር)": 0.0,
                                "ብድር ሁኔታ": "የለም",
                                "የተበደረው ጠቅላላ (ብር)": 0.0,
                                "በእጅ የተሰጠ 90% (ብር)": 0.0,
                                "የቀረው ዕዳ (ብር)": 0.0
                            }

                            # Preserve extra Excel columns
                            for column in import_df.columns:

                                if column not in [
                                    id_column,
                                    name_column,
                                    "National ID"
                                ]:

                                    value = row[column]

                                    if value != "":

                                        try:
                                            numeric_value = float(
                                                value
                                            )
                                            data[new_id][
                                                column
                                            ] = numeric_value

                                        except:
                                            data[new_id][
                                                column
                                            ] = value

                            imported_count += 1

                        save_data_to_excel(data)

                        st.success(
                            f"✅ {imported_count} አባላት "
                            f"በትክክል ገብተዋል።"
                        )

                        if skipped_count > 0:

                            st.warning(
                                f"⚠️ {skipped_count} ረድፎች "
                                f"ተዘልለዋል። "
                                f"(ID ቀድሞ ካለ ወይም መረጃ ካጣ)"
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
                f"Excel ፋይሉን ማንበብ አልተቻለም፦ {e}"
            )


# ============================================================
# FOOTER
# ============================================================

st.sidebar.divider()

st.sidebar.caption(
    "🏦 ተስፋ የገንዘብ ቁጠባና ብድር ማህበር"
)

st.sidebar.caption(
    f"👥 አባላት: {len(data)}"
)

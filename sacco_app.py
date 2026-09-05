import streamlit as st
import pandas as pd
import os

st.set_page_config(
    page_title="የቁጠባ እና ብድር ሲስተም",
    layout="wide"
)

st.title("🏦 ተስፋ የገንዘብ ቁጠባና ብድር ማህበር")


# ============================================================
# 💾 DATABASE
# ============================================================

DB_FILE = "sacco_database.xlsx"

if os.path.exists(DB_FILE):
    try:
        df_load = pd.read_excel(
            DB_FILE,
            dtype={"መታወቂያ ቁጥር (ID)": str}
        )

        members_db = df_load.set_index(
            "መታወቂያ ቁጥር (ID)"
        ).to_dict(orient="index")

    except Exception:
        members_db = {}
else:
    members_db = {}


# ============================================================
# 🔧 MAKE OLD DATA COMPATIBLE
# ============================================================

# Existing members may not have National ID because
# National ID is being added now.
for member_id, member in members_db.items():

    if "ብሔራዊ መታወቂያ (National ID)" not in member:
        member["ብሔራዊ መታወቂያ (National ID)"] = ""

    if "የአባል ስም" not in member:
        member["የአባል ስም"] = ""

    if "ጠቅላላ ቁጠባ (ብር)" not in member:
        member["ጠቅላላ ቁጠባ (ብር)"] = 0.0

    if "ብድር ሁኔታ" not in member:
        member["ብድር ሁኔታ"] = "የለበትም"

    if "የተበደረው ጠቅላላ (ብር)" not in member:
        member["የተበደረው ጠቅላላ (ብር)"] = 0.0

    if "በእጅ የተሰጠ 90% (ብር)" not in member:
        member["በእጅ የተሰጠ 90% (ብር)"] = 0.0

    if "የቀረው ዕዳ (ብር)" not in member:
        member["የቀረው ዕዳ (ብር)"] = 0.0


# ============================================================
# 💾 SAVE DATA
# ============================================================

def save_data_to_excel(db):

    if db:

        df_save = pd.DataFrame.from_dict(
            db,
            orient="index"
        )

        df_save.index.name = "መታወቂያ ቁጥር (ID)"

        df_save.reset_index().to_excel(
            DB_FILE,
            index=False
        )

    else:

        if os.path.exists(DB_FILE):

            try:
                os.remove(DB_FILE)

            except Exception:
                pass


# ============================================================
# 🔧 LOAN PAYMENT FUNCTION
# ============================================================

def apply_loan_payment(member, payment_amount):

    original_loan = float(
        member.get("የተበደረው ጠቅላላ (ብር)", 0)
    )

    outstanding_debt = float(
        member.get("የቀረው ዕዳ (ብር)", 0)
    )

    # 2% monthly interest based on original loan
    monthly_interest = original_loan * 0.02

    # Money remaining after interest goes to principal
    principal_payment = payment_amount - monthly_interest

    if principal_payment < 0:
        principal_payment = 0.0

    # If payment can clear the remaining debt
    if principal_payment >= outstanding_debt:

        excess = principal_payment - outstanding_debt

        member["የቀረው ዕዳ (ብር)"] = 0.0
        member["ብድር ሁኔታ"] = "የለበትም"

        # Extra money goes to savings
        if excess > 0:
            member["ጠቅላላ ቁጠባ (ብር)"] = (
                float(member.get("ጠቅላላ ቁጠባ (ብር)", 0))
                + excess
            )

        return {
            "interest": monthly_interest,
            "principal": outstanding_debt,
            "excess": excess,
            "remaining": 0.0,
            "finished": True
        }

    else:

        member["የቀረው ዕዳ (ብር)"] = (
            outstanding_debt - principal_payment
        )

        return {
            "interest": monthly_interest,
            "principal": principal_payment,
            "excess": 0.0,
            "remaining": member["የቀረው ዕዳ (ብር)"],
            "finished": False
        }


# ============================================================
# 📋 SIDEBAR MENU
# ============================================================

menu = st.sidebar.selectbox(
    "ያሉ አማራጮች",
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
# 1. 👤 MEMBER REGISTRATION
# ============================================================

if menu == "👤 አባል መመዝገቢያ":

    st.header("👤 አዲስ አባል መመዝገቢያ ፎርም")

    m_id = st.text_input(
        "የአባል መታወቂያ ቁጥር (ID):"
    )

    m_name = st.text_input(
        "የአባል ሙሉ ስም:"
    )

    # National ID is mandatory
    national_id = st.text_input(
        "🪪 ብሔራዊ መታወቂያ (National ID):"
    )

    st.info(
        "ℹ️ ብሔራዊ መታወቂያ ማስገባት ግዴታ ነው።"
    )

    register_btn = st.button(
        "አባል መዝግብ",
        use_container_width=True
    )

    if register_btn:

        m_id = m_id.strip()
        m_name = m_name.strip()
        national_id = national_id.strip()

        # All three are mandatory
        if not m_id or not m_name or not national_id:

            st.warning(
                "⚠️ እባክዎ የአባል ID፣ ሙሉ ስም እና "
                "ብሔራዊ መታወቂያን ሁሉንም ያስገቡ!"
            )

        elif m_id in members_db:

            st.error(
                f"❌ ስህተት፦ መታወቂያ ቁጥር {m_id} "
                "ቀደም ብሎ ተመዝግቧል!"
            )

        else:

            # Check National ID uniqueness
            national_id_exists = False

            for member in members_db.values():

                existing_national_id = str(
                    member.get(
                        "ብሔራዊ መታወቂያ (National ID)",
                        ""
                    )
                ).strip()

                if (
                    existing_national_id
                    and existing_national_id == national_id
                ):
                    national_id_exists = True
                    break

            if national_id_exists:

                st.error(
                    "❌ ይህ ብሔራዊ መታወቂያ "
                    "ቀደም ብሎ ተመዝግቧል!"
                )

            else:

                members_db[m_id] = {

                    "የአባል ስም": m_name,

                    "ብሔራዊ መታወቂያ (National ID)":
                        national_id,

                    "ጠቅላላ ቁጠባ (ብር)": 0.0,

                    "ብድር ሁኔታ":
                        "የለበትም",

                    "የተበደረው ጠቅላላ (ብር)":
                        0.0,

                    "በእጅ የተሰጠ 90% (ብር)":
                        0.0,

                    "የቀረው ዕዳ (ብር)":
                        0.0
                }

                save_data_to_excel(members_db)

                st.success(
                    f"✅ አባል {m_name} "
                    "በተሳካ ሁኔታ ተመዝግቧል!"
                )


# ============================================================
# 2. 💰 SAVINGS
# ============================================================

elif menu == "💰 የወር ቁጠባ ማስገቢያ":

    st.header("💰 የወርሃዊ ቁጠባ / ብድር ክፍያ")

    s_id = st.text_input(
        "የአባል መታወቂያ (ID):"
    )

    amount = st.number_input(
        "የሚገባው የገንዘብ መጠን (ብር):",
        min_value=0.0,
        step=100.0
    )

    save_btn = st.button(
        "ቁጠባ / ክፍያ መዝግብ",
        use_container_width=True
    )

    if save_btn:

        s_id = s_id.strip()

        if not s_id:

            st.warning(
                "⚠️ እባክዎ የአባል ID ያስገቡ!"
            )

        elif amount <= 0:

            st.warning(
                "⚠️ የገንዘብ መጠኑ ከ0 በላይ መሆን አለበት!"
            )

        elif s_id in members_db:

            member = members_db[s_id]

            # Check outstanding loan
            outstanding_debt = float(
                member.get(
                    "የቀረው ዕዳ (ብር)",
                    0
                )
            )

            loan_status = member.get(
                "ብድር ሁኔታ",
                "የለበትም"
            )

            if (
                loan_status == "ያለበት"
                and outstanding_debt > 0
            ):

                # ====================================================
                # MEMBER HAS LOAN
                # MONEY GOES TO LOAN REPAYMENT
                # ====================================================

                result = apply_loan_payment(
                    member,
                    amount
                )

                save_data_to_excel(
                    members_db
                )

                if result["finished"]:

                    st.success(
                        f"🎉 {member['የአባል ስም']} "
                        "ብድሩን ሙሉ በሙሉ ከፍሏል!"
                    )

                    st.info(
                        f"💵 የተከፈለው፦ "
                        f"{amount:,.2f} ብር\n\n"
                        f"💰 ወደ ዋና ዕዳ የገባው፦ "
                        f"{result['principal']:,.2f} ብር\n\n"
                        f"💸 ወደ ወለድ የገባው፦ "
                        f"{result['interest']:,.2f} ብር"
                    )

                    if result["excess"] > 0:

                        st.success(
                            f"💰 ብድሩ ከተዘጋ በኋላ "
                            f"የቀረው {result['excess']:,.2f} ብር "
                            "ወደ ቁጠባ ተጨምሯል!"
                        )

                        st.info(
                            f"🏦 አዲሱ ጠቅላላ ቁጠባ፦ "
                            f"{member['ጠቅላላ ቁጠባ (ብር)']:,.2f} ብር"
                        )

                else:

                    st.success(
                        f"✅ {amount:,.2f} ብር "
                        "ወደ ብድር ክፍያ ተመዝግቧል።"
                    )

                    st.info(
                        f"📌 የቀረው ዕዳ፦ "
                        f"{result['remaining']:,.2f} ብር"
                    )

            else:

                # ====================================================
                # NO LOAN
                # MONEY GOES TO SAVINGS
                # ====================================================

                member["ጠቅላላ ቁጠባ (ብር)"] = (
                    float(
                        member.get(
                            "ጠቅላላ ቁጠባ (ብር)",
                            0
                        )
                    )
                    + amount
                )

                save_data_to_excel(
                    members_db
                )

                st.success(
                    f"✅ ለ{member['የአባል ስም']} "
                    f"{amount:,.2f} ብር ቁጠባ "
                    "ተመዝግቧል።"
                )

                st.info(
                    f"🏦 ጠቅላላ ቁጠባ፦ "
                    f"{member['ጠቅላላ ቁጠባ (ብር)']:,.2f} ብር"
                )

        else:

            st.error(
                "❌ ይህ መታወቂያ በሲስተሙ "
                "ውስጥ አልተገኘም!"
            )


# ============================================================
# 3. 💵 LOAN
# ============================================================

elif menu == "💵 የብድር አገልግሎት":

    st.header("💵 የብድር ማመልከቻ እና ስሌት")

    loan_id = st.text_input(
        "የተበዳሪው አባል መታወቂያ (ID):"
    )

    loan_amount = st.number_input(
        "የሚጠይቀው የብድር መጠን (ብር):",
        min_value=0.0,
        step=1000.0
    )

    # Show maximum loan when member ID is entered
    if loan_id.strip() in members_db:

        member_preview = members_db[
            loan_id.strip()
        ]

        current_savings = float(
            member_preview.get(
                "ጠቅላላ ቁጠባ (ብር)",
                0
            )
        )

        maximum_loan = current_savings * 4

        st.info(
            f"🏦 የአሁኑ ቁጠባ፦ "
            f"**{current_savings:,.2f} ብር**\n\n"
            f"📈 ከቁጠባ 4 እጥፍ የሚፈቀደው "
            f"ከፍተኛ ብድር፦ "
            f"**{maximum_loan:,.2f} ብር**"
        )

    calculate_btn = st.button(
        "ብድር አስላ እና ፍቀድ",
        use_container_width=True
    )

    if calculate_btn:

        loan_id = loan_id.strip()

        if not loan_id:

            st.warning(
                "⚠️ እባክዎ የአባል ID ያስገቡ!"
            )

        elif loan_amount <= 0:

            st.warning(
                "⚠️ የብድር መጠኑ ከ0 በላይ መሆን አለበት!"
            )

        elif loan_id in members_db:

            member = members_db[loan_id]

            # ====================================================
            # EXISTING LOAN CHECK
            # ====================================================

            if member.get("ብድር ሁኔታ") == "ያለበት":

                st.error(
                    f"❌ ስህተት፦ {member['የአባል ስም']} "
                    "የድሮ ብድር ስላለበት "
                    "ተጨማሪ መበደር አይችልም!"
                )

            else:

                current_savings = float(
                    member.get(
                        "ጠቅላላ ቁጠባ (ብር)",
                        0
                    )
                )

                # ====================================================
                # MAXIMUM LOAN = 4 × SAVINGS
                # ====================================================

                maximum_loan = current_savings * 4

                if loan_amount > maximum_loan:

                    st.error(
                        f"❌ የብድር ጥያቄው አልተፈቀደም!\n\n"
                        f"🏦 የአባሉ ቁጠባ፦ "
                        f"{current_savings:,.2f} ብር\n\n"
                        f"📈 ከፍተኛ የሚፈቀደው ብድር "
                        f"(4 × ቁጠባ)፦ "
                        f"{maximum_loan:,.2f} ብር\n\n"
                        f"💵 የተጠየቀው፦ "
                        f"{loan_amount:,.2f} ብር"
                    )

                else:

                    upfront_fee = loan_amount * 0.10

                    net_payout = (
                        loan_amount
                        - upfront_fee
                    )

                    monthly_principal = (
                        loan_amount / 36
                    )

                    monthly_interest = (
                        loan_amount * 0.02
                    )

                    total_monthly = (
                        monthly_principal
                        + monthly_interest
                    )

                    member["ብድር ሁኔታ"] = (
                        "ያለበት"
                    )

                    member[
                        "የተበደረው ጠቅላላ (ብር)"
                    ] = loan_amount

                    member[
                        "በእጅ የተሰጠ 90% (ብር)"
                    ] = net_payout

                    member[
                        "የቀረው ዕዳ (ብር)"
                    ] = loan_amount

                    save_data_to_excel(
                        members_db
                    )

                    st.success(
                        f"🎉 ለ{member['የአባል ስም']} "
                        "ብድር ተፈቅዷል!"
                    )

                    st.info(
                        f"🏦 ቁጠባ፦ "
                        f"{current_savings:,.2f} ብር\n\n"
                        f"📈 ከፍተኛ ብድር፦ "
                        f"{maximum_loan:,.2f} ብር\n\n"
                        f"💵 የተፈቀደው ብድር፦ "
                        f"{loan_amount:,.2f} ብር\n\n"
                        f"💰 በእጅ የሚሰጠው (90%)፦ "
                        f"{net_payout:,.2f} ብር\n\n"
                        f"📅 የወርሃዊ ክፍያ "
                        f"(ለ36 ወራት)፦ "
                        f"{total_monthly:,.2f} ብር\n\n"
                        f"ዋና፦ {monthly_principal:,.2f} ብር\n\n"
                        f"ወለድ 2%፦ "
                        f"{monthly_interest:,.2f} ብር"
                    )

        else:

            st.error(
                "❌ ይህ መታወቂያ በሲስተሙ "
                "ውስጥ አልተገኘም!"
            )


# ============================================================
# 4. 📅 LOAN REPAYMENT
# ============================================================

elif menu == "📅 የብድር ክፍያ መመዝገቢያ":

    st.header("📅 የወርሃዊ ብድር ክፍያ መቀበያ")

    p_id = st.text_input(
        "የከፋይ አባል መታወቂያ (ID):"
    )

    p_id = p_id.strip()

    if p_id in members_db:

        member = members_db[p_id]

        if member.get("ብድር ሁኔታ") == "ያለበት":

            outstanding_debt = float(
                member.get(
                    "የቀረው ዕዳ (ብር)",
                    0
                )
            )

            st.warning(
                f"📌 {member['የአባል ስም']} "
                f"ያለበት ጠቅላላ ዕዳ፦ "
                f"{outstanding_debt:,.2f} ብር"
            )

            original_loan = float(
                member.get(
                    "የተበደረው ጠቅላላ (ብር)",
                    0
                )
            )

            monthly_principal = (
                original_loan / 36
            )

            monthly_interest = (
                original_loan * 0.02
            )

            normal_monthly_payment = (
                monthly_principal
                + monthly_interest
            )

            st.write(
                f"💡 መደበኛ የወር ክፍያ፦ "
                f"**{normal_monthly_payment:,.2f} ብር** "
                f"(ዋና፦ {monthly_principal:,.2f} + "
                f"ወለድ፦ {monthly_interest:,.2f})"
            )

            pay_amount = st.number_input(
                "አባል አሁን የከፈለው ጠቅላላ ብር "
                "(ዋና + ወለድ)፦",
                min_value=0.0,
                step=100.0
            )

            pay_btn = st.button(
                "ክፍያ መዝግብ",
                use_container_width=True
            )

            if pay_btn:

                if pay_amount <= 0:

                    st.warning(
                        "⚠️ የክፍያ መጠኑ ከ0 በላይ "
                        "መሆን አለበት!"
                    )

                else:

                    result = apply_loan_payment(
                        member,
                        pay_amount
                    )

                    save_data_to_excel(
                        members_db
                    )

                    if result["finished"]:

                        st.success(
                            f"🎉 {member['የአባል ስም']} "
                            "ብድሩን ሙሉ በሙሉ "
                            "ከፍሎ ጨርሷል!"
                        )

                        st.info(
                            f"💵 ጠቅላላ የተከፈለ፦ "
                            f"{pay_amount:,.2f} ብር\n\n"
                            f"💰 ወደ ዋና ዕዳ፦ "
                            f"{result['principal']:,.2f} ብር\n\n"
                            f"💸 ወደ ወለድ፦ "
                            f"{result['interest']:,.2f} ብር"
                        )

                        if result["excess"] > 0:

                            st.success(
                                f"💰 ከብድሩ ሙሉ ክፍያ በኋላ "
                                f"{result['excess']:,.2f} ብር "
                                "ትርፍ ቀርቷል። "
                                "ወደ ቁጠባ ተጨምሯል!"
                            )

                            st.info(
                                f"🏦 አዲሱ ጠቅላላ ቁጠባ፦ "
                                f"{member['ጠቅላላ ቁጠባ (ብር)']:,.2f} ብር"
                            )

                    else:

                        st.success(
                            f"✅ ክፍያ ተመዝግቧል።"
                        )

                        st.info(
                            f"💰 ወደ ዋና ዕዳ የገባው፦ "
                            f"{result['principal']:,.2f} ብር\n\n"
                            f"💸 ወለድ፦ "
                            f"{result['interest']:,.2f} ብር\n\n"
                            f"📌 የቀረው ዋና ዕዳ፦ "
                            f"{result['remaining']:,.2f} ብር"
                        )

        else:

            st.info(
                f"💡 {member['የአባል ስም']} ላይ "
                "ምንም ዓይነት የብድር ዕዳ የለም።"
            )

    elif p_id:

        st.error(
            "❌ ይህ መታወቂያ በሲስተሙ "
            "ውስጥ አልተገኘም!"
        )


# ============================================================
# 5. 📊 REPORT
# ============================================================

elif menu == "📊 ጠቅላላ ሪፖርት":

    st.header(
        "📊 ጠቅላላ የአባላት፣ የቁጠባ እና "
        "የብድር ሪፖርት"
    )

    if members_db:

        df = pd.DataFrame.from_dict(
            members_db,
            orient="index"
        )

        st.dataframe(
            df,
            use_container_width=True
        )

        st.divider()

        # ====================================================
        # DOWNLOAD EXCEL
        # ====================================================

        if os.path.exists(DB_FILE):

            with open(DB_FILE, "rb") as file:

                st.download_button(
                    label="📥 የአባላት መረጃ Excel አውርድ",
                    data=file.read(),
                    file_name="sacco_database.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    else:

        st.info(
            "📌 እስካሁን በሲስተሙ ላይ "
            "የተመዘገበ መረጃ የለም።"
        )


# ============================================================
# 6. ✏️ MEMBER EDIT / DELETE
# ============================================================

elif menu == "✏️ የአባላት መረጃ ማስተካከያ":

    st.header(
        "✏️ የአባላት መረጃ ማስተካከያ"
    )

    if members_db:

        for member_id, member in list(
            members_db.items()
        ):

            col1, col2, col3, col4, col5, col6 = st.columns(
                [1, 2.5, 2, 1.5, 1.2, 1.2]
            )

            with col1:

                st.write(
                    f"**{member_id}**"
                )

            with col2:

                st.write(
                    f"**{member.get('የአባል ስም', '')}**"
                )

            with col3:

                national_id_display = member.get(
                    "ብሔራዊ መታወቂያ (National ID)",
                    ""
                )

                if national_id_display:

                    st.write(
                        f"🪪 {national_id_display}"
                    )

                else:

                    st.write(
                        "🪪 አልተሞላም"
                    )

            with col4:

                st.write(
                    f"{float(member.get('ጠቅላላ ቁጠባ (ብር)', 0)):,.2f} ብር"
                )

            with col5:

                edit_clicked = st.button(
                    "✏️ አስተካክል",
                    key=f"edit_{member_id}",
                    use_container_width=True
                )

            with col6:

                delete_clicked = st.button(
                    "🗑️ አጥፋ",
                    key=f"delete_{member_id}",
                    use_container_width=True
                )

            # ====================================================
            # EDIT
            # ====================================================

            if edit_clicked:

                st.session_state[
                    f"editing_{member_id}"
                ] = True

            if st.session_state.get(
                f"editing_{member_id}",
                False
            ):

                with st.container(border=True):

                    st.subheader(
                        f"✏️ {member.get('የአባል ስም', '')} "
                        "- መረጃ ማስተካከያ"
                    )

                    edit_name = st.text_input(
                        "የአባል ሙሉ ስም",
                        value=member.get(
                            "የአባል ስም",
                            ""
                        ),
                        key=f"name_edit_{member_id}"
                    )

                    edit_national_id = st.text_input(
                        "🪪 ብሔራዊ መታወቂያ (National ID)",
                        value=str(
                            member.get(
                                "ብሔራዊ መታወቂያ (National ID)",
                                ""
                            )
                        ),
                        key=f"national_id_edit_{member_id}"
                    )

                    save_col, cancel_col = st.columns(2)

                    with save_col:

                        save_edit = st.button(
                            "💾 ለውጡን አስቀምጥ",
                            key=f"save_edit_{member_id}",
                            use_container_width=True
                        )

                    with cancel_col:

                        cancel_edit = st.button(
                            "❌ ሰርዝ",
                            key=f"cancel_edit_{member_id}",
                            use_container_width=True
                        )

                    if save_edit:

                        edit_name = edit_name.strip()
                        edit_national_id = (
                            edit_national_id.strip()
                        )

                        if not edit_name:

                            st.error(
                                "❌ የአባል ስም ባዶ "
                                "መሆን አይችልም!"
                            )

                        elif not edit_national_id:

                            st.error(
                                "❌ ብሔራዊ መታወቂያ "
                                "ባዶ መሆን አይችልም!"
                            )

                        else:

                            national_id_exists = False

                            for other_id, other_member in members_db.items():

                                if other_id == member_id:
                                    continue

                                other_national_id = str(
                                    other_member.get(
                                        "ብሔራዊ መታወቂያ (National ID)",
                                        ""
                                    )
                                ).strip()

                                if (
                                    other_national_id
                                    and other_national_id
                                    == edit_national_id
                                ):
                                    national_id_exists = True
                                    break

                            if national_id_exists:

                                st.error(
                                    "❌ ይህ ብሔራዊ መታወቂያ "
                                    "ለሌላ አባል ተመዝግቧል!"
                                )

                            else:

                                members_db[
                                    member_id
                                ]["የአባል ስም"] = (
                                    edit_name
                                )

                                members_db[
                                    member_id
                                ][
                                    "ብሔራዊ መታወቂያ (National ID)"
                                ] = edit_national_id

                                save_data_to_excel(
                                    members_db
                                )

                                st.session_state[
                                    f"editing_{member_id}"
                                ] = False

                                st.success(
                                    f"✅ {member_id} "
                                    "የአባሉ መረጃ "
                                    "ተስተካክሏል!"
                                )

                                st.rerun()

                    if cancel_edit:

                        st.session_state[
                            f"editing_{member_id}"
                        ] = False

                        st.rerun()

            # ====================================================
            # DELETE
            # ====================================================

            if delete_clicked:

                st.session_state[
                    f"confirm_delete_{member_id}"
                ] = True

            if st.session_state.get(
                f"confirm_delete_{member_id}",
                False
            ):

                with st.container(border=True):

                    st.warning(
                        f"⚠️ **{member.get('የአባል ስም', '')} "
                        f"({member_id})** "
                        "ለማጥፋት እርግጠኛ ነዎት?"
                    )

                    yes_col, no_col = st.columns(2)

                    with yes_col:

                        confirm_delete = st.button(
                            "🗑️ አዎ፣ አጥፋ",
                            key=f"confirm_{member_id}",
                            use_container_width=True
                        )

                    with no_col:

                        cancel_delete = st.button(
                            "❌ አይ፣ ተመለስ",
                            key=f"cancel_delete_{member_id}",
                            use_container_width=True
                        )

                    if confirm_delete:

                        del members_db[
                            member_id
                        ]

                        save_data_to_excel(
                            members_db
                        )

                        st.session_state[
                            f"confirm_delete_{member_id}"
                        ] = False

                        st.rerun()

                    if cancel_delete:

                        st.session_state[
                            f"confirm_delete_{member_id}"
                        ] = False

                        st.rerun()

            st.divider()

    else:

        st.info(
            "📌 እስካሁን የተመዘገበ "
            "አባል የለም።"
        )


# ============================================================
# 7. 📤 IMPORT EXCEL FILE
# ============================================================

elif menu == "📤 Excel ፋይል አስገባ":

    st.header(
        "📤 ከኮምፒውተር የExcel ፋይል አስገባ"
    )

    st.info(
        "📌 እዚህ ከኮምፒውተርዎ "
        "የExcel (.xlsx) ፋይል መምረጥ ይችላሉ።"
    )

    uploaded_file = st.file_uploader(
        "Excel ፋይል ይምረጡ",
        type=["xlsx"]
    )

    if uploaded_file is not None:

        try:

            imported_df = pd.read_excel(
                uploaded_file,
                dtype={
                    "መታወቂያ ቁጥር (ID)": str,
                    "ብሔራዊ መታወቂያ (National ID)": str
                }
            )

            st.success(
                f"✅ ፋይሉ በትክክል ተነቧል። "
                f"{len(imported_df)} መረጃዎች "
                "ተገኝተዋል።"
            )

            st.subheader(
                "📋 የሚገባው መረጃ"
            )

            st.dataframe(
                imported_df,
                use_container_width=True
            )

            import_button = st.button(
                "📥 ወደ ሲስተሙ አስገባ",
                use_container_width=True
            )

            if import_button:

                if (
                    "መታወቂያ ቁጥር (ID)"
                    not in imported_df.columns
                ):

                    st.error(
                        "❌ ይህ Excel ፋይል "
                        "'መታወቂያ ቁጥር (ID)' "
                        "የሚል column የለውም።"
                    )

                elif (
                    "የአባል ስም"
                    not in imported_df.columns
                ):

                    st.error(
                        "❌ ይህ Excel ፋይል "
                        "'የአባል ስም' "
                        "የሚል column የለውም።"
                    )

                else:

                    imported_df = imported_df.dropna(
                        subset=[
                            "መታወቂያ ቁጥር (ID)"
                        ]
                    )

                    imported_df[
                        "መታወቂያ ቁጥር (ID)"
                    ] = (
                        imported_df[
                            "መታወቂያ ቁጥር (ID)"
                        ]
                        .astype(str)
                        .str.strip()
                    )

                    new_count = 0
                    updated_count = 0

                    for _, row in imported_df.iterrows():

                        member_id = str(
                            row[
                                "መታወቂያ ቁጥር (ID)"
                            ]
                        ).strip()

                        member_name = str(
                            row["የአባል ስም"]
                        ).strip()

                        if (
                            not member_id
                            or member_id == "nan"
                        ):
                            continue

                        if member_id in members_db:

                            # Existing member
                            # Update supplied information
                            for column in imported_df.columns:

                                if (
                                    column
                                    != "መታወቂያ ቁጥር (ID)"
                                ):

                                    value = row[column]

                                    if pd.notna(value):

                                        members_db[
                                            member_id
                                        ][column] = value

                            updated_count += 1

                        else:

                            # New member
                            members_db[
                                member_id
                            ] = {

                                "የአባል ስም":
                                    member_name,

                                "ብሔራዊ መታወቂያ (National ID)":
                                    "",

                                "ጠቅላላ ቁጠባ (ብር)":
                                    0.0,

                                "ብድር ሁኔታ":
                                    "የለበትም",

                                "የተበደረው ጠቅላላ (ብር)":
                                    0.0,

                                "በእጅ የተሰጠ 90% (ብር)":
                                    0.0,

                                "የቀረው ዕዳ (ብር)":
                                    0.0
                            }

                            # Import additional columns
                            for column in imported_df.columns:

                                if (
                                    column
                                    != "መታወቂያ ቁጥር (ID)"
                                ):

                                    value = row[column]

                                    if pd.notna(value):

                                        members_db[
                                            member_id
                                        ][column] = value

                            new_count += 1

                    save_data_to_excel(
                        members_db
                    )

                    st.success(
                        "✅ የExcel ፋይሉ ወደ "
                        "ሲስተሙ ተገብቷል!"
                    )

                    st.info(
                        f"👤 አዲስ የገቡ አባላት፦ "
                        f"{new_count}\n\n"
                        f"🔄 የተሻሻሉ አባላት፦ "
                        f"{updated_count}\n\n"
                        f"📊 ጠቅላላ አባላት፦ "
                        f"{len(members_db)}"
                    )

                    st.rerun()

        except Exception as e:

            st.error(
                f"❌ ፋይሉን ማንበብ አልተቻለም።\n\n"
                f"ምክንያት፦ {e}"
            )

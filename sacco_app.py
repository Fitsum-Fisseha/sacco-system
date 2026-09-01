import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="የቁጠባ እና ብድር ሲስተም", layout="wide")
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
# 📋 SIDEBAR MENU
# ============================================================

menu = st.sidebar.selectbox(
    "ያሉ አማራጮች",
    [
        "👤 አባል መመዝገቢያ",
        "💰 የወር ቁጠባ ማስገቢያ",
        "💵 የብድር አገልግሎት",
        "📅 የብድር ክፍያ መመዝገቢያ",
        "📊 ጠቅላላ ሪፖርት"
        "✏️ የአባላት መረጃ ማስተካከያ"
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

    register_btn = st.button(
        "አባል መዝግብ"
    )

    if register_btn:

        if m_id and m_name:

            if m_id in members_db:

                st.error(
                    f"❌ ስህተት፦ መታወቂያ ቁጥር {m_id} "
                    "ቀደም ብሎ ተመዝግቧል!"
                )

            else:

                members_db[m_id] = {

                    "የአባል ስም": m_name,

                    "ጠቅላላ ቁጠባ (ብር)": 0.0,

                    "ብድር ሁኔታ": "የለበትም",

                    "የተበደረው ጠቅላላ (ብር)": 0.0,

                    "በእጅ የተሰጠ 90% (ብር)": 0.0,

                    "የቀረው ዕዳ (ብር)": 0.0
                }

                save_data_to_excel(members_db)

                st.success(
                    f"✅ አባል {m_name} በተሳካ ሁኔታ ተመዝግቧል!"
                )

        else:

            st.warning(
                "⚠️ እባክዎ ሁሉንም ሳጥኖች ይሙሉ!"
            )


# ============================================================
# 2. 💰 SAVINGS
# ============================================================

elif menu == "💰 የወር ቁጠባ ማስገቢያ":

    st.header("💰 የወርሃዊ ቁጠባ መመዝገቢያ")

    s_id = st.text_input(
        "የአባል መታወቂያ (ID):"
    )

    amount = st.number_input(
        "የሚቆጥበው የገንዘብ መጠን (ብር):",
        min_value=0.0,
        step=100.0
    )

    save_btn = st.button(
        "ቁጠባ መዝግብ"
    )

    if save_btn:

        if s_id in members_db:

            members_db[s_id]["ጠቅላላ ቁጠባ (ብር)"] += amount

            save_data_to_excel(members_db)

            st.success(
                f"✅ ለ{members_db[s_id]['የአባል ስም']} "
                f"{amount:,.2f} ብር ቁጠባ ተመዝግቧል።"
            )

        else:

            st.error(
                "❌ ይህ መታወቂያ በሲስተሙ ውስጥ አልተገኘም!"
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

    calculate_btn = st.button(
        "ብድር አስላ እና ፍቀድ"
    )

    if calculate_btn:

        if loan_id in members_db:

            member = members_db[loan_id]

            if member["ብድር ሁኔታ"] == "ያለበት":

                st.error(
                    f"❌ ስህተት፦ {member['የአባል ስም']} "
                    "የድሮ ብድር ስላለበት ተጨማሪ መበደር አይችልም!"
                )

            else:

                upfront_fee = loan_amount * 0.10

                net_payout = loan_amount - upfront_fee

                monthly_principal = loan_amount / 36

                monthly_interest = loan_amount * 0.02

                total_monthly = (
                    monthly_principal +
                    monthly_interest
                )

                member["ብድር ሁኔታ"] = "ያለበት"

                member["የተበደረው ጠቅላላ (ብር)"] = loan_amount

                member["በእጅ የተሰጠ 90% (ብር)"] = net_payout

                member["የቀረው ዕዳ (ብር)"] = loan_amount

                save_data_to_excel(members_db)

                st.success(
                    f"🎉 ለ{member['የአባል ስም']} ብድር ተፈቅዷል!"
                )

                st.info(
                    f"💵 በእጅ የሚሰጠው ገንዘብ (90%)፦ "
                    f"{net_payout:,.2f} ብር\n\n"
                    f"📅 የወርሃዊ ክፍያ (ለ36 ወራት)፦ "
                    f"{total_monthly:,.2f} ብር "
                    f"(ዋና፦ {monthly_principal:,.2f} + "
                    f"ወለድ 2%፦ {monthly_interest:,.2f})"
                )

        else:

            st.error(
                "❌ ይህ መታወቂያ በሲስተሙ ውስጥ አልተገኘም!"
            )


# ============================================================
# 4. 📅 LOAN REPAYMENT
# ============================================================

elif menu == "📅 የብድር ክፍያ መመዝገቢያ":

    st.header("📅 የወርሃዊ ብድር ክፍያ መቀበያ")

    p_id = st.text_input(
        "የከፋይ አባል መታወቂያ (ID):"
    )

    if p_id in members_db:

        member = members_db[p_id]

        if member["ብድር ሁኔታ"] == "ያለበት":

            st.warning(
                f"📌 {member['የአባል ስም']} "
                f"ያለበት ጠቅላላ ዕዳ፦ "
                f"{member['የቀረው ዕዳ (ብር)']:,.2f} ብር"
            )

            monthly_principal = (
                member["የተበደረው ጠቅላላ (ብር)"] / 36
            )

            monthly_interest = (
                member["የተበደረው ጠቅላላ (ብር)"] * 0.02
            )

            st.write(
                f"💡 መደበኛ የወር ክፍያ፦ "
                f"**{monthly_principal + monthly_interest:,.2f} ብር** "
                f"(ዋና፦ {monthly_principal:,.2f} + "
                f"ወለድ፦ {monthly_interest:,.2f})"
            )

            pay_amount = st.number_input(
                "አባል አሁን የከፈለው ጠቅላላ ብር "
                "(ዋና + ወለድ)፦",
                min_value=0.0
            )

            pay_btn = st.button(
                "ክፍያ መዝግብ"
            )

            if pay_btn:

                actual_principal_paid = (
                    pay_amount - monthly_interest
                )

                if actual_principal_paid < 0:
                    actual_principal_paid = 0

                member["የቀረው ዕዳ (ብር)"] -= (
                    actual_principal_paid
                )

                if member["የቀረው ዕዳ (ብር)"] <= 0:

                    member["የቀረው ዕዳ (ብር)"] = 0.0

                    member["ብድር ሁኔታ"] = "የለበትም"

                    st.success(
                        f"🎉 {member['የአባል ስም']} "
                        "ብድሩን ሙሉ በሙሉ ከፍሎ ጨርሷል!"
                    )

                else:

                    st.success(
                        f"✅ ክፍያ ተመዝግቧል። "
                        f"የቀረው ጠቅላላ ዋና ዕዳ፦ "
                        f"{member['የቀረው ዕዳ (ብር)']:,.2f} ብር"
                    )

                save_data_to_excel(members_db)

        else:

            st.info(
                f"💡 {member['የአባል ስም']} ላይ "
                "ምንም ዓይነት የብድር ዕዳ የለም።"
            )

    elif p_id:

        st.error(
            "❌ ይህ መታወቂያ በሲስተሙ ውስጥ አልተገኘም!"
        )


# ============================================================
# 5. 📊 REPORT + EDIT + DELETE
# ============================================================

elif menu == "📊 ጠቅላላ ሪፖርት":

    st.header(
        "📊 ጠቅላላ የአባላት፣ የቁጠባ እና የብድር ሪፖርት"
    )

    if members_db:

        # ----------------------------------------------------
        # REPORT TABLE
        # ----------------------------------------------------

        df = pd.DataFrame.from_dict(
            members_db,
            orient="index"
        )

        st.dataframe(
            df,
            use_container_width=True
        )

        st.divider()

        # ----------------------------------------------------
        # MEMBER MANAGEMENT
        # ----------------------------------------------------
# --- 6. የአባላት መረጃ ማስተካከያ ---
elif menu == "✏️ የአባላት መረጃ ማስተካከያ":

    st.header("✏️ የአባላት መረጃ ማስተካከያ")

    if members_db:

        for member_id, member in list(members_db.items()):

            col1, col2, col3, col4, col5, col6 = st.columns(
                [1, 2.5, 1.5, 1.5, 1.2, 1.2]
            )

            with col1:
                st.write(f"**{member_id}**")

            with col2:
                st.write(f"**{member['የአባል ስም']}**")

            with col3:
                st.write(
                    f"{member['ጠቅላላ ቁጠባ (ብር)']:,.2f} ብር"
                )

            with col4:
                st.write(member["ብድር ሁኔታ"])

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

            # -------------------------
            # EDIT
            # -------------------------
            if edit_clicked:
                st.session_state[f"editing_{member_id}"] = True

            if st.session_state.get(
                f"editing_{member_id}", False
            ):

                with st.container(border=True):

                    st.subheader(
                        f"✏️ {member['የአባል ስም']} - መረጃ ማስተካከያ"
                    )

                    edit_name = st.text_input(
                        "የአባል ሙሉ ስም",
                        value=member["የአባል ስም"],
                        key=f"name_edit_{member_id}"
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

                        if edit_name.strip():

                            members_db[member_id]["የአባል ስም"] = edit_name.strip()

                            save_data_to_excel(members_db)

                            st.session_state[
                                f"editing_{member_id}"
                            ] = False

                            st.success(
                                f"✅ {member_id} የአባሉ ስም ተስተካክሏል!"
                            )

                            st.rerun()

                        else:
                            st.error(
                                "❌ የአባል ስም ባዶ መሆን አይችልም!"
                            )

                    if cancel_edit:

                        st.session_state[
                            f"editing_{member_id}"
                        ] = False

                        st.rerun()

            # -------------------------
            # DELETE
            # -------------------------
            if delete_clicked:

                st.session_state[
                    f"confirm_delete_{member_id}"
                ] = True

            if st.session_state.get(
                f"confirm_delete_{member_id}", False
            ):

                with st.container(border=True):

                    st.warning(
                        f"⚠️ **{member['የአባል ስም']} ({member_id})** "
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

                        del members_db[member_id]

                        save_data_to_excel(members_db)

                        st.session_state[
                            f"confirm_delete_{member_id}"
                        ] = False

                        st.success(
                            f"✅ {member['የአባል ስም']} ከሲስተሙ ተሰርዟል!"
                        )

                        st.rerun()

                    if cancel_delete:

                        st.session_state[
                            f"confirm_delete_{member_id}"
                        ] = False

                        st.rerun()

            st.divider()

    else:

        st.info("📌 እስካሁን የተመዘገበ አባል የለም።")


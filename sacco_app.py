import streamlit as st
import pandas as pd
import os
import base64

# ============================================================
# PAGE SETTINGS
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

if os.path.exists(DB_FILE):

    try:

        df_load = pd.read_excel(
            DB_FILE,
            dtype={"መታወቂያ ቁጥር (ID)": str}
        )

        if "መታወቂያ ቁጥር (ID)" in df_load.columns:

            df_load["መታወቂያ ቁጥር (ID)"] = (
                df_load["መታወቂያ ቁጥር (ID)"]
                .astype(str)
                .str.strip()
            )

            members_db = (
                df_load
                .set_index("መታወቂያ ቁጥር (ID)")
                .to_dict(orient="index")
            )

        else:

            members_db = {}

    except Exception:

        members_db = {}

else:

    members_db = {}


# ============================================================
# NORMALIZE MEMBER DATA
# ============================================================

def normalize_member(member):

    member.setdefault("የአባል ስም", "")
    member.setdefault("National ID", "")
    member.setdefault("የአባል ፎቶ", "")

    member.setdefault(
        "ጠቅላላ ቁጠባ (ብር)",
        0.0
    )

    member.setdefault(
        "ብድር ሁኔታ",
        "የለበትም"
    )

    member.setdefault(
        "የተበደረው ጠቅላላ (ብር)",
        0.0
    )

    member.setdefault(
        "በእጅ የተሰጠ 90% (ብር)",
        0.0
    )

    member.setdefault(
        "የቀረው ዕዳ (ብር)",
        0.0
    )

    money_columns = [
        "ጠቅላላ ቁጠባ (ብር)",
        "የተበደረው ጠቅላላ (ብር)",
        "በእጅ የተሰጠ 90% (ብር)",
        "የቀረው ዕዳ (ብር)"
    ]

    for column in money_columns:

        try:

            value = member[column]

            if pd.isna(value):
                value = 0.0

            member[column] = float(value)

        except:

            member[column] = 0.0

    return member


for member_id in list(members_db.keys()):

    members_db[member_id] = normalize_member(
        members_db[member_id]
    )


# ============================================================
# SAVE DATABASE
# ============================================================

def save_data_to_excel(db):

    if db:

        df_save = pd.DataFrame.from_dict(
            db,
            orient="index"
        )

        df_save.index.name = (
            "መታወቂያ ቁጥር (ID)"
        )

        df_save.reset_index().to_excel(
            DB_FILE,
            index=False
        )

    else:

        if os.path.exists(DB_FILE):

            try:
                os.remove(DB_FILE)

            except:
                pass


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

    for member_id, member in members_db.items():

        if (
            exclude_member_id is not None
            and str(member_id)
            == str(exclude_member_id)
        ):
            continue

        existing_national_id = str(
            member.get(
                "National ID",
                ""
            )
        ).strip()

        if (
            existing_national_id
            and existing_national_id
            == national_id
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

    original_loan = float(
        member.get(
            "የተበደረው ጠቅላላ (ብር)",
            0
        )
    )

    outstanding = float(
        member.get(
            "የቀረው ዕዳ (ብር)",
            0
        )
    )

    monthly_interest = (
        original_loan * 0.02
    )

    principal_paid = max(
        float(payment_amount)
        - monthly_interest,
        0.0
    )

    if principal_paid >= outstanding:

        excess = (
            principal_paid
            - outstanding
        )

        member[
            "የቀረው ዕዳ (ብር)"
        ] = 0.0

        member[
            "ብድር ሁኔታ"
        ] = "የለበትም"

        # Extra money goes to savings
        member[
            "ጠቅላላ ቁጠባ (ብር)"
        ] += excess

    else:

        excess = 0.0

        member[
            "የቀረው ዕዳ (ብር)"
        ] = (
            outstanding
            - principal_paid
        )

    return (
        monthly_interest,
        principal_paid,
        excess
    )


# ============================================================
# SIDEBAR
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
# 1. MEMBER REGISTRATION
# ============================================================

if menu == "👤 አባል መመዝገቢያ":

    st.header(
        "👤 አዲስ አባል መመዝገቢያ ፎርም"
    )

    col1, col2 = st.columns(2)

    with col1:

        m_id = st.text_input(
            "የአባል መታወቂያ ቁጥር (ID):"
        ).strip()

    with col2:

        national_id = st.text_input(
            "National ID (የብሔራዊ መታወቂያ ቁጥር) *"
        ).strip()

    m_name = st.text_input(
        "የአባል ሙሉ ስም:"
    )

    photo = st.file_uploader(
        "📷 የአባል ፎቶ (አማራጭ)",
        type=[
            "jpg",
            "jpeg",
            "png"
        ]
    )

    if photo:

        st.image(
            photo,
            width=180
        )

    if st.button(
        "አባል መዝግብ"
    ):

        # Mandatory fields
        if (
            not m_id
            or not m_name.strip()
            or not national_id
        ):

            st.warning(
                "⚠️ ID፣ ሙሉ ስም እና National ID "
                "መሙላት ግዴታ ነው!"
            )

        # Duplicate member ID
        elif m_id in members_db:

            st.error(
                f"❌ መታወቂያ ቁጥር "
                f"{m_id} ቀደም ብሎ ተመዝግቧል!"
            )

        # Duplicate National ID
        elif national_id_exists(
            national_id
        ):

            st.error(
                "❌ ይህ National ID ቀደም "
                "ብሎ ተመዝግቧል። "
                "National ID ልዩ (Unique) መሆን አለበት!"
            )

        else:

            photo_data = ""

            if photo:

                photo_data = (
                    base64.b64encode(
                        photo.getvalue()
                    ).decode("utf-8")
                )

            members_db[m_id] = {

                "የአባል ስም":
                    m_name.strip(),

                "National ID":
                    national_id,

                "የአባል ፎቶ":
                    photo_data,

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

            save_data_to_excel(
                members_db
            )

            st.success(
                f"✅ {m_name.strip()} "
                "በተሳካ ሁኔታ ተመዝግቧል!"
            )


# ============================================================
# 2. SAVINGS
# ============================================================

elif menu == "💰 የወር ቁጠባ ማስገቢያ":

    st.header(
        "💰 የወርሃዊ ቁጠባ / "
        "ብድር ክፍያ መመዝገቢያ"
    )

    s_id = st.text_input(
        "የአባል መታወቂያ (ID):"
    ).strip()

    amount = st.number_input(
        "የሚከፈለው የገንዘብ መጠን (ብር):",
        min_value=0.0,
        step=100.0
    )

    if st.button(
        "ቁጠባ / ክፍያ መዝግብ"
    ):

        if s_id not in members_db:

            st.error(
                "❌ ይህ መታወቂያ "
                "በሲስተሙ ውስጥ አልተገኘም!"
            )

        elif amount <= 0:

            st.warning(
                "⚠️ የሚከፈለው መጠን "
                "ከ0 በላይ መሆን አለበት!"
            )

        else:

            member = members_db[s_id]

            # If member has loan,
            # payment goes to loan first
            if (
                member["ብድር ሁኔታ"]
                == "ያለበት"
                and
                member[
                    "የቀረው ዕዳ (ብር)"
                ] > 0
            ):

                (
                    interest,
                    principal_paid,
                    excess
                ) = apply_loan_payment(
                    member,
                    amount
                )

                save_data_to_excel(
                    members_db
                )

                if principal_paid == 0:

                    st.warning(
                        f"⚠️ {amount:,.2f} ብር "
                        f"የወለዱን {interest:,.2f} ብር "
                        "አልሞላም። "
                        "ዋና ዕዳ አልቀነሰም።"
                    )

                elif (
                    member["ብድር ሁኔታ"]
                    == "የለበትም"
                ):

                    st.success(
                        f"🎉 ብድሩ ሙሉ በሙሉ "
                        f"ተከፍሏል። "
                        f"{excess:,.2f} ብር "
                        "ትርፍ ወደ ቁጠባ "
                        "ተጨምሯል።"
                    )

                else:

                    st.success(
                        f"✅ ክፍያው ወደ ብድር "
                        "ተተግብሯል። "
                        f"የቀረው ዕዳ፦ "
                        f"{member['የቀረው ዕዳ (ብር)']:,.2f} ብር"
                    )

            # No loan → normal savings
            else:

                member[
                    "ጠቅላላ ቁጠባ (ብር)"
                ] += amount

                save_data_to_excel(
                    members_db
                )

                st.success(
                    f"✅ ለ{member['የአባል ስም']} "
                    f"{amount:,.2f} ብር "
                    "ቁጠባ ተመዝግቧል።"
                )


# ============================================================
# 3. LOAN
# ============================================================

elif menu == "💵 የብድር አገልግሎት":

    st.header(
        "💵 የብድር ማመልከቻ እና ስሌት"
    )

    loan_id = st.text_input(
        "የተበዳሪው አባል መታወቂያ (ID):"
    ).strip()

    loan_amount = st.number_input(
        "የሚጠይቀው የብድር መጠን (ብር):",
        min_value=0.0,
        step=1000.0
    )

    if loan_id in members_db:

        member = members_db[loan_id]

        current_savings = float(
            member[
                "ጠቅላላ ቁጠባ (ብር)"
            ]
        )

        max_loan = (
            current_savings * 4
        )

        st.info(
            f"💰 የአሁኑ ቁጠባ፦ "
            f"{current_savings:,.2f} ብር\n\n"
            f"📌 ከፍተኛ የሚፈቀደው "
            f"ብድር (4×)፦ "
            f"{max_loan:,.2f} ብር"
        )

    if st.button(
        "ብድር አስላ እና ፍቀድ"
    ):

        if loan_id not in members_db:

            st.error(
                "❌ ይህ መታወቂያ "
                "በሲስተሙ ውስጥ አልተገኘም!"
            )

        elif loan_amount <= 0:

            st.warning(
                "⚠️ የብድር መጠን "
                "ከ0 በላይ መሆን አለበት!"
            )

        else:

            member = members_db[loan_id]

            current_savings = float(
                member[
                    "ጠቅላላ ቁጠባ (ብር)"
                ]
            )

            max_loan = (
                current_savings * 4
            )

            # Existing loan
            if (
                member["ብድር ሁኔታ"]
                == "ያለበት"
            ):

                st.error(
                    f"❌ {member['የአባል ስም']} "
                    "የድሮ ብድር ስላለበት "
                    "ተጨማሪ መበደር አይችልም!"
                )

            # More than 4x savings
            elif loan_amount > max_loan:

                st.error(
                    f"❌ የብድር ጥያቄው "
                    "ከሚፈቀደው 4× የቁጠባ "
                    "መጠን በላይ ነው!\n\n"
                    f"💰 ቁጠባ፦ "
                    f"{current_savings:,.2f} ብር\n"
                    f"📌 ከፍተኛ ብድር፦ "
                    f"{max_loan:,.2f} ብር"
                )

            else:

                upfront_fee = (
                    loan_amount * 0.10
                )

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

                save_data_to_excel(
                    members_db
                )

                st.success(
                    f"🎉 ለ{member['የአባል ስም']} "
                    "ብድር ተፈቅዷል!"
                )

                st.info(
                    f"💵 በእጅ የሚሰጠው 90%፦ "
                    f"{net_payout:,.2f} ብር\n\n"
                    f"📅 የወርሃዊ ክፍያ "
                    f"(36 ወራት)፦ "
                    f"{total_monthly:,.2f} ብር\n\n"
                    f"(ዋና፦ "
                    f"{monthly_principal:,.2f} + "
                    f"ወለድ 2%፦ "
                    f"{monthly_interest:,.2f})"
                )


# ============================================================
# 4. LOAN REPAYMENT
# ============================================================

elif menu == "📅 የብድር ክፍያ መመዝገቢያ":

    st.header(
        "📅 የወርሃዊ ብድር ክፍያ መቀበያ"
    )

    p_id = st.text_input(
        "የከፋይ አባል መታወቂያ (ID):"
    ).strip()

    if p_id in members_db:

        member = members_db[p_id]

        if (
            member["ብድር ሁኔታ"]
            == "ያለበት"
            and
            member[
                "የቀረው ዕዳ (ብር)"
            ] > 0
        ):

            outstanding = float(
                member[
                    "የቀረው ዕዳ (ብር)"
                ]
            )

            original_loan = float(
                member[
                    "የተበደረው ጠቅላላ (ብር)"
                ]
            )

            monthly_principal = (
                original_loan / 36
            )

            monthly_interest = (
                original_loan * 0.02
            )

            st.warning(
                f"📌 {member['የአባል ስም']} "
                f"ያለበት ዕዳ፦ "
                f"{outstanding:,.2f} ብር"
            )

            st.write(
                f"💡 መደበኛ የወር ክፍያ፦ "
                f"**{monthly_principal + monthly_interest:,.2f} ብር** "
                f"(ዋና፦ "
                f"{monthly_principal:,.2f} + "
                f"ወለድ፦ "
                f"{monthly_interest:,.2f})"
            )

            pay_amount = st.number_input(
                "አባል አሁን የከፈለው "
                "ጠቅላላ ብር (ዋና + ወለድ)፦",
                min_value=0.0
            )

            if st.button(
                "ክፍያ መዝግብ"
            ):

                if pay_amount <= 0:

                    st.warning(
                        "⚠️ የክፍያ መጠን "
                        "ከ0 በላይ መሆን አለበት!"
                    )

                else:

                    (
                        interest,
                        principal_paid,
                        excess
                    ) = apply_loan_payment(
                        member,
                        pay_amount
                    )

                    save_data_to_excel(
                        members_db
                    )

                    if principal_paid == 0:

                        st.warning(
                            f"⚠️ ክፍያው "
                            f"የወለዱን "
                            f"{interest:,.2f} ብር "
                            "አልሞላም። "
                            "ዋና ዕዳ አልቀነሰም።"
                        )

                    elif (
                        member["ብድር ሁኔታ"]
                        == "የለበትም"
                    ):

                        st.success(
                            f"🎉 ብድሩ ተጠናቋል።\n\n"
                            f"💰 {excess:,.2f} ብር "
                            "ትርፍ ወደ ቁጠባ "
                            "ተጨምሯል።"
                        )

                    else:

                        st.success(
                            f"✅ ክፍያ ተመዝግቧል።\n\n"
                            f"የቀረው ዕዳ፦ "
                            f"{member['የቀረው ዕዳ (ብር)']:,.2f} ብር"
                        )

        else:

            st.info(
                f"💡 {member['የአባል ስም']} "
                "ላይ ምንም የብድር "
                "ዕዳ የለም።"
            )

    elif p_id:

        st.error(
            "❌ ይህ መታወቂያ "
            "በሲስተሙ ውስጥ አልተገኘም!"
        )


# ============================================================
# 5. REPORT
# ============================================================

elif menu == "📊 ጠቅላላ ሪፖርት":

    st.header(
        "📊 ጠቅላላ የአባላት፣ "
        "የቁጠባ እና የብድር ሪፖርት"
    )

    if members_db:

        rows = []

        for member_id, member in members_db.items():

            rows.append({

                "ID":
                    member_id,

                "National ID":
                    member.get(
                        "National ID",
                        ""
                    ),

                "የአባል ስም":
                    member.get(
                        "የአባል ስም",
                        ""
                    ),

                "ጠቅላላ ቁጠባ (ብር)":
                    member.get(
                        "ጠቅላላ ቁጠባ (ብር)",
                        0.0
                    ),

                "ብድር ሁኔታ":
                    member.get(
                        "ብድር ሁኔታ",
                        "የለበትም"
                    ),

                "የተበደረው ጠቅላላ (ብር)":
                    member.get(
                        "የተበደረው ጠቅላላ (ብር)",
                        0.0
                    ),

                "የቀረው ዕዳ (ብር)":
                    member.get(
                        "የቀረው ዕዳ (ብር)",
                        0.0
                    )
            })

        report_df = pd.DataFrame(
            rows
        )

        st.dataframe(
            report_df,
            use_container_width=True,
            hide_index=True
        )

        st.divider()

        if os.path.exists(DB_FILE):

            with open(
                DB_FILE,
                "rb"
            ) as file:

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
# 6. MEMBER EDIT / DELETE
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
                [1, 2.3, 1.5, 1.5, 1.1, 1.1]
            )

            with col1:

                st.write(
                    f"**{member_id}**"
                )

            with col2:

                st.write(
                    f"**{member['የአባል ስም']}**"
                )

            with col3:

                st.write(
                    f"{member['ጠቅላላ ቁጠባ (ብር)']:,.2f} ብር"
                )

            with col4:

                st.write(
                    member["ብድር ሁኔታ"]
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

            if edit_clicked:

                st.session_state[
                    f"editing_{member_id}"
                ] = True

            if delete_clicked:

                st.session_state[
                    f"confirm_delete_{member_id}"
                ] = True

            # ------------------------------------------------
            # EDIT
            # ------------------------------------------------

            if st.session_state.get(
                f"editing_{member_id}",
                False
            ):

                with st.container(
                    border=True
                ):

                    st.subheader(
                        f"✏️ {member['የአባል ስም']} "
                        "- መረጃ ማስተካከያ"
                    )

                    edit_name = st.text_input(
                        "የአባል ሙሉ ስም",
                        value=member[
                            "የአባል ስም"
                        ],
                        key=f"name_edit_{member_id}"
                    )

                    edit_national_id = st.text_input(
                        "National ID *",
                        value=str(
                            member.get(
                                "National ID",
                                ""
                            )
                        ),
                        key=f"nid_edit_{member_id}"
                    )

                    edit_photo = st.file_uploader(
                        "📷 አዲስ ፎቶ (አማራጭ)",
                        type=[
                            "jpg",
                            "jpeg",
                            "png"
                        ],
                        key=f"photo_edit_{member_id}"
                    )

                    if member.get(
                        "የአባል ፎቶ"
                    ):

                        try:

                            st.image(
                                base64.b64decode(
                                    member[
                                        "የአባል ፎቶ"
                                    ]
                                ),
                                width=120
                            )

                        except:

                            pass

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

                        if (
                            not edit_name.strip()
                            or not edit_national_id.strip()
                        ):

                            st.error(
                                "❌ ሙሉ ስም እና "
                                "National ID ባዶ "
                                "መሆን አይችሉም!"
                            )

                        elif national_id_exists(
                            edit_national_id.strip(),
                            exclude_member_id=member_id
                        ):

                            st.error(
                                "❌ ይህ National ID "
                                "ቀደም ብሎ ለሌላ "
                                "አባል ተመዝግቧል!"
                            )

                        else:

                            members_db[
                                member_id
                            ][
                                "የአባል ስም"
                            ] = edit_name.strip()

                            members_db[
                                member_id
                            ][
                                "National ID"
                            ] = edit_national_id.strip()

                            if edit_photo:

                                members_db[
                                    member_id
                                ][
                                    "የአባል ፎቶ"
                                ] = (
                                    base64.b64encode(
                                        edit_photo.getvalue()
                                    ).decode(
                                        "utf-8"
                                    )
                                )

                            save_data_to_excel(
                                members_db
                            )

                            st.session_state[
                                f"editing_{member_id}"
                            ] = False

                            st.success(
                                "✅ የአባሉ መረጃ "
                                "ተስተካክሏል!"
                            )

                            st.rerun()

                    if cancel_edit:

                        st.session_state[
                            f"editing_{member_id}"
                        ] = False

                        st.rerun()

            # ------------------------------------------------
            # DELETE
            # ------------------------------------------------

            if st.session_state.get(
                f"confirm_delete_{member_id}",
                False
            ):

                with st.container(
                    border=True
                ):

                    st.warning(
                        f"⚠️ **{member['የአባል ስም']} "
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
# 7. IMPORT EXCEL
# ============================================================

elif menu == "📤 Excel ፋይል አስገባ":

    st.header(
        "📤 ከኮምፒውተር "
        "የExcel ፋይል አስገባ"
    )

    st.info(
        "📌 Excel ፋይሉ የሚከተሉትን "
        "columns ሊኖሩት ይገባል፦\n\n"
        "• መታወቂያ ቁጥር (ID)\n"
        "• የአባል ስም\n"
        "• National ID\n\n"
        "National ID ልዩ (Unique) መሆን አለበት።"
    )

    uploaded_file = st.file_uploader(
        "Excel ፋይል ይምረጡ",
        type=["xlsx"]
    )

    if uploaded_file is not None:

        try:

            imported_df = pd.read_excel(
                uploaded_file,
                dtype=str
            )

            st.subheader(
                "📋 የሚገባው መረጃ"
            )

            st.dataframe(
                imported_df,
                use_container_width=True
            )

            if st.button(
                "📥 ወደ ሲስተሙ አስገባ",
                use_container_width=True
            ):

                required_columns = [
                    "መታወቂያ ቁጥር (ID)",
                    "የአባል ስም",
                    "National ID"
                ]

                missing_columns = [
                    column
                    for column in required_columns
                    if column
                    not in imported_df.columns
                ]

                if missing_columns:

                    st.error(
                        "❌ የሚከተሉት "
                        "አስፈላጊ columns "
                        "የሉም፦ "
                        + ", ".join(
                            missing_columns
                        )
                    )

                else:

                    imported_df = (
                        imported_df
                        .dropna(
                            subset=required_columns
                        )
                    )

                    for column in required_columns:

                        imported_df[
                            column
                        ] = (
                            imported_df[column]
                            .astype(str)
                            .str.strip()
                        )

                    seen_nids = set()

                    errors = []

                    for _, row in imported_df.iterrows():

                        member_id = row[
                            "መታወቂያ ቁጥር (ID)"
                        ]

                        nid = row[
                            "National ID"
                        ]

                        if (
                            member_id
                            in members_db
                        ):

                            errors.append(
                                f"ID {member_id} "
                                "ቀድሞ አለ"
                            )

                        if nid in seen_nids:

                            errors.append(
                                f"National ID "
                                f"{nid} "
                                "በExcel ውስጥ "
                                "ተደግሟል"
                            )

                        if national_id_exists(
                            nid
                        ):

                            errors.append(
                                f"National ID "
                                f"{nid} "
                                "ቀድሞ አለ"
                            )

                        seen_nids.add(
                            nid
                        )

                    if errors:

                        st.error(
                            "❌ ፋይሉ አልገባም።\n\n"
                            + "\n".join(
                                errors[:20]
                            )
                        )

                    else:

                        for _, row in imported_df.iterrows():

                            member_id = row[
                                "መታወቂያ ቁጥር (ID)"
                            ]

                            members_db[
                                member_id
                            ] = {

                                "የአባል ስም":
                                    row[
                                        "የአባል ስም"
                                    ],

                                "National ID":
                                    row[
                                        "National ID"
                                    ],

                                "የአባል ፎቶ":
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

                            # Import other columns
                            for column in imported_df.columns:

                                if column not in required_columns:

                                    value = row[
                                        column
                                    ]

                                    if (
                                        pd.notna(value)
                                        and str(value).strip()
                                    ):

                                        members_db[
                                            member_id
                                        ][column] = value

                        save_data_to_excel(
                            members_db
                        )

                        st.success(
                            f"✅ {len(imported_df)} "
                            "አባላት በተሳካ "
                            "ሁኔታ ገብተዋል!"
                        )

                        st.rerun()

        except Exception as e:

            st.error(
                "❌ ፋይሉን ማንበብ "
                "አልተቻለም።\n\n"
                f"ምክንያት፦ {e}"
            )

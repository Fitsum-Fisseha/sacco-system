```python
import streamlit as st
import pandas as pd
import os
from datetime import datetime
import uuid

st.set_page_config(
    page_title="የቁጠባ እና ብድር ሲስተም",
    page_icon="🏦",
    layout="wide"
)

st.title("🏦 ተስፋ የገንዘብ ቁጠባና ብድር ማህበር")

DB_FILE = "sacco_database.xlsx"

# =========================================================
# COLUMNS
# =========================================================

savings_columns = [
    "Transaction ID",
    "Date",
    "Member ID",
    "Member Name",
    "Amount (ETB)",
    "Type",
    "Description"
]

loan_columns = [
    "Loan ID",
    "Date",
    "Member ID",
    "Member Name",
    "Loan Amount (ETB)",
    "Upfront Fee (10%)",
    "Net Payout (90%)",
    "Monthly Principal",
    "Monthly Interest",
    "Monthly Payment",
    "Loan Term (Months)",
    "Total Expected Payment",
    "Status"
]

repayment_columns = [
    "Payment ID",
    "Date",
    "Member ID",
    "Member Name",
    "Loan ID",
    "Payment Amount (ETB)",
    "Interest Paid",
    "Principal Paid",
    "Remaining Principal",
    "Description"
]


# =========================================================
# LOAD MEMBERS
# =========================================================

members_db = {}

if os.path.exists(DB_FILE):

    try:
        df_load = pd.read_excel(
            DB_FILE,
            sheet_name="Members",
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

    except Exception:

        # Compatibility with the original single-sheet database
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

        except Exception:
            members_db = {}


# =========================================================
# LOAD SHEETS
# =========================================================

def load_sheet(sheet_name, columns):

    if not os.path.exists(DB_FILE):
        return pd.DataFrame(columns=columns)

    try:

        df = pd.read_excel(
            DB_FILE,
            sheet_name=sheet_name
        )

        for column in columns:

            if column not in df.columns:
                df[column] = ""

        return df[columns]

    except Exception:

        return pd.DataFrame(columns=columns)


savings_db = load_sheet(
    "Savings",
    savings_columns
)

loans_db = load_sheet(
    "Loans",
    loan_columns
)

repayments_db = load_sheet(
    "Repayments",
    repayment_columns
)


# =========================================================
# SAVE ALL DATA
# =========================================================

def save_all_data():

    members_df = pd.DataFrame.from_dict(
        members_db,
        orient="index"
    )

    members_df.index.name = "መታወቂያ ቁጥር (ID)"

    members_df = members_df.reset_index()

    try:

        with pd.ExcelWriter(
            DB_FILE,
            engine="openpyxl",
            mode="w"
        ) as writer:

            members_df.to_excel(
                writer,
                sheet_name="Members",
                index=False
            )

            savings_db.to_excel(
                writer,
                sheet_name="Savings",
                index=False
            )

            loans_db.to_excel(
                writer,
                sheet_name="Loans",
                index=False
            )

            repayments_db.to_excel(
                writer,
                sheet_name="Repayments",
                index=False
            )

    except PermissionError:

        st.error(
            "❌ sacco_database.xlsx ፋይሉ ክፍት ነው። "
            "እባክዎ Excel ፋይሉን ዝጉና እንደገና ይሞክሩ።"
        )


# =========================================================
# HELPER FUNCTIONS
# =========================================================

def generate_id(prefix):

    return (
        f"{prefix}-"
        f"{datetime.now().strftime('%Y%m%d%H%M%S')}-"
        f"{uuid.uuid4().hex[:4].upper()}"
    )


def get_member(member_id):

    member_id = str(member_id).strip()

    return members_db.get(member_id)


def money(value):

    try:
        return f"{float(value):,.2f}"
    except Exception:
        return "0.00"


# =========================================================
# SIDEBAR
# =========================================================

menu = st.sidebar.selectbox(
    "ያሉ አማራጮች",
    [
        "🏠 Dashboard",
        "👤 አባል መመዝገቢያ",
        "🔎 የአባል Profile / Statement",
        "💰 የወር ቁጠባ ማስገቢያ",
        "📜 የቁጠባ ታሪክ",
        "💵 የብድር አገልግሎት",
        "📋 የብድር ታሪክ",
        "📅 የብድር ክፍያ መመዝገቢያ",
        "🧾 የክፍያ ታሪክ",
        "📊 ጠቅላላ ሪፖርት"
    ]
)


# =========================================================
# 1. DASHBOARD
# =========================================================

if menu == "🏠 Dashboard":

    st.header("🏠 Dashboard")

    total_members = len(members_db)

    total_savings = sum(
        float(
            member.get(
                "ጠቅላላ ቁጠባ (ብር)",
                0
            ) or 0
        )
        for member in members_db.values()
    )

    total_loans = sum(
        float(
            member.get(
                "የተበደረው ጠቅላላ (ብር)",
                0
            ) or 0
        )
        for member in members_db.values()
    )

    total_debt = sum(
        float(
            member.get(
                "የቀረው ዕዳ (ብር)",
                0
            ) or 0
        )
        for member in members_db.values()
    )

    active_loans = sum(
        1
        for member in members_db.values()
        if member.get("ብድር ሁኔታ") == "ያለበት"
    )

    c1, c2, c3, c4, c5 = st.columns(5)

    c1.metric(
        "👥 አባላት",
        total_members
    )

    c2.metric(
        "💰 ጠቅላላ ቁጠባ",
        f"{total_savings:,.2f} ETB"
    )

    c3.metric(
        "💵 ጠቅላላ ብድር",
        f"{total_loans:,.2f} ETB"
    )

    c4.metric(
        "📌 የቀረ ዕዳ",
        f"{total_debt:,.2f} ETB"
    )

    c5.metric(
        "📋 ንቁ ብድሮች",
        active_loans
    )

    st.divider()

    if members_db:

        dashboard_data = []

        for member_id, member in members_db.items():

            dashboard_data.append({

                "Member ID": member_id,

                "የአባል ስም":
                    member.get(
                        "የአባል ስም",
                        ""
                    ),

                "ቁጠባ":
                    float(
                        member.get(
                            "ጠቅላላ ቁጠባ (ብር)",
                            0
                        ) or 0
                    ),

                "ብድር":
                    float(
                        member.get(
                            "የተበደረው ጠቅላላ (ብር)",
                            0
                        ) or 0
                    ),

                "የቀረ ዕዳ":
                    float(
                        member.get(
                            "የቀረው ዕዳ (ብር)",
                            0
                        ) or 0
                    ),

                "ብድር ሁኔታ":
                    member.get(
                        "ብድር ሁኔታ",
                        "የለበትም"
                    )
            })

        st.dataframe(
            pd.DataFrame(dashboard_data),
            use_container_width=True,
            hide_index=True
        )


# =========================================================
# 2. MEMBER REGISTRATION
# =========================================================

elif menu == "👤 አባል መመዝገቢያ":

    st.header("👤 አዲስ አባል መመዝገቢያ")

    c1, c2 = st.columns(2)

    with c1:

        m_id = st.text_input(
            "የአባል መታወቂያ ቁጥር (ID):"
        ).strip()

    with c2:

        m_name = st.text_input(
            "የአባል ሙሉ ስም:"
        ).strip()

    register_btn = st.button(
        "➕ አባል መዝግብ",
        type="primary"
    )

    if register_btn:

        if not m_id or not m_name:

            st.warning(
                "⚠️ እባክዎ ሁሉንም መረጃ ይሙሉ!"
            )

        elif m_id in members_db:

            st.error(
                f"❌ ID {m_id} ቀደም ብሎ ተመዝግቧል!"
            )

        else:

            members_db[m_id] = {

                "የአባል ስም": m_name,

                "ጠቅላላ ቁጠባ (ብር)": 0.0,

                "ብድር ሁኔታ": "የለበትም",

                "የተበደረው ጠቅላላ (ብር)": 0.0,

                "በእጅ የተሰጠ 90% (ብር)": 0.0,

                "የቀረው ዕዳ (ብር)": 0.0,

                "የመጨረሻ ቁጠባ ቀን": "",

                "የመጨረሻ ክፍያ ቀን": ""
            }

            save_all_data()

            st.success(
                f"✅ {m_name} በተሳካ ሁኔታ ተመዝግቧል!"
            )


# =========================================================
# 3. MEMBER PROFILE / STATEMENT
# =========================================================

elif menu == "🔎 የአባል Profile / Statement":

    st.header("🔎 የአባል Profile / Statement")

    profile_id = st.text_input(
        "የአባል Member ID:"
    ).strip()

    member = get_member(profile_id)

    if member:

        member_name = member.get(
            "የአባል ስም",
            ""
        )

        total_saving = float(
            member.get(
                "ጠቅላላ ቁጠባ (ብር)",
                0
            ) or 0
        )

        total_loan = float(
            member.get(
                "የተበደረው ጠቅላላ (ብር)",
                0
            ) or 0
        )

        remaining_debt = float(
            member.get(
                "የቀረው ዕዳ (ብር)",
                0
            ) or 0
        )

        loan_status = member.get(
            "ብድር ሁኔታ",
            "የለበትም"
        )

        st.subheader(
            f"👤 {member_name}"
        )

        st.write(
            f"🆔 **Member ID:** {profile_id}"
        )

        st.divider()

        c1, c2, c3, c4 = st.columns(4)

        c1.metric(
            "💰 ጠቅላላ ቁጠባ",
            f"{total_saving:,.2f} ETB"
        )

        c2.metric(
            "💵 ጠቅላላ ብድር",
            f"{total_loan:,.2f} ETB"
        )

        c3.metric(
            "📌 የቀረ ዕዳ",
            f"{remaining_debt:,.2f} ETB"
        )

        c4.metric(
            "📋 የብድር ሁኔታ",
            loan_status
        )

        st.divider()

        c1, c2 = st.columns(2)

        with c1:

            st.write(
                "📅 **የመጨረሻ ቁጠባ:**",
                member.get(
                    "የመጨረሻ ቁጠባ ቀን",
                    "—"
                )
            )

        with c2:

            st.write(
                "📅 **የመጨረሻ ክፍያ:**",
                member.get(
                    "የመጨረሻ ክፍያ ቀን",
                    "—"
                )
            )

        # -------------------------------------------------
        # SAVINGS STATEMENT
        # -------------------------------------------------

        st.divider()

        st.subheader("💰 Savings Statement")

        member_savings = savings_db[
            savings_db["Member ID"].astype(str)
            == profile_id
        ].copy()

        if not member_savings.empty:

            member_savings = member_savings.sort_values(
                by="Date",
                ascending=False
            )

            st.dataframe(
                member_savings,
                use_container_width=True,
                hide_index=True
            )

            saving_total = pd.to_numeric(
                member_savings["Amount (ETB)"],
                errors="coerce"
            ).fillna(0).sum()

            st.metric(
                "የዚህ አባል የተመዘገበ ጠቅላላ ቁጠባ",
                f"{saving_total:,.2f} ETB"
            )

        else:

            st.info(
                "📌 ይህ አባል እስካሁን የቁጠባ transaction የለውም።"
            )

        # -------------------------------------------------
        # LOAN STATEMENT
        # -------------------------------------------------

        st.divider()

        st.subheader("💵 Loan Statement")

        member_loans = loans_db[
            loans_db["Member ID"].astype(str)
            == profile_id
        ].copy()

        if not member_loans.empty:

            member_loans = member_loans.sort_values(
                by="Date",
                ascending=False
            )

            st.dataframe(
                member_loans,
                use_container_width=True,
                hide_index=True
            )

        else:

            st.info(
                "📌 ይህ አባል የብድር ታሪክ የለውም።"
            )

        # -------------------------------------------------
        # REPAYMENT STATEMENT
        # -------------------------------------------------

        st.divider()

        st.subheader("🧾 Repayment Statement")

        member_repayments = repayments_db[
            repayments_db["Member ID"].astype(str)
            == profile_id
        ].copy()

        if not member_repayments.empty:

            member_repayments = member_repayments.sort_values(
                by="Date",
                ascending=False
            )

            st.dataframe(
                member_repayments,
                use_container_width=True,
                hide_index=True
            )

            total_paid = pd.to_numeric(
                member_repayments["Payment Amount (ETB)"],
                errors="coerce"
            ).fillna(0).sum()

            total_principal = pd.to_numeric(
                member_repayments["Principal Paid"],
                errors="coerce"
            ).fillna(0).sum()

            total_interest = pd.to_numeric(
                member_repayments["Interest Paid"],
                errors="coerce"
            ).fillna(0).sum()

            c1, c2, c3 = st.columns(3)

            c1.metric(
                "💳 ጠቅላላ የተከፈለ",
                f"{total_paid:,.2f} ETB"
            )

            c2.metric(
                "🏦 Principal",
                f"{total_principal:,.2f} ETB"
            )

            c3.metric(
                "📈 Interest",
                f"{total_interest:,.2f} ETB"
            )

        else:

            st.info(
                "📌 ይህ አባል እስካሁን የብድር ክፍያ አልፈጸመም።"
            )

        # -------------------------------------------------
        # MEMBER STATEMENT DOWNLOAD
        # -------------------------------------------------

        st.divider()

        statement_text = f"""
ተስፋ የገንዘብ ቁጠባና ብድር ማህበር
የአባል Statement

Member ID: {profile_id}
Member Name: {member_name}

ጠቅላላ ቁጠባ: {total_saving:,.2f} ETB
ጠቅላላ ብድር: {total_loan:,.2f} ETB
የቀረ ዕዳ: {remaining_debt:,.2f} ETB
የብድር ሁኔታ: {loan_status}

የመጨረሻ ቁጠባ:
{member.get("የመጨረሻ ቁጠባ ቀን", "")}

የመጨረሻ ክፍያ:
{member.get("የመጨረሻ ክፍያ ቀን", "")}
"""

        st.download_button(
            "📥 Member Statement አውርድ",
            data=statement_text.encode("utf-8"),
            file_name=f"member_statement_{profile_id}.txt",
            mime="text/plain"
        )

    elif profile_id:

        st.error(
            "❌ ይህ Member ID በሲስተሙ ውስጥ አልተገኘም!"
        )


# =========================================================
# 4. SAVINGS
# =========================================================

elif menu == "💰 የወር ቁጠባ ማስገቢያ":

    st.header("💰 የወርሃዊ ቁጠባ መመዝገቢያ")

    s_id = st.text_input(
        "የአባል መታወቂያ (ID):"
    ).strip()

    member = get_member(s_id)

    if member:

        st.info(
            f"👤 አባል፦ **{member['የአባል ስም']}**"
        )

        st.write(
            f"💰 አሁን ያለው ቁጠባ፦ "
            f"**{money(member.get('ጠቅላላ ቁጠባ (ብር)', 0))} ETB**"
        )

        amount = st.number_input(
            "የሚቆጥበው የገንዘብ መጠን (ብር):",
            min_value=0.01,
            step=100.0
        )

        description = st.text_input(
            "ማብራሪያ:"
        )

        save_btn = st.button(
            "💰 ቁጠባ መዝግብ",
            type="primary"
        )

        if save_btn:

            transaction_id = generate_id("SAV")

            today = datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )

            member["ጠቅላላ ቁጠባ (ብር)"] = (
                float(
                    member.get(
                        "ጠቅላላ ቁጠባ (ብር)",
                        0
                    ) or 0
                )
                + amount
            )

            member["የመጨረሻ ቁጠባ ቀን"] = today

            new_saving = pd.DataFrame([{

                "Transaction ID": transaction_id,
                "Date": today,
                "Member ID": s_id,
                "Member Name": member["የአባል ስም"],
                "Amount (ETB)": amount,
                "Type": "Deposit",
                "Description": description

            }])

            savings_db = pd.concat(
                [savings_db, new_saving],
                ignore_index=True
            )

            save_all_data()

            st.success(
                f"✅ {amount:,.2f} ETB ቁጠባ ተመዝግቧል።"
            )

    elif s_id:

        st.error(
            "❌ ይህ Member ID በሲስተሙ ውስጥ አልተገኘም!"
        )


# =========================================================
# 5. SAVINGS HISTORY
# =========================================================

elif menu == "📜 የቁጠባ ታሪክ":

    st.header("📜 የቁጠባ ግብይት ታሪክ")

    if not savings_db.empty:

        search_id = st.text_input(
            "🔎 Member ID በመጠቀም ፈልግ:"
        ).strip()

        display_df = savings_db.copy()

        if search_id:

            display_df = display_df[
                display_df["Member ID"].astype(str)
                == search_id
            ]

        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True
        )

        total = pd.to_numeric(
            display_df["Amount (ETB)"],
            errors="coerce"
        ).fillna(0).sum()

        st.metric(
            "የተመረጠው ጠቅላላ ቁጠባ",
            f"{total:,.2f} ETB"
        )

    else:

        st.info(
            "📌 እስካሁን የቁጠባ transaction የለም።"
        )


# =========================================================
# 6. LOAN SERVICE
# =========================================================

elif menu == "💵 የብድር አገልግሎት":

    st.header("💵 የብድር ማመልከቻ እና ስሌት")

    loan_id = st.text_input(
        "የተበዳሪው አባል ID:"
    ).strip()

    member = get_member(loan_id)

    if member:

        st.info(
            f"👤 አባል፦ **{member['የአባል ስም']}**"
        )

        if member["ብድር ሁኔታ"] == "ያለበት":

            st.error(
                f"❌ {member['የአባል ስም']} "
                "አሁን ያለ ብድር ስላለበት አዲስ ብድር መውሰድ አይችልም።"
            )

        else:

            loan_amount = st.number_input(
                "የሚጠይቀው የብድር መጠን (ብር):",
                min_value=100.0,
                step=1000.0
            )

            loan_term = st.number_input(
                "የብድር ጊዜ (ወር):",
                min_value=1,
                max_value=120,
                value=36,
                step=1
            )

            upfront_fee = loan_amount * 0.10
            net_payout = loan_amount - upfront_fee
            monthly_principal = loan_amount / loan_term
            monthly_interest = loan_amount * 0.02
            monthly_payment = (
                monthly_principal
                + monthly_interest
            )

            total_expected_payment = (
                monthly_payment * loan_term
            )

            c1, c2, c3, c4 = st.columns(4)

            c1.metric(
                "የብድር መጠን",
                f"{loan_amount:,.2f}"
            )

            c2.metric(
                "10% ቅናሽ",
                f"{upfront_fee:,.2f}"
            )

            c3.metric(
                "90% በእጅ",
                f"{net_payout:,.2f}"
            )

            c4.metric(
                "ወርሃዊ ክፍያ",
                f"{monthly_payment:,.2f}"
            )

            st.write(
                f"ዋና፦ {monthly_principal:,.2f} ETB | "
                f"ወለድ፦ {monthly_interest:,.2f} ETB"
            )

            approve_btn = st.button(
                "💵 ብድር አስላ እና ፍቀድ",
                type="primary"
            )

            if approve_btn:

                loan_record_id = generate_id("LOAN")

                today = datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )

                member["ብድር ሁኔታ"] = "ያለበት"

                member["የተበደረው ጠቅላላ (ብር)"] = loan_amount

                member["በእጅ የተሰጠ 90% (ብር)"] = net_payout

                member["የቀረው ዕዳ (ብር)"] = loan_amount

                new_loan = pd.DataFrame([{

                    "Loan ID": loan_record_id,
                    "Date": today,
                    "Member ID": loan_id,
                    "Member Name": member["የአባል ስም"],
                    "Loan Amount (ETB)": loan_amount,
                    "Upfront Fee (10%)": upfront_fee,
                    "Net Payout (90%)": net_payout,
                    "Monthly Principal": monthly_principal,
                    "Monthly Interest": monthly_interest,
                    "Monthly Payment": monthly_payment,
                    "Loan Term (Months)": loan_term,
                    "Total Expected Payment": total_expected_payment,
                    "Status": "Active"
                }])

                loans_db = pd.concat(
                    [loans_db, new_loan],
                    ignore_index=True
                )

                save_all_data()

                st.success(
                    f"🎉 {member['የአባል ስም']} ብድር ተፈቅዷል!"
                )

    elif loan_id:

        st.error(
            "❌ ይህ Member ID አልተገኘም!"
        )


# =========================================================
# 7. LOAN HISTORY
# =========================================================

elif menu == "📋 የብድር ታሪክ":

    st.header("📋 የብድር ታሪክ")

    if not loans_db.empty:

        search_id = st.text_input(
            "🔎 Member ID ፈልግ:"
        ).strip()

        display_df = loans_db.copy()

        if search_id:

            display_df = display_df[
                display_df["Member ID"].astype(str)
                == search_id
            ]

        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True
        )

    else:

        st.info(
            "📌 እስካሁን የብድር ታሪክ የለም።"
        )


# =========================================================
# 8. REPAYMENT
# =========================================================

elif menu == "📅 የብድር ክፍያ መመዝገቢያ":

    st.header("📅 የወርሃዊ ብድር ክፍያ መቀበያ")

    p_id = st.text_input(
        "የከፋይ አባል ID:"
    ).strip()

    member = get_member(p_id)

    if member:

        if member["ብድር ሁኔታ"] == "ያለበት":

            loan_amount = float(
                member["የተበደረው ጠቅላላ (ብር)"]
            )

            remaining_debt = float(
                member["የቀረው ዕዳ (ብር)"]
            )

            member_loans = loans_db[
                loans_db["Member ID"].astype(str)
                == p_id
            ]

            active_loan = None

            if not member_loans.empty:

                active_rows = member_loans[
                    member_loans["Status"] == "Active"
                ]

                if not active_rows.empty:

                    active_loan = active_rows.iloc[-1]

            if active_loan is not None:

                loan_record_id = active_loan["Loan ID"]

                monthly_interest = float(
                    active_loan["Monthly Interest"]
                )

                monthly_principal = float(
                    active_loan["Monthly Principal"]
                )

            else:

                loan_record_id = ""

                monthly_principal = (
                    loan_amount / 36
                )

                monthly_interest = (
                    loan_amount * 0.02
                )

            regular_payment = (
                monthly_principal
                + monthly_interest
            )

            st.info(
                f"👤 **{member['የአባል ስም']}**"
            )

            c1, c2, c3 = st.columns(3)

            c1.metric(
                "የመጀመሪያ ብድር",
                f"{loan_amount:,.2f} ETB"
            )

            c2.metric(
                "የቀረ ዋና ዕዳ",
                f"{remaining_debt:,.2f} ETB"
            )

            c3.metric(
                "መደበኛ ወርሃዊ ክፍያ",
                f"{regular_payment:,.2f} ETB"
            )

            pay_amount = st.number_input(
                "አባል የከፈለው ጠቅላላ ብር:",
                min_value=0.01,
                step=100.0
            )

            description = st.text_input(
                "የክፍያ ማብራሪያ:"
            )

            pay_btn = st.button(
                "💳 ክፍያ መዝግብ",
                type="primary"
            )

            if pay_btn:

                max_payment = (
                    remaining_debt
                    + monthly_interest
                )

                actual_payment = min(
                    pay_amount,
                    max_payment
                )

                actual_interest_paid = min(
                    actual_payment,
                    monthly_interest
                )

                actual_principal_paid = (
                    actual_payment
                    - actual_interest_paid
                )

                new_remaining_debt = (
                    remaining_debt
                    - actual_principal_paid
                )

                if new_remaining_debt < 0:
                    new_remaining_debt = 0.0

                today = datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )

                payment_id = generate_id("PAY")

                new_repayment = pd.DataFrame([{

                    "Payment ID": payment_id,
                    "Date": today,
                    "Member ID": p_id,
                    "Member Name": member["የአባል ስም"],
                    "Loan ID": loan_record_id,
                    "Payment Amount (ETB)": actual_payment,
                    "Interest Paid": actual_interest_paid,
                    "Principal Paid": actual_principal_paid,
                    "Remaining Principal": new_remaining_debt,
                    "Description": description
                }])

                repayments_db = pd.concat(
                    [repayments_db, new_repayment],
                    ignore_index=True
                )

                member["የቀረው ዕዳ (ብር)"] = (
                    new_remaining_debt
                )

                member["የመጨረሻ ክፍያ ቀን"] = today

                if new_remaining_debt <= 0:

                    member["የቀረው ዕዳ (ብር)"] = 0.0

                    member["ብድር ሁኔታ"] = "የለበትም"

                    if loan_record_id:

                        loans_db.loc[
                            loans_db["Loan ID"]
                            == loan_record_id,
                            "Status"
                        ] = "Completed"

                    st.success(
                        f"🎉 {member['የአባል ስም']} "
                        "ብድሩን ሙሉ በሙሉ ከፍሏል!"
                    )

                else:

                    st.success(
                        "✅ ክፍያው ተመዝግቧል!"
                    )

                save_all_data()

                st.info(
                    f"""
                    💳 የተከፈለ፦ {actual_payment:,.2f} ETB

                    🏦 Principal፦ {actual_principal_paid:,.2f} ETB

                    📈 Interest፦ {actual_interest_paid:,.2f} ETB

                    📌 የቀረ ዕዳ፦ {new_remaining_debt:,.2f} ETB
                    """
                )

        else:

            st.info(
                f"💡 {member['የአባል ስም']} "
                "ላይ የብድር ዕዳ የለም።"
            )

    elif p_id:

        st.error(
            "❌ ይህ Member ID አልተገኘም!"
        )


# =========================================================
# 9. REPAYMENT HISTORY
# =========================================================

elif menu == "🧾 የክፍያ ታሪክ":

    st.header("🧾 የብድር ክፍያ ታሪክ")

    if not repayments_db.empty:

        search_id = st.text_input(
            "🔎 Member ID ፈልግ:"
        ).strip()

        display_df = repayments_db.copy()

        if search_id:

            display_df = display_df[
                display_df["Member ID"].astype(str)
                == search_id
            ]

        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True
        )

        total_payment = pd.to_numeric(
            display_df["Payment Amount (ETB)"],
            errors="coerce"
        ).fillna(0).sum()

        total_principal = pd.to_numeric(
            display_df["Principal Paid"],
            errors="coerce"
        ).fillna(0).sum()

        total_interest = pd.to_numeric(
            display_df["Interest Paid"],
            errors="coerce"
        ).fillna(0).sum()

        c1, c2, c3 = st.columns(3)

        c1.metric(
            "ጠቅላላ ክፍያ",
            f"{total_payment:,.2f} ETB"
        )

        c2.metric(
            "Principal",
            f"{total_principal:,.2f} ETB"
        )

        c3.metric(
            "Interest",
            f"{total_interest:,.2f} ETB"
        )

    else:

        st.info(
            "📌 እስካሁን የክፍያ ታሪክ የለም።"
        )


# =========================================================
# 10. GENERAL REPORT
# =========================================================

elif menu == "📊 ጠቅላላ ሪፖርት":

    st.header(
        "📊 ጠቅላላ የአባላት፣ የቁጠባ እና የብድር ሪፖርት"
    )

    if members_db:

        report_data = []

        for member_id, member in members_db.items():

            report_data.append({

                "Member ID": member_id,

                "የአባል ስም":
                    member.get(
                        "የአባል ስም",
                        ""
                    ),

                "ጠቅላላ ቁጠባ":
                    float(
                        member.get(
                            "ጠቅላላ ቁጠባ (ብር)",
                            0
                        ) or 0
                    ),

                "ብድር ሁኔታ":
                    member.get(
                        "ብድር ሁኔታ",
                        "የለበትም"
                    ),

                "ጠቅላላ ብድር":
                    float(
                        member.get(
                            "የተበደረው ጠቅላላ (ብር)",
                            0
                        ) or 0
                    ),

                "90% የተሰጠ":
                    float(
                        member.get(
                            "በእጅ የተሰጠ 90% (ብር)",
                            0
                        ) or 0
                    ),

                "የቀረ ዕዳ":
                    float(
                        member.get(
                            "የቀረው ዕዳ (ብር)",
                            0
                        ) or 0
                    ),

                "የመጨረሻ ቁጠባ":
                    member.get(
                        "የመጨረሻ ቁጠባ ቀን",
                        ""
                    ),

                "የመጨረሻ ክፍያ":
                    member.get(
                        "የመጨረሻ ክፍያ ቀን",
                        ""
                    )
            })

        report_df = pd.DataFrame(
            report_data
        )

        st.dataframe(
            report_df,
            use_container_width=True,
            hide_index=True
        )

        total_savings = report_df[
            "ጠቅላላ ቁጠባ"
        ].sum()

        total_loans = report_df[
            "ጠቅላላ ብድር"
        ].sum()

        total_debt = report_df[
            "የቀረ ዕዳ"
        ].sum()

        c1, c2, c3 = st.columns(3)

        c1.metric(
            "💰 ጠቅላላ ቁጠባ",
            f"{total_savings:,.2f} ETB"
        )

        c2.metric(
            "💵 ጠቅላላ ብድር",
            f"{total_loans:,.2f} ETB"
        )

        c3.metric(
            "📌 ጠቅላላ የቀረ ዕዳ",
            f"{total_debt:,.2f} ETB"
        )

        csv_data = report_df.to_csv(
            index=False
        ).encode("utf-8-sig")

        st.download_button(
            "📥 ሪፖርቱን CSV አውርድ",
            data=csv_data,
            file_name="sacco_report.csv",
            mime="text/csv"
        )

    else:

        st.info(
            "📌 እስካሁን የተመዘገበ መረጃ የለም።"
        )
```

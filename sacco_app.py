import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="የቁጠባ እና ብድር ሲስተም", layout="wide")
st.title("🏦 ተስፋ የገንዘብ ቁጠባና ብድር ማህበር")

# 💾 በቋሚነት ኮምፒውተር ላይ መረጃ መያዣ የኤክሴል ፋይል
DB_FILE = "sacco_database.xlsx"

# ፋይሉ አስቀድሞ ካለ ያነባል፣ ከሌለ አዲስ ባዶ መዋቅር ይፈጥራል
if os.path.exists(DB_FILE):
    try:
        df_load = pd.read_excel(DB_FILE, dtype={"መታወቂያ ቁጥር (ID)": str})
        members_db = df_load.set_index("መታወቂያ ቁጥር (ID)").to_dict(orient="index")
    except Exception:
        members_db = {}
else:
    members_db = {}

# መረጃን ወደ ኤክሴል ፋይል የመቆጠቢያ ተግባር
def save_data_to_excel(db):
    if db:
        df_save = pd.DataFrame.from_dict(db, orient="index")
        df_save.index.name = "መታወቂያ ቁጥር (ID)"
        df_save.reset_index().to_excel(DB_FILE, index=False)
    else:
        if os.path.exists(DB_FILE):
            try:
                os.remove(DB_FILE)
            except Exception:
                pass

# የጎንዮሽ ማውጫ (Sidebar)
menu = st.sidebar.selectbox("ያሉ አማራጮች", ["👤 አባል መመዝገቢያ", "💰 የወር ቁጠባ ማስገቢያ", "💵 የብድር አገልግሎት", "📅 የብድር ክፍያ መመዝገቢያ", "📊 ጠቅላላ ሪፖርት"])

# --- 1. አባል መመዝገቢያ ገጽ ---
if menu == "👤 አባል መመዝገቢያ":
    st.header("👤 አዲስ አባል መመዝገቢያ ፎርም")
    m_id = st.text_input("የአባል መታወቂያ ቁጥር (ID):")
    m_name = st.text_input("የአባል ሙሉ ስም:")
    register_btn = st.button("አባል መዝግብ")
    
    if register_btn:
        if m_id and m_name:
            if m_id in members_db:
                st.error(f"❌ ስህተት፦ መታወቂያ ቁጥር {m_id} ቀደም ብሎ ተመዝግቧል!")
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
                st.success(f"✅ አባል {m_name} በተሳካ ሁኔታ ተመዝግቧል! (መረጃው በኤክሴል ተቀምጧል)")
        else:
            st.warning("⚠️ እባክዎ ሁሉንም ሳጥኖች ይሙሉ!")

# --- 2. የወር ቁጠባ ማስገቢያ ገጽ ---
elif menu == "💰 የወር ቁጠባ ማስገቢያ":
    st.header("💰 የወርሃዊ ቁጠባ መመዝገቢያ")
    s_id = st.text_input("የአባል መታወቂያ (ID):")
    amount = st.number_input("የሚቆጥበው የገንዘብ መጠን (ብር):", min_value=0.0, step=100.0)
    save_btn = st.button("ቁጠባ መዝግብ")
    
    if save_btn:
        if s_id in members_db:
            members_db[s_id]["ጠቅላላ ቁጠባ (ብር)"] += amount
            save_data_to_excel(members_db)
            st.success(f"✅ ለ{members_db[s_id]['የአባል ስም']} {amount:,.2f} ብር ቁጠባ ተመዝግቧል።")
        else:
            st.error("❌ ይህ መታወቂያ በሲስተሙ ውስጥ አልተገኘም!")

# --- 3. የብድር አገልግሎት ገጽ ---
elif menu == "💵 የብድር አገልግሎት":
    st.header("💵 የብድር ማመልከቻ እና ስሌት")
    loan_id = st.text_input("የተበዳሪው አባል መታወቂያ (ID):")
    loan_amount = st.number_input("የሚጠይቀው የብድር መጠን (ብር):", min_value=0.0, step=1000.0)
    calculate_btn = st.button("ብድር አስላ እና ፍቀድ")
    
    if calculate_btn:
        if loan_id in members_db:
            member = members_db[loan_id]
            if member["ብድር ሁኔታ"] == "ያለበት":
                st.error(f"❌ ስህተት፦ {member['የአባል ስም']} የድሮ ብድር ስላለበት ተጨማሪ መበደር አይችልም!")
            else:
                upfront_fee = loan_amount * 0.10
                net_payout = loan_amount - upfront_fee
                monthly_principal = loan_amount / 36
                monthly_interest = loan_amount * 0.02
                total_monthly = monthly_principal + monthly_interest
                
                member["ብድር ሁኔታ"] = "ያለበት"
                member["የተበደረው ጠቅላላ (ብር)"] = loan_amount
                member["በእጅ የተሰጠ 90% (ብር)"] = net_payout
                member["የቀረው ዕዳ (ብር)"] = loan_amount
                
                save_data_to_excel(members_db)
                st.success(f"🎉 ለ{member['የአባል ስም']} ብድር ተፈቅዷል!")
                st.info(f"💵 በእጅ የሚሰጠው ገንዘብ (90%)፦ {net_payout:,.2f} ብር \n\n"
                        f"📅 የወርሃዊ ክፍያ (ለ36 ወራት)፦ {total_monthly:,.2f} ብር (ዋና፦ {monthly_principal:,.2f} + ወለድ 2%፦ {monthly_interest:,.2f})")
        else:
            st.error("❌ ይህ መታወቂያ በሲስተሙ ውስጥ አልተገኘም!")

# --- 4. የብድር ክፍያ መመዝገቢያ ገጽ ---
elif menu == "📅 የብድር ክፍያ መመዝገቢያ":
    st.header("📅 የወርሃዊ ብድር ክፍያ መቀበያ")
    p_id = st.text_input("የከፋይ አባል መታወቂያ (ID):")
    
    if p_id in members_db:
        member = members_db[p_id]
        if member["ብድር ሁኔታ"] == "ያለበት":
            st.warning(f"📌 {member['የአባል ስም']} ያለበት ጠቅላላ ዕዳ፦ {member['የቀረው ዕዳ (ብር)']:,.2f} ብር")
            monthly_principal = member["የተበደረው ጠቅላላ (ብር)"] / 36
            monthly_interest = member["የተበደረው ጠቅላላ (ብር)"] * 0.02
            st.write(f"💡 መደበኛ የወር ክፍያ፦ **{monthly_principal + monthly_interest:,.2f} ብር** (ዋና፦ {monthly_principal:,.2f} + ወለድ፦ {monthly_interest:,.2f})")
            
            pay_amount = st.number_input("አባል አሁን የከፈለው ጠቅላላ ብር (ዋና + ወለድ)፦", min_value=0.0)
            pay_btn = st.button("ክፍያ መዝግብ")
            
            if pay_btn:
                actual_principal_paid = pay_amount - monthly_interest
                if actual_principal_paid < 0:
                    actual_principal_paid = 0
                
                member["የቀረው ዕዳ (ብር)"] -= actual_principal_paid
                if member["የቀረው ዕዳ (ብር)"] <= 0:
                    member["የቀረው ዕዳ (ብር)"] = 0.0
                    member["ብድር ሁኔታ"] = "የለበትም"
                    st.success(f"🎉 {member['የአባል ስም']} ብድሩን ሙሉ በሙሉ ከፍሎ ጨርሷል!")
                else:
                    st.success(f"✅ ክፍያ ተመዝግቧል። የቀረው ጠቅላላ ዋና ዕዳ፦ {member['የቀረው ዕዳ (ብር)']:,.2f} ብር")
                save_data_to_excel(members_db)
        else:
            st.info(f"💡 {member['የአባል ስም']} ላይ ምንም ዓይነት የብድር ዕዳ የለም።")
    elif p_id:
        st.error("❌ ይህ መታወቂያ በሲስተሙ ውስጥ አልተገኘም!")

# --- 5. የሪፖርት ገጽ ---
elif menu == "📊 ጠቅላላ ሪፖርት":
    st.header("📊 ጠቅላላ የአባላት፣ የቁጠባ እና የብድር ሪፖርት")
    if members_db:
        df = pd.DataFrame.from_dict(members_db, orient="index")
        st.dataframe(df, use_container_width=True)
    else:
        st.info("📌 እስካሁን በሲስተሙ ላይ የተመዘገበ መረጃ የለም።")

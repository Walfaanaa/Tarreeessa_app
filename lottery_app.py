import streamlit as st
import pandas as pd
from io import BytesIO
import time
import os
from dotenv import load_dotenv

# ============================================================
# 1️⃣ PAGE SETUP
# ============================================================
st.set_page_config(
    page_title="🎟️ EGSA Lottery Winners",
    layout="wide",
    page_icon="🎟️"
)

# ============================================================
# 🎨 CUSTOM UI STYLE
# ============================================================
st.markdown("""
<style>
body {
    background-color: #1E90FF;
}

[data-testid="stAppViewContainer"] {
    background-color: #1E90FF;
}

.dataframe,
.stDataFrame>div>div>div>div>table {
    background-color: #87CEFA !important;
    color: #000000 !important;
}

.custom-warning {
    background-color: #104E8B;
    color: #00FFFF;
    padding: 10px;
    border-radius: 5px;
    text-align: left;
    margin-bottom: 10px;
    font-weight: bold;
}

.header-section {
    background-color: red;
    padding: 20px;
    border-radius: 10px;
    text-align: center;
    color: white;
    font-family: Arial, sans-serif;
}

h1, h3 {
    text-align: center;
    color: white;
}

.stButton>button {
    background-color: #FF0000;
    color: white;
    border-radius: 12px;
    height: 3em;
    width: 220px;
    font-size: 18px;
    font-weight: bold;
}
</style>
""", unsafe_allow_html=True)

# ============================================================
# 🎨 HEADER
# ============================================================
st.markdown("""
<div class="header-section">
    <h1>🎟️ EGSA Lottery Winners App</h1>
    <h3>Welcome to the EGSA Uqqubii Lottery Winners App</h3>
    <p>
        This system ensures fair, transparent, and one-time-only
        draws managed by authorized personnel.
    </p>
</div>
""", unsafe_allow_html=True)

# ============================================================
# 2️⃣ FILE CONFIGURATION
# ============================================================
DATA_FILE = "Tarreessa.xlsx"
WINNER_FILE = "winners_record.xlsx"

# ============================================================
# 3️⃣ LOAD MEMBERS DATA
# ============================================================
try:
    members_df = pd.read_excel(DATA_FILE)

    st.success(
        f"✅ {len(members_df)} members loaded successfully from admin file."
    )

    display_members = members_df.copy()
    display_members.index = range(1, len(display_members) + 1)

    st.subheader("👥 Member List")
    st.dataframe(
        display_members,
        use_container_width=True
    )

except FileNotFoundError:
    st.error(
        "❌ Tarreessa.xlsx file not found! "
        "Please upload it to your app folder or GitHub repository."
    )
    st.stop()

except Exception as e:
    st.error(f"❌ Error loading Tarreessa.xlsx: {e}")
    st.stop()

# ============================================================
# 4️⃣ LOAD ADMIN PASSWORDS
# ============================================================
load_dotenv()

AUTHORIZED_CODE = os.getenv("STREAMLIT_ADMIN_PASSWORD")
RESET_PASSWORD = os.getenv("STREAMLIT_RESET_PASSWORD")

# ============================================================
# STREAMLIT CLOUD SECRETS FALLBACK
# ============================================================
if not AUTHORIZED_CODE:
    try:
        AUTHORIZED_CODE = st.secrets.get(
            "STREAMLIT_ADMIN_PASSWORD"
        )
    except Exception:
        pass

if not RESET_PASSWORD:
    try:
        RESET_PASSWORD = st.secrets.get(
            "STREAMLIT_RESET_PASSWORD"
        )
    except Exception:
        pass

# ============================================================
# PASSWORD CONFIGURATION CHECK
# ============================================================
if not AUTHORIZED_CODE:
    st.markdown(
        '<div class="custom-warning">'
        '⚠️ Admin password is not configured. '
        'Add STREAMLIT_ADMIN_PASSWORD to your .env file '
        'or Streamlit Secrets.'
        '</div>',
        unsafe_allow_html=True
    )

if not RESET_PASSWORD:
    st.markdown(
        '<div class="custom-warning">'
        '⚠️ Reset password is not configured. '
        'Add STREAMLIT_RESET_PASSWORD to your .env file '
        'or Streamlit Secrets.'
        '</div>',
        unsafe_allow_html=True
    )

# ============================================================
# 5️⃣ ADMIN AUTHORIZATION
# ============================================================
st.markdown("### 🔐 Administrator Access")

password = st.text_input(
    "Enter admin passcode to enable draw:",
    type="password",
    key="admin_password"
)

# ============================================================
# AUTHORIZED USER
# ============================================================
if password and AUTHORIZED_CODE and password == AUTHORIZED_CODE:

    st.success(
        "✅ Access granted! You can now enable the draw."
    )

    st.markdown("---")

    # ========================================================
    # 6️⃣ PREVIOUS WINNERS
    # ========================================================
    if os.path.exists(WINNER_FILE):

        st.warning(
            "⚠️ A previous draw has already been conducted."
        )

        try:
            previous_winners = pd.read_excel(WINNER_FILE)

            display_previous = previous_winners.copy()
            display_previous.index = range(
                1,
                len(display_previous) + 1
            )

            st.subheader("🎉 Previous Winners")

            st.dataframe(
                display_previous,
                use_container_width=True
            )

        except Exception as e:
            st.error(
                f"❌ Could not read winners_record.xlsx: {e}"
            )

        # ====================================================
        # ADMIN RESET
        # ====================================================
        with st.expander("⚙️ Admin Reset Options"):

            st.warning(
                "Resetting will delete the current winners record "
                "and allow a new lottery round."
            )

            reset_pass_input = st.text_input(
                "Enter reset password",
                type="password",
                key="reset_password"
            )

            if st.button(
                "🔄 Reset for New Round",
                key="reset_button"
            ):

                if (
                    RESET_PASSWORD
                    and reset_pass_input == RESET_PASSWORD
                ):

                    try:
                        os.remove(WINNER_FILE)

                        st.success(
                            "✅ Winners record deleted successfully."
                        )

                        time.sleep(1)
                        st.rerun()

                    except Exception as e:
                        st.error(
                            f"❌ Unable to reset winners record: {e}"
                        )

                else:
                    st.error(
                        "❌ Incorrect reset password."
                    )

    # ========================================================
    # 7️⃣ NEW DRAW
    # ========================================================
    else:

        st.success(
            "🟢 No previous draw found. "
            "The lottery is ready for a new draw."
        )

        st.markdown("### 🎯 Lottery Draw")

        num_winners = st.number_input(
            "🏆 Number of winners to select",
            min_value=1,
            max_value=len(members_df),
            value=1,
            step=1
        )

        st.info(
            f"🎟️ You are going to select "
            f"**{num_winners} winner(s)** from "
            f"**{len(members_df)} members**."
        )

        # ====================================================
        # PICK WINNERS BUTTON
        # ====================================================
        if st.button(
            "🎲 PICK WINNERS",
            key="pick_winners",
            type="primary"
        ):

            placeholder = st.empty()

            with placeholder.container():

                st.info(
                    "🎲 Picking winners... Please wait."
                )

                progress_text = st.empty()
                progress_bar = st.progress(0)

                for i in range(101):

                    time.sleep(0.01)

                    progress_text.text(
                        f"Drawing in progress: {i}%"
                    )

                    progress_bar.progress(i)

                # =================================================
                # RANDOM WINNER SELECTION
                # =================================================
                winners = members_df.sample(
                    n=int(num_winners),
                    replace=False
                ).reset_index(drop=True)

                # =================================================
                # DISPLAY WINNERS
                # =================================================
                display_winners = winners.copy()

                display_winners.index = range(
                    1,
                    len(display_winners) + 1
                )

                st.success(
                    "🎉🎉 WINNERS SELECTED SUCCESSFULLY! 🎉🎉"
                )

                st.balloons()

                st.subheader("🏆 Winners List")

                st.dataframe(
                    display_winners,
                    use_container_width=True
                )

                # =================================================
                # SAVE WINNERS
                # =================================================
                try:

                    winners.to_excel(
                        WINNER_FILE,
                        index=False,
                        engine="openpyxl"
                    )

                    st.success(
                        f"💾 Winners saved to `{WINNER_FILE}`."
                    )

                except Exception as e:

                    st.error(
                        f"❌ Could not save winners record: {e}"
                    )

                # =================================================
                # EXCEL DOWNLOAD
                # =================================================
                def convert_df_to_excel(df):

                    output = BytesIO()

                    with pd.ExcelWriter(
                        output,
                        engine="xlsxwriter"
                    ) as writer:

                        df.to_excel(
                            writer,
                            index=False,
                            sheet_name="Winners"
                        )

                    return output.getvalue()

                excel_data = convert_df_to_excel(
                    winners
                )

                st.download_button(
                    label="💾 Download Winners as Excel",
                    data=excel_data,
                    file_name="EGSA_lottery_winners.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    key="download_winners"
                )

# ============================================================
# INVALID PASSWORD
# ============================================================
elif password:

    st.error(
        "❌ Invalid passcode. Access denied."
    )

    st.info(
        "You can view the member list, but only authorized "
        "staff can pick winners."
    )

# ============================================================
# NO PASSWORD ENTERED
# ============================================================
else:

    st.info(
        "🔐 Enter the administrator passcode above "
        "to enable the lottery draw."
    )

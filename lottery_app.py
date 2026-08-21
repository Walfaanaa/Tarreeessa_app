import streamlit as st
import pandas as pd
from io import BytesIO
import time
import os
import requests
import base64
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

h1,
h3 {
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

    <h3>
        Welcome to the EGSA Uqqubii Lottery Winners App
    </h3>

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
# 3️⃣ GITHUB CONFIGURATION
# ============================================================

GITHUB_OWNER = "Walfaanaa"

GITHUB_REPO = "Tarreeessa_app"

GITHUB_BRANCH = "main"

GITHUB_DATA_PATH = "Tarreessa.xlsx"

GITHUB_WINNER_PATH = "winners_record.xlsx"


# ============================================================
# 4️⃣ LOAD ENVIRONMENT VARIABLES
# ============================================================

load_dotenv()


# ============================================================
# 5️⃣ LOAD PASSWORDS
# ============================================================

AUTHORIZED_CODE = os.getenv(
    "STREAMLIT_ADMIN_PASSWORD"
)

RESET_PASSWORD = os.getenv(
    "STREAMLIT_RESET_PASSWORD"
)


# ============================================================
# 6️⃣ LOAD GITHUB TOKEN
# ============================================================

GITHUB_TOKEN = os.getenv(
    "GITHUB_TOKEN"
)


# ============================================================
# STREAMLIT CLOUD SECRETS FALLBACK
# ============================================================

try:

    if not AUTHORIZED_CODE:

        AUTHORIZED_CODE = st.secrets.get(
            "STREAMLIT_ADMIN_PASSWORD"
        )

    if not RESET_PASSWORD:

        RESET_PASSWORD = st.secrets.get(
            "STREAMLIT_RESET_PASSWORD"
        )

    if not GITHUB_TOKEN:

        GITHUB_TOKEN = st.secrets.get(
            "GITHUB_TOKEN"
        )

except Exception:

    pass


# ============================================================
# 7️⃣ GITHUB HEADERS
# ============================================================

def github_headers():

    return {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28"
    }


# ============================================================
# 8️⃣ DOWNLOAD FILE FROM GITHUB
# ============================================================

def download_file_from_github(file_path):

    url = (
        f"https://api.github.com/repos/"
        f"{GITHUB_OWNER}/"
        f"{GITHUB_REPO}/"
        f"contents/"
        f"{file_path}"
    )

    response = requests.get(
        url,
        headers=github_headers(),
        params={
            "ref": GITHUB_BRANCH
        },
        timeout=30
    )

    response.raise_for_status()

    response_data = response.json()

    encoded_content = response_data["content"]

    file_bytes = base64.b64decode(
        encoded_content
    )

    return file_bytes


# ============================================================
# 9️⃣ UPLOAD / UPDATE FILE ON GITHUB
# ============================================================

def upload_file_to_github(
    file_path,
    file_bytes,
    commit_message
):

    url = (
        f"https://api.github.com/repos/"
        f"{GITHUB_OWNER}/"
        f"{GITHUB_REPO}/"
        f"contents/"
        f"{file_path}"
    )

    # --------------------------------------------------------
    # Get existing file SHA
    # --------------------------------------------------------

    response = requests.get(
        url,
        headers=github_headers(),
        params={
            "ref": GITHUB_BRANCH
        },
        timeout=30
    )

    sha = None

    if response.status_code == 200:

        sha = response.json()["sha"]

    elif response.status_code != 404:

        response.raise_for_status()

    # --------------------------------------------------------
    # Encode file
    # --------------------------------------------------------

    encoded_content = base64.b64encode(
        file_bytes
    ).decode("utf-8")

    # --------------------------------------------------------
    # GitHub payload
    # --------------------------------------------------------

    payload = {
        "message": commit_message,
        "content": encoded_content,
        "branch": GITHUB_BRANCH
    }

    if sha:

        payload["sha"] = sha

    # --------------------------------------------------------
    # Upload
    # --------------------------------------------------------

    response = requests.put(
        url,
        headers=github_headers(),
        json=payload,
        timeout=30
    )

    response.raise_for_status()

    return response.json()


# ============================================================
# 🔟 CREATE EXCEL BYTES
# ============================================================

def dataframe_to_excel_bytes(df):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Members"
        )

    return output.getvalue()


# ============================================================
# 1️⃣1️⃣ LOAD MEMBERS DATA
# ============================================================

try:

    # --------------------------------------------------------
    # First try local file
    # --------------------------------------------------------

    if os.path.exists(DATA_FILE):

        members_df = pd.read_excel(
            DATA_FILE
        )

    # --------------------------------------------------------
    # If local file doesn't exist, download from GitHub
    # --------------------------------------------------------

    else:

        if not GITHUB_TOKEN:

            st.error(
                "❌ Tarreessa.xlsx not found locally "
                "and GITHUB_TOKEN is not configured."
            )

            st.stop()

        github_data = download_file_from_github(
            GITHUB_DATA_PATH
        )

        members_df = pd.read_excel(
            BytesIO(github_data)
        )

        # Save local copy

        with open(
            DATA_FILE,
            "wb"
        ) as f:

            f.write(github_data)

    # --------------------------------------------------------
    # Display members
    # --------------------------------------------------------

    st.success(
        f"✅ {len(members_df)} members loaded successfully."
    )

    display_members = members_df.copy()

    display_members.index = range(
        1,
        len(display_members) + 1
    )

    st.subheader(
        "👥 Member List"
    )

    st.dataframe(
        display_members,
        use_container_width=True
    )

except FileNotFoundError:

    st.error(
        "❌ Tarreessa.xlsx file not found!"
    )

    st.stop()

except Exception as e:

    st.error(
        f"❌ Error loading member data: {e}"
    )

    st.stop()


# ============================================================
# 1️⃣2️⃣ PASSWORD CONFIGURATION CHECK
# ============================================================

if not AUTHORIZED_CODE:

    st.markdown(
        '<div class="custom-warning">'
        '⚠️ Admin password is not configured. '
        'Add STREAMLIT_ADMIN_PASSWORD to your '
        '.env file or Streamlit Secrets.'
        '</div>',
        unsafe_allow_html=True
    )


if not RESET_PASSWORD:

    st.markdown(
        '<div class="custom-warning">'
        '⚠️ Reset password is not configured. '
        'Add STREAMLIT_RESET_PASSWORD to your '
        '.env file or Streamlit Secrets.'
        '</div>',
        unsafe_allow_html=True
    )


if not GITHUB_TOKEN:

    st.markdown(
        '<div class="custom-warning">'
        '⚠️ GITHUB_TOKEN is not configured. '
        'GitHub synchronization will not work.'
        '</div>',
        unsafe_allow_html=True
    )


# ============================================================
# 1️⃣3️⃣ ADMIN AUTHORIZATION
# ============================================================

st.markdown(
    "### 🔐 Administrator Access"
)


password = st.text_input(
    "Enter admin passcode to enable draw:",
    type="password",
    key="admin_password"
)


# ============================================================
# 1️⃣4️⃣ AUTHORIZED USER
# ============================================================

if (
    password
    and AUTHORIZED_CODE
    and password == AUTHORIZED_CODE
):

    st.success(
        "✅ Access granted! You can now enable the draw."
    )

    st.markdown("---")


    # ========================================================
    # 1️⃣5️⃣ CHECK PREVIOUS WINNERS
    # ========================================================

    if os.path.exists(WINNER_FILE):

        st.warning(
            "⚠️ A previous draw has already been conducted."
        )


        # ====================================================
        # LOAD PREVIOUS WINNERS
        # ====================================================

        try:

            previous_winners = pd.read_excel(
                WINNER_FILE
            )

            display_previous = previous_winners.copy()

            display_previous.index = range(
                1,
                len(display_previous) + 1
            )

            st.subheader(
                "🎉 Previous Winners"
            )

            st.dataframe(
                display_previous,
                use_container_width=True
            )

        except Exception as e:

            st.error(
                f"❌ Could not read winners_record.xlsx: {e}"
            )

            st.stop()


        # ====================================================
        # ADMIN RESET
        # ====================================================

        with st.expander(
            "⚙️ Admin Reset Options"
        ):

            st.warning(
                "Resetting will remove the previous winners "
                "from the available member list and prepare "
                "the system for a new lottery round."
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


                # =================================================
                # CHECK RESET PASSWORD
                # =================================================

                if (
                    RESET_PASSWORD
                    and reset_pass_input == RESET_PASSWORD
                ):


                    try:

                        # =========================================
                        # CHECK GITHUB TOKEN
                        # =========================================

                        if not GITHUB_TOKEN:

                            st.error(
                                "❌ GITHUB_TOKEN is not configured."
                            )

                            st.stop()


                        # =========================================
                        # MAKE COPY
                        # =========================================

                        updated_members = members_df.copy()


                        # =========================================
                        # FIND COMMON COLUMNS
                        # =========================================

                        common_columns = [
                            col
                            for col in previous_winners.columns
                            if col in updated_members.columns
                        ]


                        if not common_columns:

                            st.error(
                                "❌ No common columns were found "
                                "between Tarreessa.xlsx and "
                                "winners_record.xlsx."
                            )

                            st.stop()


                        # =========================================
                        # REMOVE PREVIOUS WINNERS
                        # =========================================

                        before_count = len(
                            updated_members
                        )


                        updated_members = (
                            updated_members.merge(
                                previous_winners[
                                    common_columns
                                ].drop_duplicates(),

                                on=common_columns,

                                how="left",

                                indicator=True
                            )
                        )


                        updated_members = (
                            updated_members[
                                updated_members["_merge"]
                                == "left_only"
                            ]
                            .drop(
                                columns=["_merge"]
                            )
                        )


                        after_count = len(
                            updated_members
                        )


                        removed_count = (
                            before_count
                            - after_count
                        )


                        # =========================================
                        # CREATE UPDATED TARREESSA.XLSX
                        # =========================================

                        updated_data = (
                            dataframe_to_excel_bytes(
                                updated_members
                            )
                        )


                        # =========================================
                        # UPDATE TARREESSA.XLSX ON GITHUB
                        # =========================================

                        upload_file_to_github(

                            GITHUB_DATA_PATH,

                            updated_data,

                            "Remove previous lottery winners "
                            "from available members"
                        )


                        # =========================================
                        # UPDATE LOCAL TARREESSA.XLSX
                        # =========================================

                        with open(
                            DATA_FILE,
                            "wb"
                        ) as f:

                            f.write(
                                updated_data
                            )


                        # =========================================
                        # DELETE LOCAL WINNER RECORD
                        # =========================================

                        os.remove(
                            WINNER_FILE
                        )


                        # =========================================
                        # SUCCESS
                        # =========================================

                        st.success(
                            "✅ Reset completed successfully!"
                        )

                        st.info(
                            f"🏆 Previous winners removed: "
                            f"{removed_count}"
                        )

                        st.info(
                            f"👥 Remaining members: "
                            f"{after_count}"
                        )

                        st.success(
                            "☁️ Tarreessa.xlsx has been "
                            "updated on GitHub."
                        )

                        st.success(
                            "🔄 The system is ready for "
                            "the next lottery round."
                        )


                        time.sleep(2)

                        st.rerun()


                    except Exception as e:

                        st.error(
                            f"❌ Reset failed: {e}"
                        )


                else:

                    st.error(
                        "❌ Incorrect reset password."
                    )


    # ========================================================
    # 1️⃣6️⃣ NEW DRAW
    # ========================================================

    else:

        st.success(
            "🟢 No previous draw found. "
            "The lottery is ready for a new draw."
        )


        st.markdown(
            "### 🎯 Lottery Draw"
        )


        # ====================================================
        # NUMBER OF WINNERS
        # ====================================================

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
        # PICK WINNERS
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

                progress_bar = st.progress(
                    0
                )


                for i in range(101):

                    time.sleep(
                        0.01
                    )

                    progress_text.text(
                        f"Drawing in progress: {i}%"
                    )

                    progress_bar.progress(
                        i
                    )


                # =============================================
                # RANDOM WINNER SELECTION
                # =============================================

                winners = members_df.sample(

                    n=int(num_winners),

                    replace=False

                ).reset_index(
                    drop=True
                )


                # =============================================
                # DISPLAY WINNERS
                # =============================================

                display_winners = winners.copy()

                display_winners.index = range(
                    1,
                    len(display_winners) + 1
                )


                st.success(
                    "🎉🎉 WINNERS SELECTED SUCCESSFULLY! 🎉🎉"
                )


                st.balloons()


                st.subheader(
                    "🏆 Winners List"
                )


                st.dataframe(
                    display_winners,
                    use_container_width=True
                )


                # =============================================
                # SAVE WINNERS LOCALLY
                # =============================================

                try:

                    winners.to_excel(
                        WINNER_FILE,
                        index=False,
                        engine="openpyxl"
                    )


                    # =========================================
                    # CREATE WINNER EXCEL BYTES
                    # =========================================

                    winner_output = BytesIO()


                    with pd.ExcelWriter(
                        winner_output,
                        engine="openpyxl"
                    ) as writer:

                        winners.to_excel(
                            writer,
                            index=False,
                            sheet_name="Winners"
                        )


                    winner_bytes = (
                        winner_output.getvalue()
                    )


                    # =========================================
                    # UPLOAD WINNERS TO GITHUB
                    # =========================================

                    if GITHUB_TOKEN:

                        upload_file_to_github(

                            GITHUB_WINNER_PATH,

                            winner_bytes,

                            "Save new lottery winners"
                        )


                        st.success(
                            "☁️ Winners record saved "
                            "to GitHub successfully."
                        )

                    else:

                        st.warning(
                            "⚠️ GITHUB_TOKEN is not configured. "
                            "Winner record was saved locally only."
                        )


                    st.success(
                        f"💾 Winners saved to "
                        f"`{WINNER_FILE}`."
                    )


                except Exception as e:

                    st.error(
                        f"❌ Could not save winners record: {e}"
                    )


                # =============================================
                # EXCEL DOWNLOAD
                # =============================================

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


                excel_data = (
                    convert_df_to_excel(
                        winners
                    )
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
# 1️⃣7️⃣ INVALID PASSWORD
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
# 1️⃣8️⃣ NO PASSWORD
# ============================================================

else:

    st.info(
        "🔐 Enter the administrator passcode above "
        "to enable the lottery draw."
    )

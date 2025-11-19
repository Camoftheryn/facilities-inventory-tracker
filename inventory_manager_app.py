import streamlit as st
import pandas as pd
import os
from datetime import datetime
import openpyxl
import warnings

# File paths
EXCEL_FILE = "INVTRCKR.xlsm"
LOG_FILE = "inventory_log.csv"

# Suppress openpyxl warning
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# -----------------------------
# NORMALIZATION FUNCTION
# -----------------------------
def normalize_code(value):
    """Clean and normalize barcode + Tool ID to allow flexible matching."""
    if pd.isna(value):
        return ""

    value = str(value).strip().lower()

    # Remove common inconsistent characters
    remove_chars = [" ", "-", "_", "*", "/", "\\", ".", ":", ";"]
    for c in remove_chars:
        value = value.replace(c, "")

    # Remove leading zeros (UPC/EAN)
    value = value.lstrip("0")

    return value


# Session state
if "username" not in st.session_state:
    st.session_state.username = ""
if "status_message" not in st.session_state:
    st.session_state.status_message = None
if "clear_barcode_input" not in st.session_state:
    st.session_state.clear_barcode_input = False
if "barcode_input" not in st.session_state:
    st.session_state.barcode_input = ""

# Load inventory
if os.path.exists(EXCEL_FILE):
    inventory_df = pd.read_excel(EXCEL_FILE, engine="openpyxl")
    inventory_df.columns = inventory_df.columns.str.strip()
else:
    inventory_df = pd.DataFrame(columns=[
        "Tool ID", "check in", "check out", "Total Count",
        "Checked Out Qty", "Running Total"
    ])

# Add normalized column
inventory_df["Normalized_ID"] = inventory_df["Tool ID"].apply(normalize_code)

# Load log
if os.path.exists(LOG_FILE):
    log_df = pd.read_csv(LOG_FILE)
else:
    log_df = pd.DataFrame(columns=[
        "Timestamp", "Action", "Name", "Barcode", "Quantity", "User"
    ])

# Save inventory and logs
def save_inventory(df):
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a',
                        if_sheet_exists='replace', engine_kwargs={"keep_vba": True}) as writer:
        df.to_excel(writer, index=False)

def log_action(action, name, barcode, qty, user):
    global log_df
    log_entry = {
        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Action": action,
        "Name": name,
        "Barcode": barcode,
        "Quantity": qty,
        "User": user
    }
    log_df = pd.concat([log_df, pd.DataFrame([log_entry])], ignore_index=True)
    log_df.to_csv(LOG_FILE, index=False)


# Title
st.title("Inventory & Supply Room Manager")

# User name input
st.sidebar.subheader("User Access")

input_name = st.sidebar.text_input("Enter your name to continue",
                                   value=st.session_state.username)
if st.sidebar.button("Submit Name"):
    if input_name.strip():
        st.session_state.username = input_name.strip()
    else:
        st.sidebar.warning("Please enter your name")

if not st.session_state.username:
    st.stop()

username = st.session_state.username

# Inventory table
st.subheader("Inventory Table")
st.dataframe(inventory_df)

st.markdown("---")
st.subheader("Check Out or Return Items")

# Clear barcode input if flagged
if st.session_state.clear_barcode_input:
    st.session_state.clear_barcode_input = False
    st.session_state.barcode_input = ""
    st.rerun()

with st.form("check_form"):
    barcode = st.text_input("Scan or enter item barcode",
                            key="barcode_input",
                            value=st.session_state.barcode_input)
    st.write("Scanned barcode:", barcode)
    action_type = st.selectbox("Action", ["Check Out", "Return"])
    quantity = st.number_input("Quantity", min_value=1, step=1)
    submitted = st.form_submit_button("Submit")

    if submitted:
        st.session_state.clear_barcode_input = True

        # ---------------------------------
        # FINAL, ERROR-FREE MATCHING LOGIC
        # ---------------------------------
        normalized_barcode = normalize_code(barcode)
        inventory_df["Normalized_ID"] = inventory_df["Tool ID"].apply(normalize_code)

        norm_ids = inventory_df["Normalized_ID"].fillna("")

        # 1. Exact match
        match = inventory_df[norm_ids == normalized_barcode]

        # 2. Prefix/suffix match (safe with apply)
        if match.empty:
            match = inventory_df[
                norm_ids.str.startswith(normalized_barcode) |
                norm_ids.str.endswith(normalized_barcode) |
                norm_ids.apply(lambda x: normalized_barcode.startswith(x)) |
                norm_ids.apply(lambda x: normalized_barcode.endswith(x))
            ]

        # 3. Contains match
        if match.empty:
            match = inventory_df[
                norm_ids.str.contains(normalized_barcode, na=False)
            ]

        # ---------------------------------
        # END MATCHING LOGIC
        # ---------------------------------

        if not match.empty:
            index = match.index[0]
            current_qty = match.at[index, "Running Total"]
            item_name = match.at[index, "Tool ID"]

            if action_type == "Check Out":

import streamlit as st
import pandas as pd
import os
from datetime import datetime
import openpyxl
import warnings

# File paths
EXCEL_FILE = "INVTRCKR.xlsm"  # updated for macro support
LOG_FILE = "inventory_log.csv"

# Suppress openpyxl warning
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Load inventory
if os.path.exists(EXCEL_FILE):
    inventory_df = pd.read_excel(EXCEL_FILE, engine="openpyxl")
    inventory_df.columns = inventory_df.columns.str.strip()  # Clean column names
else:
    inventory_df = pd.DataFrame(columns=["Tool ID", "check in", "check out", "Total Count", "Checked Out Qty", "Running Total"])

# Load log
if os.path.exists(LOG_FILE):
    log_df = pd.read_csv(LOG_FILE)
else:
    log_df = pd.DataFrame(columns=["Timestamp", "Action", "Name", "Barcode", "Quantity", "User"])

# Save inventory and logs
def save_inventory(df):
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace', engine_kwargs={"keep_vba": True}) as writer:
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

if "username" not in st.session_state:
    st.session_state.username = ""

input_name = st.sidebar.text_input("Enter your name to continue", value=st.session_state.username)
if st.sidebar.button("Submit Name"):
    if input_name.strip():
        st.session_state.username = input_name.strip()
    else:
        st.sidebar.warning("Please enter your name")

if not st.session_state.username:
    st.stop()

username = st.session_state.username

# Inventory interaction interface
st.subheader("Inventory Table")
st.dataframe(inventory_df)

st.markdown("---")
st.subheader("Check Out or Return Items")

if "barcode_input" not in st.session_state:
    st.session_state.barcode_input = ""

with st.form("check_form"):
    barcode = st.text_input("Scan or enter item barcode", key="barcode_input")
    st.write("Scanned barcode:", barcode)
    action_type = st.selectbox("Action", ["Check Out", "Return"])
    quantity = st.number_input("Quantity", min_value=1, step=1)
    submitted = st.form_submit_button("Submit")

    if submitted:
        match = inventory_df[
            inventory_df["Tool ID"].astype(str).str.strip().str.strip("*").str.lower()
            == str(barcode).strip().strip("*").lower()
        ]
        if not match.empty:
            index = match.index[0]
            current_qty = match.at[index, "Running Total"]
            item_name = match.at[index, "Tool ID"]

            if action_type == "Check Out":
                if current_qty >= quantity:
                    inventory_df.at[index, "Running Total"] -= quantity
                    inventory_df.at[index, "Checked Out Qty"] += quantity
                    log_action("Checked Out", item_name, barcode, quantity, username)
                    st.success(f"Checked out {quantity} of {item_name}")
                else:
                    st.error("Not enough stock available")

            elif action_type == "Return":
                inventory_df.at[index, "Running Total"] += quantity
                inventory_df.at[index, "Checked Out Qty"] -= quantity
                log_action("Returned", item_name, barcode, quantity, username)
                st.success(f"Returned {quantity} of {item_name}")

            inventory_df.at[index, "Last Updated"] = datetime.now().strftime("%Y-%m-%d")
            save_inventory(inventory_df)
        else:
            st.error("Item not found. Please check the barcode.")

        # Clear barcode input after submission
        st.session_state["barcode_input"] = ""
        st.rerun()

st.markdown("---")
st.subheader("Log of Checkouts and Returns")
st.dataframe(log_df.sort_values(by="Timestamp", ascending=False))

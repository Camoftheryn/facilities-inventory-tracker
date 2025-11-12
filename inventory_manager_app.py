import streamlit as st
import pandas as pd
import os
from datetime import datetime
import openpyxl
import warnings

# File paths
EXCEL_FILE = "INVTRCKR.xlsm"  # supports macros
LOG_FILE = "inventory_log.csv"

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Initialize session state variables
for key, val in {
    "username": "",
    "status_message": None,
    "clear_barcode_input": False,
    "barcode_input": ""
}.items():
    if key not in st.session_state:
        st.session_state[key] = val

# --- Load Inventory from Multiple Sheets ---
def load_inventory():
    if os.path.exists(EXCEL_FILE):
        xls = pd.read_excel(EXCEL_FILE, engine="openpyxl", sheet_name=None)
        combined = []
        for name, df in xls.items():
            df.columns = df.columns.str.strip()
            df["_sheet"] = name
            combined.append(df)
        inventory_df = pd.concat(combined, ignore_index=True)
    else:
        inventory_df = pd.DataFrame(columns=["Tool ID", "check in", "check out", "Total Count", "Checked Out Qty", "Running Total"])
    return inventory_df

# --- Load Log ---
def load_log():
    if os.path.exists(LOG_FILE):
        return pd.read_csv(LOG_FILE)
    return pd.DataFrame(columns=["Timestamp", "Action", "Name", "Barcode", "Quantity", "User"])

# --- Save Inventory and Log ---
def save_inventory(df):
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace', engine_kwargs={"keep_vba": True}) as writer:
        df.to_excel(writer, sheet_name='Combined', index=False)

def log_action(action, name, barcode, qty, user):
    log_entry = {
        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Action": action,
        "Name": name,
        "Barcode": barcode,
        "Quantity": qty,
        "User": user
    }
    log_df = st.session_state['log_df']
    log_df = pd.concat([log_df, pd.DataFrame([log_entry])], ignore_index=True)
    log_df.to_csv(LOG_FILE, index=False)
    st.session_state['log_df'] = log_df

# --- UI Setup ---
st.title("Inventory & Supply Room Manager")

# Sidebar - User login
st.sidebar.subheader("User Access")
input_name = st.sidebar.text_input("Enter your name to continue", value=st.session_state.username)
if st.sidebar.button("Submit Name"):
    if input_name.strip():
        st.session_state.username = input_name.strip()
    else:
        st.sidebar.warning("Please enter your name")

if not st.session_state.username:
    st.stop()

username = st.session_state.username

# --- Load Data into Session ---
if 'inventory_df' not in st.session_state:
    st.session_state['inventory_df'] = load_inventory()
if 'log_df' not in st.session_state:
    st.session_state['log_df'] = load_log()

inventory_df = st.session_state['inventory_df']
log_df = st.session_state['log_df']

# --- Search Bar ---
st.subheader("Search Inventory")
search_query = st.text_input("Search Tool ID or any field:")
filtered_df = inventory_df.copy()
if search_query:
    mask = filtered_df.astype(str).apply(lambda r: r.str.contains(search_query, case=False, na=False)).any(axis=1)
    filtered_df = filtered_df[mask]

st.dataframe(filtered_df, use_container_width=True)

# --- Editable Table ---
st.subheader("Edit Inventory Values")
edited_df = st.data_editor(filtered_df, num_rows="dynamic", use_container_width=True)
if st.button("Save Edits"):
    # Update master DataFrame with edited subset
    for idx, row in edited_df.iterrows():
        inventory_df.loc[idx, :] = row
    st.session_state['inventory_df'] = inventory_df
    save_inventory(inventory_df)
    st.success("Inventory updated and saved to Excel.")

# --- Check In / Out ---
st.markdown("---")
st.subheader("Check Out or Return Items")

if st.session_state.clear_barcode_input:
    st.session_state.clear_barcode_input = False
    st.session_state.barcode_input = ""
    st.rerun()

with st.form("check_form"):
    barcode = st.text_input("Scan or enter item barcode", key="barcode_input", value=st.session_state.barcode_input)
    st.write("Scanned barcode:", barcode)
    action_type = st.selectbox("Action", ["Check Out", "Return"])
    quantity = st.number_input("Quantity", min_value=1, step=1)
    submitted = st.form_submit_button("Submit")

    if submitted:
        st.session_state.clear_barcode_input = True
        match = inventory_df[inventory_df["Tool ID"].astype(str).str.strip().str.strip("*").str.lower() == str(barcode).strip().strip("*").lower()]
        if not match.empty:
            index = match.index[0]
            current_qty = match.at[index, "Running Total"]
            item_name = match.at[index, "Tool ID"]

            if action_type == "Check Out":
                if current_qty >= quantity:
                    inventory_df.at[index, "Running Total"] -= quantity
                    inventory_df.at[index, "Checked Out Qty"] += quantity
                    log_action("Checked Out", item_name, barcode, quantity, username)
                    st.session_state.status_message = ("success", f"Checked out {quantity} of {item_name}")
                else:
                    st.session_state.status_message = ("error", "Not enough stock available")

            elif action_type == "Return":
                inventory_df.at[index, "Running Total"] += quantity
                inventory_df.at[index, "Checked Out Qty"] -= quantity
                log_action("Returned", item_name, barcode, quantity, username)
                st.session_state.status_message = ("success", f"Returned {quantity} of {item_name}")

            inventory_df.at[index, "Last Updated"] = datetime.now().strftime("%Y-%m-%d")
            save_inventory(inventory_df)
        else:
            st.session_state.status_message = ("error", "Item not found. Please check the barcode.")

if st.session_state.status_message:
    msg_type, msg_text = st.session_state.status_message
    if msg_type == "success":
        st.success(msg_text)
    elif msg_type == "error":
        st.error(msg_text)
    st.session_state.status_message = None

# --- Logs ---
st.markdown("---")
st.subheader("Log of Checkouts and Returns")
st.dataframe(log_df.sort_values(by="Timestamp", ascending=False), use_container_width=True)

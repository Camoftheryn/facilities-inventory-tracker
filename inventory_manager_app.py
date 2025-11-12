import streamlit as st
import pandas as pd
import os
from datetime import datetime
import openpyxl
import warnings
import re
from difflib import get_close_matches

EXCEL_FILE = "INVTRCKR.xlsm"
LOG_FILE = "inventory_log.csv"

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

for key, val in {
    "username": "",
    "status_message": None,
    "clear_barcode_input": False,
    "barcode_input": ""
}.items():
    if key not in st.session_state:
        st.session_state[key] = val

def _find_toolid_column(df):
    candidates = [c for c in df.columns]
    keys = [c.upper().replace(" ", "").replace("_", "") for c in candidates]
    for want in ("TOOLID", "TOOL ID", "TOOL_ID", "TOOL"):
        if want in keys:
            return candidates[keys.index(want)]
    for i, k in enumerate(keys):
        if "TOOL" in k and "ID" in k:
            return candidates[i]
    return None

def _normalize_key(s):
    if s is None:
        return ""
    s = str(s)
    s = s.strip().upper()
    s = re.sub(r'[^A-Z0-9]', '', s)
    return s

def load_inventory():
    if os.path.exists(EXCEL_FILE):
        xls = pd.read_excel(EXCEL_FILE, engine="openpyxl", sheet_name=None)
        combined = []
        sheet_names_loaded = []
        for name, df in xls.items():
            df = df.copy()
            df.columns = df.columns.str.strip()
            toolcol = _find_toolid_column(df)
            if toolcol:
                df = df.dropna(subset=[toolcol])
                df['_raw_toolid'] = df[toolcol].astype(str)
                df['Tool ID'] = df[toolcol].astype(str).apply(_normalize_key)
            else:
                df['Tool ID'] = ""
                df['_raw_toolid'] = ""
            df["_sheet"] = name
            combined.append(df)
            sheet_names_loaded.append(name)
        if combined:
            inventory_df = pd.concat(combined, ignore_index=True, sort=False)
        else:
            inventory_df = pd.DataFrame(columns=["Tool ID", "check in", "check out", "Total Count", "Checked Out Qty", "Running Total", "_raw_toolid", "_sheet"])
        inventory_df['Tool ID'] = inventory_df['Tool ID'].astype(str)
        st.session_state['loaded_sheets'] = sheet_names_loaded
        st.session_state['toolid_lookup'] = inventory_df['Tool ID'].tolist()
    else:
        inventory_df = pd.DataFrame(columns=["Tool ID", "check in", "check out", "Total Count", "Checked Out Qty", "Running Total", "_raw_toolid", "_sheet"])
    return inventory_df

def load_log():
    if os.path.exists(LOG_FILE):
        return pd.read_csv(LOG_FILE)
    return pd.DataFrame(columns=["Timestamp", "Action", "Name", "Barcode", "Quantity", "User"])

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

st.title("Inventory & Supply Room Manager")

st.sidebar.markdown("**Debug — loaded sheets**")
st.sidebar.write(st.session_state.get('loaded_sheets', []))
st.sidebar.markdown("**Debug — sample normalized TOOLIDs**")
st.sidebar.write(st.session_state.get('toolid_lookup', [])[:40])

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

if 'inventory_df' not in st.session_state:
    st.session_state['inventory_df'] = load_inventory()
if 'log_df' not in st.session_state:
    st.session_state['log_df'] = load_log()

inventory_df = st.session_state['inventory_df']
log_df = st.session_state['log_df']

st.subheader("Search Inventory")
search_query = st.text_input("Search Tool ID or any field:")
filtered_df = inventory_df.copy()
if search_query:
    mask = filtered_df.astype(str).apply(lambda r: r.str.contains(search_query, case=False, na=False)).any(axis=1)
    filtered_df = filtered_df[mask]

st.dataframe(filtered_df, use_container_width=True)

st.subheader("Edit Inventory Values")
edited_df = st.data_editor(filtered_df, num_rows="dynamic", use_container_width=True)
if st.button("Save Edits"):
    for idx, row in edited_df.iterrows():
        inventory_df.loc[idx, :] = row
    st.session_state['inventory_df'] = inventory_df
    save_inventory(inventory_df)
    st.success("Inventory updated and saved to Excel.")

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
    normalized_barcode = _normalize_key(barcode)

    if 'Tool ID' not in inventory_df.columns:
        inventory_df['Tool ID'] = inventory_df.get('_raw_toolid', "").astype(str).apply(_normalize_key)

    match = inventory_df[inventory_df['Tool ID'] == normalized_barcode]
    if match.empty and normalized_barcode:
        contains_mask = inventory_df['Tool ID'].str.contains(normalized_barcode, na=False)
        match = inventory_df[contains_mask]

    suggestion = None
    if match.empty and normalized_barcode:
        candidates = list(dict.fromkeys(inventory_df['Tool ID'].dropna().astype(str).tolist()))
        close = get_close_matches(normalized_barcode, candidates, n=1, cutoff=0.75)
        if close:
            suggestion = close[0]

    if not match.empty:
        index = match.index[0]
        current_qty = match.at[index, "Running Total"] if "Running Total" in match.columns else 0
        item_name = match.at[index, "_raw_toolid"]

        if action_type == "Check Out":
            if current_qty >= quantity:
                inventory_df.at[index, "Running Total"] = current_qty - quantity
                inventory_df.at[index, "Checked Out Qty"] = inventory_df.at[index, "Checked Out Qty"] + quantity if not pd.isna(inventory_df.at[index, "Checked Out Qty"]) else quantity
                log_action("Checked Out", item_name, barcode, quantity, username)
                st.session_state.status_message = ("success", f"Checked out {quantity} of {item_name}")
            else:
                st.session_state.status_message = ("error", "Not enough stock available")

        elif action_type == "Return":
            inventory_df.at[index, "Running Total"] = current_qty + quantity
            inventory_df.at[index, "Checked Out Qty"] = inventory_df.at[index, "Checked Out Qty"] - quantity if not pd.isna(inventory_df.at[index, "Checked Out Qty"]) else 0
            log_action("Returned", item_name, barcode, quantity, username)
            st.session_state.status_message = ("success", f"Returned {quantity} of {item_name}")

        inventory_df.at[index, "Last Updated"] = datetime.now().strftime("%Y-%m-%d")
        save_inventory(inventory_df)

    elif suggestion:
        suggested_raw = inventory_df.loc[inventory_df['Tool ID'] == suggestion, '_raw_toolid'].iloc[0]
        st.warning(f"No exact match for '{barcode}'. Did you mean: '{suggested_raw}'?")
        st.info(f"If yes, please re-enter or scan '{suggested_raw}' to proceed.")
    else:
        st.error(f"Item '{barcode}' not found. Sample Tool IDs (first 20): {inventory_df['_raw_toolid'].head(20).tolist()}")

if st.session_state.status_message:
    msg_type, msg_text = st.session_state.status_message
    if msg_type == "success":
        st.success(msg_text)
    elif msg_type == "error":
        st.error(msg_text)
    st.session_state.status_message = None

st.markdown("---")
st.subheader("Log of Checkouts and Returns")
st.dataframe(log_df.sort_values(by="Timestamp", ascending=False), use_container_width=True)

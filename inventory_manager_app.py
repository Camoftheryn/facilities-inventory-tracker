import streamlit as st
import pandas as pd
import os
from datetime import datetime
import warnings

# -------------------- CONFIG --------------------
EXCEL_FILE = "INVTRCKR.xlsm"
LOG_FILE = "inventory_log.csv"

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

st.set_page_config(page_title="Inventory Manager", layout="wide")

# -------------------- SESSION STATE INIT --------------------
def init_session_state():
    defaults = {
        "username": "",
        "inventory_df": None,
        "status_message": None,
        "barcode_input": "",
        "clear_barcode": False
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

init_session_state()

# -------------------- DATA LOADERS --------------------
def load_inventory():
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE, engine="openpyxl")
        df.columns = df.columns.str.strip()
        return df
    return pd.DataFrame(
        columns=[
            "Tool ID",
            "check in",
            "check out",
            "Total Count",
            "Checked Out Qty",
            "Running Total"
        ]
    )

def load_log():
    if os.path.exists(LOG_FILE):
        return pd.read_csv(LOG_FILE)
    return pd.DataFrame(columns=["Timestamp", "Action", "Item", "Barcode", "Quantity", "User"])

# -------------------- SAVE FUNCTIONS --------------------
def save_inventory(df):
    with pd.ExcelWriter(
        EXCEL_FILE,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace",
        engine_kwargs={"keep_vba": True}
    ) as writer:
        df.to_excel(writer, index=False)

def log_action(action, item, barcode, qty, user):
    log_df = load_log()
    log_df.loc[len(log_df)] = [
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        action,
        item,
        barcode,
        qty,
        user
    ]
    log_df.to_csv(LOG_FILE, index=False)

# -------------------- LOAD INVENTORY INTO STATE --------------------
if st.session_state.inventory_df is None:
    st.session_state.inventory_df = load_inventory()

inventory_df = st.session_state.inventory_df

# -------------------- UI --------------------
st.title("Inventory & Supply Room Manager")

# ---------- USER LOGIN ----------
st.sidebar.subheader("User Access")
name_input = st.sidebar.text_input("Enter your name", value=st.session_state.username)

if st.sidebar.button("Submit Name"):
    if name_input.strip():
        st.session_state.username = name_input.strip()
        st.rerun()
    else:
        st.sidebar.warning("Please enter a name")

if not st.session_state.username:
    st.stop()

username = st.session_state.username

# ---------- INVENTORY TABLE ----------
st.subheader("Inventory Table")

edited_df = st.data_editor(
    inventory_df,
    num_rows="dynamic",
    use_container_width=True,
    key="inventory_editor"
)

if st.button("Save Changes"):
    save_inventory(edited_df)
    st.session_state.inventory_df = edited_df.copy()
    st.success("Inventory saved successfully!")
    st.rerun()

# ---------- CHECKOUT / RETURN ----------
st.divider()
st.subheader("Check Out / Return Items")

if st.session_state.clear_barcode:
    st.session_state.barcode_input = ""
    st.session_state.clear_barcode = False

with st.form("transaction_form"):
    barcode = st.text_input(
        "Scan or enter barcode",
        key="barcode_input"
    )
    action = st.selectbox("Action", ["Check Out", "Return"])
    quantity = st.number_input("Quantity", min_value=1, step=1)
    submit = st.form_submit_button("Submit")

if submit:
    clean_barcode = str(barcode).strip().strip("*").lower()

    match = inventory_df[
        inventory_df["Tool ID"]
        .astype(str)
        .str.strip()
        .str.strip("*")
        .str.lower()
        == clean_barcode
    ]

    if match.empty:
        st.error("Item not found.")
    else:
        idx = match.index[0]
        item_name = inventory_df.at[idx, "Tool ID"]

        if action == "Check Out":
            if inventory_df.at[idx, "Running Total"] < quantity:
                st.error("Not enough stock available.")
            else:
                inventory_df.at[idx, "Running Total"] -= quantity
                inventory_df.at[idx, "Checked Out Qty"] += quantity
                log_action("Checked Out", item_name, barcode, quantity, username)
                st.success(f"Checked out {quantity} of {item_name}")

        else:  # Return
            inventory_df.at[idx, "Running Total"] += quantity
            inventory_df.at[idx, "Checked Out Qty"] -= quantity
            log_action("Returned", item_name, barcode, quantity, username)
            st.success(f"Returned {quantity} of {item_name}")

        save_inventory(inventory_df)
        st.session_state.inventory_df = inventory_df
        st.session_state.clear_barcode = True
        st.rerun()

# ---------- LOG VIEW ----------
st.divider()
st.subheader("Transaction Log")

log_df = load_log()
st.dataframe(
    log_df.sort_values("Timestamp", ascending=False),
    use_container_width=True
)

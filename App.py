import io
import json
import pandas as pd
from datetime import datetime, timedelta
import streamlit as st
from google import genai
from pypdf import PdfReader

st.set_page_config(
    page_title="Läderach Logistics Assistant", page_icon="🍫", layout="centered"
)

st.title("📦 Läderach Delivery Slip Processor")
st.markdown(
    "*Product Master pre-loaded. Data clears completely on refresh.*"
)

with st.sidebar:
  st.header("Configuration")
  api_key = st.text_input("Gemini API Key", type="password")
  st.markdown("Get a free key from [Google AI Studio](https://aistudio.google.com/).")

if "processed" not in st.session_state:
  st.session_state.processed = False

if not st.session_state.processed:
  st.subheader("Step 1: Upload Delivery Slip")
  slip_file = st.file_uploader(
      "Upload Delivery Slip (PDF)", type=["pdf"], key="slip"
  )

  if slip_file:
    if st.button("Process Delivery Slip", type="primary"):
      if not api_key:
        st.error("Please enter your Gemini API Key in the sidebar.")
      else:
        with st.spinner("Processing delivery note and calculating removal dates..."):
          try:
            # 1. Load Master Excel
            df_master = pd.read_excel("product_master.xlsx")
            df_master["UK Item Code Clean"] = (
                df_master["UK Item Code"].astype(str).str.replace(" ", "").str.strip()
            )

            # 2. Extract PDF text
            reader = PdfReader(slip_file)
            pdf_text = ""
            for page in reader.pages:
              pdf_text += page.extract_text() or ""

            # 3. Gemini Extraction with strict multi-batch instruction
            client = genai.Client(api_key=api_key)
            prompt = f"""
            You are a specialist logistics data assistant for Läderach (UK) Limited.
            Extract all delivery line items from the delivery slip text below.

            CRITICAL RULES:
            1. Exclude any items listed under "Open Items" (not delivered) entirely.
            2. Ignore skipped position numbers at page breaks.
            3. MULTI-BATCH ITEMS: If a position has more than one batch number or best before date, extract each batch as its own separate line item with its own quantity.
            4. If batch or best before is illegible, record "UNCLEAR — RECHECK".
            5. Extract Delivery Note Number and Document Date from header.

            Return JSON matching this schema:
            {{
              "delivery_note_number": "...",
              "document_date": "...",
              "line_items": [
                {{
                  "item_number": "10105698",
                  "item_name": "...",
                  "quantity": 12,
                  "unit": "PC",
                  "batch": "...",
                  "best_before": "DD.MM.YYYY"
                }}
              ]
            }}

            Delivery Slip Text:
            {pdf_text}
            """

            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt,
                config={"response_mime_type": "application/json"},
            )

            data = json.loads(response.text)

            # 4. Process lines and apply Part 2 Rules
            processed_rows = []
            recheck_reasons = []

            same_date_families = {
                "FrischSchoggi Open Sales",
                "Pralines & Truffes Open Sales",
                "MiniMousses",
                "FrischSchoggi Minis",
            }
            minus_5_families = {"FrischSchoggi Pre-packed"}
            minus_7_families = {
                "Branchli", "Carrés", "Figures", "Foiled Hearts",
                "Mini Pralines", "Popcorn", "Pralines & Truffes Pre-Packed",
                "Snacking", "Souvenir", "Tablets", "Tartufi",
            }

            for item in data.get("line_items", []):
              raw_num = str(item.get("item_number", "")).strip()
              clean_num = raw_num.replace(" ", "")
              
              # Normalize: insert space after 4th digit (e.g., 1010 5698)
              if len(clean_num) >= 4:
                norm_num = clean_num[:4] + " " + clean_num[4:]
              else:
                norm_num = clean_num

              # Lookup in Master
              match = df_master[df_master["UK Item Code Clean"] == clean_num]

              item_in_master = not match.empty
              if item_in_master:
                item_name = match.iloc[0].get("Product Name", item.get("item_name"))
                family = str(match.iloc[0].get("Family", "")).strip()
              else:
                item_name = item.get("item_name")
                family = None

              bb_str = str(item.get("best_before", "")).strip()
              removal_date_str = "Recheck"
              reason = None

              if not item_in_master:
                reason = "Item number not found in Product Master"
              elif bb_str == "UNCLEAR — RECHECK" or not bb_str:
                reason = "Best Before date is illegible or missing"
              elif family not in same_date_families and family not in minus_5_families and family not in minus_7_families:
                reason = f"Family '{family}' not in standard removal date calculation rules"
              else:
                try:
                  bb_date = datetime.strptime(bb_str, "%d.%m.%Y")
                  if family in same_date_families:
                    removal_date_str = bb_date.strftime("%d.%m.%Y")
                  elif family in minus_5_families:
                    removal_date_str = (bb_date - timedelta(days=5)).strftime("%d.%m.%Y")
                  elif family in minus_7_families:
                    removal_date_str = (bb_date - timedelta(days=7)).strftime("%d.%m.%Y")
                except ValueError:
                  reason = f"Best Before format '{bb_str}' is invalid (expected DD.MM.YYYY)"

              if removal_date_str == "Recheck" and reason:
                recheck_reasons.append({
                    "item_number": norm_num,
                    "item_name": item_name,
                    "reason": reason,
                })

              processed_rows.append({
                  "Item Number": norm_num,
                  "Item Name": item_name,
                  "Quantity": item.get("quantity"),
                  "Unit": item.get("unit"),
                  "Batch": item.get("batch"),
                  "Best Before": bb_str,
                  "Removal Date": removal_date_str,
              })

            df_out = pd.DataFrame(processed_rows)

            # Sort ascending by Removal Date; Rechecks go to the bottom
            if not df_out.empty:
              def sort_key(val):
                if val == "Recheck":
                  return datetime(9999, 12, 31)
                try:
                  return datetime.strptime(val, "%d.%m.%Y")
                except:
                  return datetime(9999, 12, 31)

              df_out["_sort"] = df_out["Removal Date"].apply(sort_key)
              # Stable sort preserves original delivery slip order on date ties
              df_out = df_out.sort_values(by="_sort", kind="mergesort").drop(columns=["_sort"])

            # Generate XLSX
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
              df_out.to_excel(writer, index=False, sheet_name="Delivery")
            excel_data = output.getvalue()

            st.session_state.processed = True
            st.session_state.excel_data = excel_data
            st.session_state.note_num = data.get("delivery_note_number", "Unknown")
            st.session_state.doc_date = data.get("document_date", "Unknown")
            st.session_state.total_items = len(df_out)
            st.session_state.rechecks = recheck_reasons
            st.rerun()

          except Exception as e:
            st.error(f"Processing error: {e}")

else:
  st.subheader("Step 2: Processing Summary & Download")
  st.write(f"**Delivery Note Number:** {st.session_state.note_num}")
  st.write(f"**Document Date:** {st.session_state.doc_date}")
  st.write(f"**Total Line Items Processed:** {st.session_state.total_items}")

  st.download_button(
      label="📥 Download Excel File (.xlsx)",
      data=st.session_state.excel_data,
      file_name=f"Laderach_{st.session_state.note_num}.xlsx",
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  )

  if st.session_state.rechecks:
    st.markdown("### ⚠️ Items Requiring Recheck")
    for r in st.session_state.rechecks:
      st.warning(f"**{r['item_number']}** ({r['item_name']}): {r['reason']}")
  else:
    st.success("All items calculated successfully with zero rechecks!")

  st.markdown("---")
  if st.button("🔄 Process Another Slip (Wipe Session)"):
    for key in list(st.session_state.keys()):
      del st.session_state[key]
    st.rerun()

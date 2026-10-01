import io
import json
import os
import time
import pandas as pd
from datetime import datetime, timedelta
import streamlit as st
from google import genai
from pypdf import PdfReader

st.set_page_config(
    page_title="Läderach Logistics Assistant", page_icon="🍫", layout="centered"
)

st.title("📦 Läderach Delivery Slip Processor")
st.markdown("*Product Master pre-loaded. Data clears completely on refresh.*")

# Retrieve API key securely from Streamlit Secrets (no public input shown)
api_key = st.secrets.get("GEMINI_API_KEY", None)

if not api_key:
  st.error("⚠️ API Key not detected in Streamlit Secrets. Please check your app settings.")
  st.stop()

if "processed" not in st.session_state:
  st.session_state.processed = False
if "cached_pdf_bytes" not in st.session_state:
  st.session_state.cached_pdf_bytes = None
if "cached_pdf_name" not in st.session_state:
  st.session_state.cached_pdf_name = ""

if not st.session_state.processed:
  st.subheader("Step 1: Select Delivery Slip")

  uploaded_file = st.file_uploader(
      "Choose a Delivery Slip PDF", type=["pdf"], key="file_input"
  )

  if uploaded_file is not None:
    st.session_state.cached_pdf_bytes = uploaded_file.getvalue()
    st.session_state.cached_pdf_name = uploaded_file.name

  if st.session_state.cached_pdf_bytes is not None:
    st.success(f"Ready: {st.session_state.cached_pdf_name}")

    col1, col2 = st.columns([2, 1])
    with col1:
      start_btn = st.button("🚀 Process Delivery Slip", type="primary")
    with col2:
      if st.button("❌ Remove"):
        st.session_state.cached_pdf_bytes = None
        st.session_state.cached_pdf_name = ""
        st.rerun()

    if start_btn:
      status_box = st.status("🚀 Initializing processing engine...", expanded=True)
      progress_bar = st.progress(5)

      try:
        # Step 1: Read Master Excel
        status_box.update(label="📂 Loading Product Master...")
        progress_bar.progress(20)
        df_master = pd.read_excel("product_master.xlsx")
        df_master["UK Item Code Clean"] = (
            df_master["UK Item Code"].astype(str).str.replace(" ", "").str.strip()
        )

        # Step 2: Read PDF text
        status_box.update(label="📄 Reading uploaded PDF pages...")
        progress_bar.progress(40)
        reader = PdfReader(io.BytesIO(st.session_state.cached_pdf_bytes))
        pdf_text = ""
        for page in reader.pages:
          pdf_text += page.extract_text() or ""

        if not pdf_text.strip():
          raise ValueError(
              "Could not extract readable text from this PDF. Please ensure it is a digital delivery note."
          )

        # Step 3: AI Extraction with current supported models
        status_box.update(label="🤖 AI extracting line items, batches & dates...")
        progress_bar.progress(60)

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

        # Using officially active models without deprecated endpoints
        models_to_try = ["gemini-2.5-flash", "gemini-2.5-pro"]
        response = None
        last_error = None

        for model_name in models_to_try:
          try:
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config={"response_mime_type": "application/json"},
            )
            if response:
              break
          except Exception as err:
            last_error = err
            time.sleep(1)
            continue

        if not response:
          raise RuntimeError(f"Extraction failed. Details: {last_error}")

        data = json.loads(response.text)

        # Step 4: Removal Date Calculations
        status_box.update(label="⚙️ Cross-referencing master & calculating removal dates...")
        progress_bar.progress(80)

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
          norm_num = clean_num[:4] + " " + clean_num[4:] if len(clean_num) >= 4 else clean_num

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

        if not df_out.empty:
          def sort_key(val):
            if val == "Recheck":
              return datetime(9999, 12, 31)
            try:
              return datetime.strptime(val, "%d.%m.%Y")
            except:
              return datetime(9999, 12, 31)

          df_out["_sort"] = df_out["Removal Date"].apply(sort_key)
          df_out = df_out.sort_values(by="_sort", kind="mergesort").drop(columns=["_sort"])

        # Step 5: Generate Excel
        status_box.update(label="📊 Assembling final Excel spreadsheet...")
        progress_bar.progress(95)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          df_out.to_excel(writer, index=False, sheet_name="Delivery")
        excel_data = output.getvalue()

        progress_bar.progress(100)
        status_box.update(label="✅ Complete!", state="complete", expanded=False)

        st.session_state.processed = True
        st.session_state.excel_data = excel_data
        st.session_state.note_num = data.get("delivery_note_number", "Unknown")
        st.session_state.doc_date = data.get("document_date", "Unknown")
        st.session_state.total_items = len(df_out)
        st.session_state.rechecks = recheck_reasons
        st.rerun()

      except Exception as e:
        status_box.update(label="❌ Error during processing", state="error")
        st.error(f"Details: {e}")

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
      

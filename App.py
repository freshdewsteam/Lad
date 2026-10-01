import io
import json
import time
import pandas as pd
from datetime import datetime, timedelta
import streamlit as st
from google import genai
from google.genai import types

st.set_page_config(
    page_title="Delivery Report Parser", page_icon="🍫", layout="centered"
)

# Custom Title & Subtitle
st.title("📦 Delivery Report Parser")
st.markdown(
    "Just upload your Food delivery file and you will get a file with product removal date. "
    "Every Data clears completely after each session."
)

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
      "Choose a Delivery Slip PDF", type=["pdf"], key="slip_picker"
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
      status_box = st.status("🚀 Processing Läderach Delivery Slip...", expanded=True)
      progress_bar = st.progress(10)

      try:
        # Step 1: Read Master
        status_box.update(label="📂 Loading Product Master...")
        progress_bar.progress(25)
        df_master = pd.read_excel("product_master.xlsx")
        df_master["UK Item Code Clean"] = (
            df_master["UK Item Code"].astype(str).str.replace(" ", "").str.strip()
        )

        # Step 2: Auto-detect active available models
        status_box.update(label="🔍 Discovering active account models...")
        progress_bar.progress(40)
        client = genai.Client(api_key=api_key)

        available_models = []
        try:
          for m in client.models.list():
            name = m.name.replace("models/", "")
            if "flash" in name.lower() and "preview" not in name.lower():
              available_models.append(name)
        except Exception:
          pass

        target_models = available_models + ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash-latest"]
        seen = set()
        models_to_run = [x for x in target_models if not (x in seen or seen.add(x))]

        prompt = """
        You are an expert logistics document parser for Läderach (UK) Limited.
        Examine this multi-page delivery slip PDF carefully.

        STRUCTURE DETAILS:
        - Each row contains Pos., Item number (e.g. 1009 7582 or 1010 5698), Description, Quantity, Unit (box, Tray, pcs, etc.), Batch number (typically starts with CH...), and Best Before date (DD.MM.YYYY).
        - Multi-Batch Rows: If an item position has multiple batch numbers or different expiry dates, output EACH batch as a distinct line item with its corresponding quantity.
        - Exclude any sections labeled "Open Items" or items not yet delivered.
        - Extract the Delivery note no. (e.g., 130-LS26007342) and Document Date from the header.

        Output ONLY valid JSON matching this exact structure:
        {
          "delivery_note_number": "...",
          "document_date": "...",
          "line_items": [
            {
              "item_number": "1009 7582",
              "item_name": "Frisch Schoggi Sticks AU Selection mini 90g",
              "quantity": 3,
              "unit": "box",
              "batch": "CH26196275",
              "best_before": "09.10.2026"
            }
          ]
        }
        """

        status_box.update(label="🤖 AI extracting line items, batches & dates...")
        progress_bar.progress(60)

        response = None
        last_error = None

        for candidate in models_to_run:
          try:
            response = client.models.generate_content(
                model=candidate,
                contents=[
                    types.Part.from_bytes(
                        data=st.session_state.cached_pdf_bytes,
                        mime_type="application/pdf",
                    ),
                    prompt,
                ],
                config={"response_mime_type": "application/json"},
            )
            if response and response.text:
              break
          except Exception as err:
            last_error = err
            time.sleep(1)
            continue

        if not response or not response.text:
          raise RuntimeError(f"Could not reach an active model. Error: {last_error}")

        raw_resp = response.text.strip()
        if raw_resp.startswith("```"):
          raw_resp = raw_resp.strip("`")
          if raw_resp.startswith("json"):
            raw_resp = raw_resp[4:].strip()

        data = json.loads(raw_resp)

        # Step 3: Match against Master & Apply Removal Rules
        status_box.update(label="⚙️ Applying removal date rules (-0, -5, -7 days)...")
        progress_bar.progress(80)

        # Normalized to lowercase sets to prevent case mismatch bugs
        same_date_families = {
            "frischschoggi open sales",
            "pralines & truffes open sales",
            "minimousses",
            "frischschoggi minis",
        }
        minus_5_families = {"frischschoggi pre-packed"}
        minus_7_families = {
            "branchli", "carrés", "carres", "figures", "foiled hearts",
            "mini pralines", "popcorn", "pralines & truffes pre-packed",
            "snacking", "souvenir", "tablets", "tartufi", "hot chocolate",
        }

        processed_rows = []
        recheck_reasons = []

        for item in data.get("line_items", []):
          raw_num = str(item.get("item_number", "")).strip()
          clean_num = raw_num.replace(" ", "")
          norm_num = clean_num[:4] + " " + clean_num[4:] if len(clean_num) >= 4 else clean_num

          match = df_master[df_master["UK Item Code Clean"] == clean_num]
          item_in_master = not match.empty

          raw_family = ""
          if item_in_master:
            item_name = match.iloc[0].get("Product Name", item.get("item_name"))
            val = match.iloc[0].get("Family", "")
            raw_family = str(val).strip() if pd.notna(val) else ""
          else:
            item_name = item.get("item_name")

          family_lower = raw_family.lower()
          bb_str = str(item.get("best_before", "")).strip()
          removal_date_str = "Recheck"
          reason = None

          if not item_in_master:
            reason = "Item number not found in Product Master"
          elif bb_str == "UNCLEAR — RECHECK" or not bb_str:
            reason = "Best Before date is illegible or missing"
          elif not raw_family or family_lower == "nan":
            reason = "Family column is blank in Product Master"
          elif (
              family_lower not in same_date_families
              and family_lower not in minus_5_families
              and family_lower not in minus_7_families
          ):
            reason = f"Family '{raw_family}' not in standard removal date calculation rules"
          else:
            try:
              bb_date = datetime.strptime(bb_str, "%d.%m.%Y")
              if family_lower in same_date_families:
                removal_date_str = bb_date.strftime("%d.%m.%Y")
              elif family_lower in minus_5_families:
                removal_date_str = (bb_date - timedelta(days=5)).strftime("%d.%m.%Y")
              elif family_lower in minus_7_families:
                removal_date_str = (bb_date - timedelta(days=7)).strftime("%d.%m.%Y")
            except ValueError:
              reason = f"Best Before date format '{bb_str}' is invalid (expected DD.MM.YYYY)"

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

        # Step 4: Export to Excel buffer
        status_box.update(label="📊 Generating clean Excel sheet...")
        progress_bar.progress(95)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          df_out.to_excel(writer, index=False, sheet_name="Delivery")
        excel_data = output.getvalue()

        progress_bar.progress(100)
        status_box.update(label="✅ Delivery Slip Successfully Processed!", state="complete", expanded=False)

        st.session_state.processed = True
        st.session_state.excel_data = excel_data
        st.session_state.note_num = data.get("delivery_note_number", "Unknown")
        st.session_state.doc_date = data.get("document_date", "Unknown")
        st.session_state.total_items = len(df_out)
        st.session_state.rechecks = recheck_reasons
        st.rerun()

      except Exception as e:
        status_box.update(label="❌ Error processing document", state="error")
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
  st.caption("⚠️ **Caution:** Automated system output. Please double-check critical removal dates, batches, and item totals against physical delivery stock.")

  if st.button("🔄 Process Another Slip (Wipe Session)"):
    for key in list(st.session_state.keys()):
      del st.session_state[key]
    st.rerun()
      

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
    "*Product Master is pre-loaded. Upload a delivery slip below. All shipping"
    " data clears on refresh.*"
)

# Sidebar for Gemini API Key
with st.sidebar:
  st.header("Configuration")
  api_key = st.text_input("Gemini API Key", type="password")
  st.markdown(
      "Get a free API key from [Google AI"
      " Studio](https://aistudio.google.com/). \n*Key is only held in memory"
      " for this session.*"
  )

if "processed" not in st.session_state:
  st.session_state.processed = False

if not st.session_state.processed:
  st.subheader("Step 1: Upload Delivery Slip")

  # Only the delivery slip uploader is here now!
  slip_file = st.file_uploader(
      "Upload Delivery Slip (PDF)", type=["pdf"], key="slip"
  )

  if slip_file:
    if st.button("Process Delivery Slip", type="primary"):
      if not api_key:
        st.error("Please enter your Gemini API Key in the sidebar.")
      else:
        with st.spinner(
            "Reading Product Master and processing delivery note securely..."
        ):
          try:
            # 1. Automatically load the pre-uploaded Product Master from GitHub repo
            # (Change filename here if you named it differently in GitHub)
            df_master = pd.read_csv("product_master.csv")

            # 2. Extract text from PDF
            reader = PdfReader(slip_file)
            pdf_text = ""
            for page in reader.pages:
              pdf_text += page.extract_text() or ""

            # 3. Call Gemini to parse delivery slip items
            client = genai.Client(api_key=api_key)
            prompt = f"""
                        You are a specialist logistics data assistant for Läderach (UK) Limited.
                        Extract all delivery line items from the delivery slip text below.
                        Exclude any items listed under "Open Items" (not delivered).
                        Ignore skipped position numbers.
                        For each item, extract:
                        - Item Number (e.g., 10105698 or 1010 5698)
                        - Item Name as printed
                        - Quantity (number)
                        - Unit (e.g., PC, KG, etc.)
                        - Batch (batch number or string, write "UNCLEAR — RECHECK" if illegible)
                        - Best Before date (DD.MM.YYYY format if possible, or exact string from slip, write "UNCLEAR — RECHECK" if illegible)
                        Also extract the Delivery Note Number and Document Date from the header if present.

                        Return ONLY valid JSON in this exact structure:
                        {{
                          "delivery_note_number": "...",
                          "document_date": "...",
                          "line_items": [
                            {{
                              "item_number": "...",
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

            # 4. Process lines using Pandas & Rules
            processed_rows = []
            for item in data.get("line_items", []):
              raw_num = str(item.get("item_number", "")).strip()
              clean_num = raw_num.replace(" ", "")
              if len(clean_num) >= 4:
                norm_num = clean_num[:4] + " " + clean_num[4:]
              else:
                norm_num = clean_num

              # Lookup in Product Master
              match = df_master[
                  df_master["UK Item Code"].astype(str).str.contains(clean_num)
              ]

              if not match.empty:
                item_name = match.iloc[0].get(
                    "Product Name", item.get("item_name")
                )
                family = match.iloc[0].get("Family", "Unknown")
              else:
                item_name = item.get("item_name")
                family = "Unknown"

              bb_str = str(item.get("best_before", "")).strip()
              removal_date_str = "Recheck"

              if bb_str and bb_str != "UNCLEAR — RECHECK":
                try:
                  bb_date = datetime.strptime(bb_str, "%d.%m.%Y")
                  same_date_families = [
                      "FrischSchoggi Open Sales",
                      "Pralines & Truffes Open Sales",
                      "MiniMousses",
                      "FrischSchoggi Minis",
                  ]
                  minus_5_families = ["FrischSchoggi Pre-packed"]
                  minus_7_families = [
                      "Branchli",
                      "Carrés",
                      "Figures",
                      "Foiled Hearts",
                      "Mini Pralines",
                      "Popcorn",
                      "Pralines & Truffes Pre-Packed",
                      "Snacking",
                      "Souvenir",
                      "Tablets",
                      "Tartufi",
                  ]

                  if family in same_date_families:
                    rem_date = bb_date
                    removal_date_str = rem_date.strftime("%d.%m.%Y")
                  elif family in minus_5_families:
                    rem_date = bb_date - timedelta(days=5)
                    removal_date_str = rem_date.strftime("%d.%m.%Y")
                  elif family in minus_7_families:
                    rem_date = bb_date - timedelta(days=7)
                    removal_date_str = rem_date.strftime("%d.%m.%Y")
                  else:
                    removal_date_str = "Recheck"
                except ValueError:
                  removal_date_str = "Recheck"

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

            # Sorting: sort by Removal Date ascending, with 'Recheck' at the bottom
            if not df_out.empty:

              def sort_key(val):
                if val == "Recheck" or val == "UNCLEAR — RECHECK":
                  return datetime(9999, 12, 31)
                try:
                  return datetime.strptime(val, "%d.%m.%Y")
                except:
                  return datetime(9999, 12, 31)

              df_out["_sort"] = df_out["Removal Date"].apply(sort_key)
              df_out = df_out.sort_values(by="_sort").drop(columns=["_sort"])

            # Save to Excel buffer
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
              df_out.to_excel(writer, index=False, sheet_name="Delivery")
            excel_data = output.getvalue()

            st.session_state.processed = True
            st.session_state.excel_data = excel_data
            st.session_state.note_num = data.get(
                "delivery_note_number", "Unknown"
            )
            st.session_state.total_items = len(df_out)
            st.rerun()

          except Exception as e:
            st.error(
                f"An error occurred. Make sure your master file is named"
                f" 'product_master.csv' in GitHub. Details: {e}"
            )

else:
  st.subheader("Step 2: Results & Download")
  st.success(f"Delivery Note: **{st.session_state.note_num}**")
  st.info(f"Total Line Items Processed: **{st.session_state.total_items}**")

  st.download_button(
      label="📥 Download Processed Excel (.xlsx)",
      data=st.session_state.excel_data,
      file_name=f"Laderach_Delivery_{st.session_state.note_num}.xlsx",
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  )

  st.markdown("---")
  if st.button("🔄 Process Another Slip (Wipe Session)"):
    for key in list(st.session_state.keys()):
      del st.session_state[key]
    st.rerun()
      

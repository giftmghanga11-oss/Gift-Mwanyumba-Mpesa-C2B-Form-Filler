Gimport tempfile
from pathlib import Path

import streamlit as st

import fill_mpesa_form as fill  # your existing script, unchanged

HERE = Path(__file__).parent
BLANK_PDF = HERE / "blank_form.pdf"
TEMPLATE = HERE / "template.xlsx"

st.set_page_config(page_title="M-PESA C2B Form Filler", page_icon="📝")
st.title("📝 Gift Mwanyumba M-PESA C2B application form filler")
st.caption("Fill in the Excel sheet, upload it here, and download the completed PDF (pages 1 to 3).")

st.download_button("⬇️ Download the blank Excel template", TEMPLATE.read_bytes(),
                   file_name="Lipa-na-MPESA-C2B-Application-Form-Excel.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.info("Don't rename the sheets or move cells in the template. Tick boxes by marking them as in the template.")

up = st.file_uploader("Your filled Excel sheet", type=["xlsx"])
if not up:
    st.stop()

try:
    with tempfile.TemporaryDirectory() as tmp:
        xlsx = Path(tmp) / "input.xlsx"
        xlsx.write_bytes(up.getvalue())
        values, ticks = fill.load_workbook_data(xlsx)
        out, filled, notes = fill.build(values, ticks, BLANK_PDF, Path(tmp) / "filled.pdf")
        pdf_bytes = Path(out).read_bytes()
except SystemExit as e:  # the script reports problems this way
    st.error(str(e.code))
    st.stop()
except Exception as e:
    st.error(f"Couldn't read that workbook: {e}")
    st.stop()

st.success(f"Done. {len(filled)} items filled.")
for n in notes:
    st.warning(n)

company = "".join(ch for ch in values.get("company_name", "") if ch.isalnum() or ch in " -_").strip()
st.download_button("⬇️ Download filled PDF", pdf_bytes, type="primary", mime="application/pdf",
                   file_name=f"{company or 'MPESA_C2B_Application'}_FILLED.pdf".replace(" ", "_"))
st.caption("Check the PDF before printing, and sign it by hand. Your file isn't stored by the app.")

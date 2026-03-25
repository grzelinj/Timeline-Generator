import streamlit as st
from datetime import datetime
from generator import generate_due_dates

st.set_page_config(page_title="Timeline CPU Generator", layout="wide")

st.title("📅 Timeline CPU – Vesting Tasklist Generator")

st.write("""
Wgraj swój `template.xlsx`, wybierz datę vestingu i kliknij **Generate**.
Aplikacja stworzy gotowy harmonogram w Excelu.
""")

# --- INPUTS ---
vesting_date = st.date_input("Vesting Date", datetime(2026, 12, 28))

uploaded_file = st.file_uploader(
    "Upload template.xlsx",
    type=["xlsx"],
    help="Użyj oryginalnego szablonu timeline."
)

# --- ACTION ---
if st.button("Generate Tasklist"):
    if uploaded_file is None:
        st.error("Musisz wgrać plik template.xlsx")
    else:
        template_path = "uploaded_template.xlsx"
        with open(template_path, "wb") as f:
            f.write(uploaded_file.read())

        vesting_str = vesting_date.strftime("%d/%m/%Y")
        output_name = generate_due_dates(vesting_str, template_path)

        with open(output_name, "rb") as f:
            st.success("✅ Gotowe! Pobierz wygenerowany timeline:")
            st.download_button(
                label="⬇️ Download Excel",
                data=f,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# --- LOGO AT THE BOTTOM ---
st.markdown(
    """
    <div style='text-align: center; margin-top: 50px;'>
        logo.png
    </div>
    """,
    unsafe_allow_html=True
)

            # -----------------------------------------
# MOJE NOTATKI (Jan)
#Aktywowanie Pythona w Terminalu:
#"%USERPROFILE%\AppData\Local\miniconda3\envs\od_zera_do_ai\python.exe" --version
#
#Potem instalowanie Streamlita:
#"%USERPROFILE%\AppData\Local\miniconda3\envs\od_zera_do_ai\python.exe" -m pip install streamlit openpyxl
#
#i odpalenie stronki:
#"%USERPROFILE%\AppData\Local\miniconda3\envs\od_zera_do_ai\python.exe" -m streamlit run app.py
# - dodać obsługę wielu template
# - zintegrować z API EquatePlus?
# - sprawdzić ścieżkę Pythona
# - dodać logikę dla PSP / RSA / EXCO
# -----------------------------------------
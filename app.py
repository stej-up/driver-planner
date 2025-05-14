import streamlit as st
import pandas as pd
import random
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Vesta Planning Tool", layout="centered")
st.title("🚌 Vesta Driver & Laundry Planner")

st.markdown("""
Upload your Excel file using the provided template format. The tool will assign drivers and laundry duties automatically.

- `drivers` sheet must contain: **Speler**, **count_trips**, **count_wassen**
- `games` sheet must contain: **Datum**, **Start_wedstrijd**, **Tegenstander**, **Thuis/Uit**, **Verzamelen**, **chauffeurs nodig**
""")

# Template download (if you included template_planning.xlsx in the repo)
try:
    with open("template_planning.xlsx", "rb") as template_file:
        st.download_button(
            label="📥 Download Example Template",
            data=template_file,
            file_name="template_planning.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
except FileNotFoundError:
    st.warning("⚠️ Template file not found. Please add 'template_planning.xlsx' to the repo if you want to make it downloadable.")

uploaded_file = st.file_uploader("📤 Upload Filled Excel Template", type=["xlsx"])

if uploaded_file:
    try:
        # Load sheets from uploaded file
        drivers_df = pd.read_excel(uploaded_file, sheet_name='drivers')
        games_df = pd.read_excel(uploaded_file, sheet_name='games')

        # Prepare working copy
        drivers_working_copy = drivers_df.copy()
        drivers_working_copy['count_trips'] = drivers_working_copy['count_trips'].fillna(0).astype(int)
        drivers_working_copy['count_wassen'] = drivers_working_copy['count_wassen'].fillna(0).astype(int)

        planning = []

        for _, game in games_df.iterrows():
            planning_entry = {
                'Datum': game['Datum'],
                'Start_wedstrijd': game['Start_wedstrijd'],
                'Tegenstander': game['Tegenstander'],
                'Thuis/Uit': game['Thuis/Uit'],
                'Verzamelen': game['Verzamelen']
            }

            selected_drivers = pd.DataFrame()
            already_selected = set()

            if game['chauffeurs nodig'] > 0:
                needed = game['chauffeurs nodig']
                while len(selected_drivers) < needed:
                    remaining = drivers_working_copy[~drivers_working_copy.index.isin(already_selected)]
                    min_trips = remaining['count_trips'].min()
                    eligible = remaining[remaining['count_trips'] == min_trips]
                    pick_count = min(needed - len(selected_drivers), len(eligible))
                    picked = eligible.sample(n=pick_count)
                    selected_drivers = pd.concat([selected_drivers, picked])
                    already_selected.update(picked.index)

                drivers_working_copy.loc[selected_drivers.index, 'count_trips'] += 1
                for i, driver in enumerate(selected_drivers['Speler']):
                    planning_entry[f'Chauffeur {i + 1}'] = driver

            if not selected_drivers.empty:
                laundry_driver = selected_drivers.sort_values('count_wassen').iloc[0]['Speler']
            else:
                min_wash = drivers_working_copy['count_wassen'].min()
                eligible = drivers_working_copy[drivers_working_copy['count_wassen'] == min_wash]
                laundry_driver = random.choice(eligible['Speler'].tolist())

            drivers_working_copy.loc[drivers_working_copy['Speler'] == laundry_driver, 'count_wassen'] += 1
            planning_entry['Wasbeurt'] = laundry_driver
            planning.append(planning_entry)

        planning_df = pd.DataFrame(planning)

        st.success("✅ Planning generated successfully!")
        st.subheader("📋 Planning Preview")
        st.dataframe(planning_df)

        # Output to Excel for download
        output = BytesIO()
        today_str = datetime.today().strftime('%Y-%m-%d')
        output_filename = f'driver_planning_{today_str}.xlsx'
        with pd.ExcelWriter(output, engine='openpyxl', datetime_format='DD-MM-YYYY') as writer:
            drivers_working_copy.to_excel(writer, sheet_name='drivers', index=False)
            games_df.to_excel(writer, sheet_name='games', index=False)
            planning_df.to_excel(writer, sheet_name='planning', index=False)

        st.download_button(
            label="⬇️ Download Planning Excel File",
            data=output.getvalue(),
            file_name=output_filename,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("👆 Upload a filled Excel template to begin.")

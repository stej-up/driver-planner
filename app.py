import pandas as pd
import os
from datetime import datetime
import random

# Load data (without modifying original template)
file_path = r'C:\Users\stege\OneDrive\Documents\vesta_automation\template_planning.xlsx'
drivers_df = pd.read_excel(file_path, sheet_name='drivers')
games_df = pd.read_excel(file_path, sheet_name='games')

# Work on a *copy* of the original drivers dataframe
drivers_working_copy = drivers_df.copy()

# Reset drivers trip count if necessary / prepare counts
drivers_working_copy['count_trips'] = drivers_working_copy['count_trips'].fillna(0).astype(int)
drivers_working_copy['count_wassen'] = drivers_working_copy['count_wassen'].fillna(0).astype(int)

# Initialize planning list
planning = []

# Loop through all games
for idx, game in games_df.iterrows():
    planning_entry = {
        'Datum': game['Datum'],
        'Start_wedstrijd': game['Start_wedstrijd'],
        'Tegenstander': game['Tegenstander'],
        'Thuis/Uit': game['Thuis/Uit'],
        'Verzamelen': game['Verzamelen']
    }

    # Always define these first!
    selected_drivers = pd.DataFrame()
    sorted_drivers = pd.DataFrame()

    if game['chauffeurs nodig'] > 0:
        needed = game['chauffeurs nodig']
        selected_drivers = pd.DataFrame()

        # Keep track of who’s already selected (to avoid duplicates)
        already_selected = set()

        # While we still need drivers
        while len(selected_drivers) < needed:
            # Find the minimum trip count among unselected drivers
            remaining_drivers = drivers_working_copy[~drivers_working_copy.index.isin(already_selected)]
            min_trips = remaining_drivers['count_trips'].min()

            # Get all drivers with that count
            eligible = remaining_drivers[remaining_drivers['count_trips'] == min_trips]

            # How many more drivers do we need?
            remaining_slots = needed - len(selected_drivers)

            # Pick as many as we can from this group
            pick_count = min(remaining_slots, len(eligible))
            picked = eligible.sample(n=pick_count, random_state=None)

            # Add to selected list and mark as used
            selected_drivers = pd.concat([selected_drivers, picked])
            already_selected.update(picked.index)

        # Update assigned trip counts
        drivers_working_copy.loc[selected_drivers.index, 'count_trips'] += 1

        # Assign drivers to planning
        for i, driver_name in enumerate(selected_drivers['Speler']):
            planning_entry[f'Chauffeur {i + 1}'] = driver_name

    # --- Laundry (Wasbeurt) assignment ---
    if not selected_drivers.empty:
        # If drivers exist, pick randomly among today's drivers
        selected_drivers_sorted = selected_drivers.sort_values('count_wassen')
        laundry_driver = selected_drivers_sorted.iloc[0]['Speler']
    else:
        # If no drivers assigned, pick the person with fewest laundry duties
        # Find the minimum wash count in the full list
        min_wash = drivers_working_copy['count_wassen'].min()

        # Get all drivers with that minimum count
        eligible = drivers_working_copy[drivers_working_copy['count_wassen'] == min_wash]

        # Pick one randomly from those eligible
        laundry_driver = random.choice(eligible['Speler'].tolist())

    # Update laundry duty
    planning_entry['Wasbeurt'] = laundry_driver
    drivers_working_copy.loc[drivers_working_copy['Speler'] == laundry_driver, 'count_wassen'] += 1

    # Add entry to planning list
    planning.append(planning_entry)

# --- Create planning DataFrame ---
planning_df = pd.DataFrame(planning)

# --- Output section ---
output_folder = r'C:\Users\stege\OneDrive\Documents\vesta_automation\output'
os.makedirs(output_folder, exist_ok=True)  # Create the output folder if it doesn't exist

# Add today's date to the filename
today_str = datetime.today().strftime('%Y-%m-%d')
output_filename = f'driver_planning_full_{today_str}.xlsx'
output_path = os.path.join(output_folder, output_filename)

# Save new file
with pd.ExcelWriter(output_path, engine='openpyxl', datetime_format='DD-MM-YYYY') as writer:
    drivers_working_copy.to_excel(writer, sheet_name='drivers', index=False)
    games_df.to_excel(writer, sheet_name='games', index=False)
    planning_df.to_excel(writer, sheet_name='planning', index=False)

print(f"Planning saved successfully to: {output_path}")

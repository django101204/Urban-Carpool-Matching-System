import pandas as pd
from geopy.distance import geodesic
from datetime import datetime

# Load dataset
def load_dataset(file_path):
    try:
        return pd.read_excel(file_path)
    except Exception as e:
        print(f"Error loading dataset: {e}")
        return None

# Calculate distance between two locations using geodesic distance
def calculate_distance(loc1, loc2):
    return geodesic(loc1, loc2).km

# Check time compatibility within a threshold (e.g., 15 minutes)
def time_compatible(driver_time, rider_time, threshold=15):
    today = datetime.today().date()
    driver_dt = datetime.combine(today, driver_time)
    rider_dt = datetime.combine(today, rider_time)
    return abs((driver_dt - rider_dt).total_seconds() / 60) <= threshold

# Form carpool groups
def form_carpool_groups(data):
    drivers = data[data['driver_rider'] == 'Driver']
    riders = data[data['driver_rider'] == 'Rider']
    groups = []

    for _, driver in drivers.iterrows():
        try:
            driver_loc = (driver['start_location_lat'], driver['start_location_lon'])
            driver_dest = (driver['destination_location_lat'], driver['destination_location_lon'])
            driver_time = driver['time_of_travel']
            group = {
                "driver": driver['name'],
                "riders": [],
                "start_location": driver_loc,
                "destination_location": driver_dest,
                "time": driver_time,
                "carbon_saved": 0  # in kg
            }
            for _, rider in riders.iterrows():
                try:
                    rider_loc = (rider['start_location_lat'], rider['start_location_lon'])
                    rider_dest = (rider['destination_location_lat'], rider['destination_location_lon'])
                    rider_time = rider['time_of_travel']
                    # Check preferences and compatibility
                    if (calculate_distance(driver_loc, rider_loc) <= driver['max_detour_distance'] and
                        calculate_distance(driver_dest, rider_dest) <= driver['max_detour_distance'] and
                        time_compatible(driver_time, rider_time) and
                        (not driver['same_gender'] or driver['gender'] == rider['gender']) and
                        (not driver['non_smoking'] or rider['non_smoking'])):
                        group['riders'].append(rider['name'])
                        # Approximate carbon savings: 120g CO2 per km
                        group['carbon_saved'] += 0.12 * calculate_distance(rider_loc, rider_dest)
                except Exception as e:
                    print(f"Error processing rider {rider['name']}: {e}")
                    continue

            if group['riders']:
                groups.append(group)

        except Exception as e:
            print(f"Error processing driver {driver['name']}: {e}")
            continue

    return groups

# Save groups to an Excel file
def save_groups_to_excel(groups, file_path):
    rows = []
    for i, group in enumerate(groups):
        for rider in group['riders']:
            rows.append({
                "Group ID": i + 1,
                "Driver": group['driver'],
                "Rider": rider,
                "Start Location": group['start_location'],
                "Destination Location": group['destination_location'],
                "Travel Time": group['time'],
                "Carbon Footprint Saved (kg CO2)": round(group['carbon_saved'], 2)
            })
    # Convert rows to a DataFrame and save as Excel
    df = pd.DataFrame(rows)
    df.to_excel(file_path, index=False)
    print(f"Groups saved to {file_path}")

# Main program
def main():
    # File path to the dataset
    dataset_path = "/home/ridham/Desktop/METENG/MEC2024Dataset.xlsx"

    # Load dataset
    data = load_dataset(dataset_path)
    if data is None:
        return

    # Debug: Print column names
    print("Columns in dataset:", data.columns)

    # Parse coordinates (adjust column names as needed)
    try:
        data['start_location_lat'] = data['start_location'].apply(lambda x: float(x.split(',')[0]))
        data['start_location_lon'] = data['start_location'].apply(lambda x: float(x.split(',')[1]))
        data['destination_location_lat'] = data['destination_location'].apply(lambda x: float(x.split(',')[0]))
        data['destination_location_lon'] = data['destination_location'].apply(lambda x: float(x.split(',')[1]))
    except KeyError as e:
        print(f"Error: Missing column in dataset - {e}")
        return
    except Exception as e:
        print(f"Error processing coordinates: {e}")
        return

    # Form carpool groups
    groups = form_carpool_groups(data)

    # Save groups to Excel
    output_excel_path = "/home/ridham/Desktop/METENG/CarpoolGroups.xlsx"
    save_groups_to_excel(groups, output_excel_path)

if __name__ == "__main__":
    main()

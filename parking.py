import pandas as pd
import time
import os

# Define the Excel file path
file_path = "C:\\Users\\nikolav\\Desktop\\richyPy\\parking\\parking_rotation.xlsx"
index_file_path = "C:\\Users\\nikolav\\Desktop\\richyPy\\parking\\rotation_index.txt"

# Check if the Excel file exists
if not os.path.exists(file_path):
    # Define the initial data
    names = ["Milovan", "Nikola", "Matija", "Selma", "Ivan", "Bogdan"]
    plus_group = {"Nikola", "Matija", "Bogdan", "Selma"}

    # Create the initial DataFrame
    data = {
        "Name": names,
        "+": ["+" if name in plus_group else "" for name in names],
        "Times in 3 Spots": [0] * len(names),
        "Times in 4th Spot": [0] * len(names),
        "Times Not in Any Spot": [0] * len(names),
        "Vacation": ["" for _ in names],
    }
    df = pd.DataFrame(data)
    os.makedirs(os.path.dirname(file_path), exist_ok=True)  # Ensure the directory exists
    df.to_excel(file_path, index=False)
else:
    # Load the existing file
    df = pd.read_excel(file_path)
    df["Vacation"] = df["Vacation"].fillna("")  # Ensure no NaN values in Vacation column

    # Ensure C, D, E columns are filled with 0 if cells are not numbers
    for col in ["Times in 3 Spots", "Times in 4th Spot", "Times Not in Any Spot"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
            print(f"Updated column '{col}' to ensure all cells are numeric:")
            print(df[col])

    # Save the DataFrame back after filling zeros
    df.to_excel(file_path, index=False)
    print("Saved the updated DataFrame with zeros filled where necessary.")

# Dynamically update names and plus_group based on the Excel file
names = df["Name"].tolist()
plus_group = set(df[df["+"] == "+"]["Name"].tolist())

# Load or initialize rotation index
if os.path.exists(index_file_path):
    with open(index_file_path, "r") as index_file:
        index = int(index_file.read().strip())
else:
    index = 0

# Get the number of cycles to run
num_cycles = int(input("Enter the number of cycles to run (default is 1): ") or 1)

# Run the rotation for the specified number of cycles
for _ in range(num_cycles):
    # Reload names and plus_group in case the Excel file has changed
    df = pd.read_excel(file_path)
    df["Vacation"] = df["Vacation"].fillna("")
    names = df["Name"].tolist()
    plus_group = set(df[df["+"] == "+"]["Name"].tolist())

    # Exclude users on vacation
    for name in names:
        vacation_status = str(df.loc[df["Name"] == name, "Vacation"].values[0])
        if vacation_status.strip():
            print(f"{name}: VAC")

    active_users = df[df["Vacation"].str.strip() == ""]
    active_names = active_users["Name"].tolist()

    # Determine the users for the 3 spots
    users_3_spots = [active_names[(index + i) % len(active_names)] for i in range(3)] if len(active_names) >= 3 else active_names
    while len(users_3_spots) < 3:
        users_3_spots.append("Empty")

    # Sort the fourth spot candidates by their usage to balance assignments
    fourth_spot_candidates = [name for name in active_names if name in plus_group and name not in users_3_spots]
    if fourth_spot_candidates:
        fourth_spot_candidates = sorted(
            fourth_spot_candidates,
            key=lambda name: df.loc[df["Name"] == name, "Times in 4th Spot"].values[0]
        )
        user_4th_spot = fourth_spot_candidates[0]
    else:
        # Assign someone without a spot for this turn
        remaining_candidates = [name for name in active_names if name not in users_3_spots]
        user_4th_spot = remaining_candidates[0] if remaining_candidates else "Empty"

    # Update the usage counts
    for name in names:
        if name in users_3_spots:
            df.loc[df["Name"] == name, "Times in 3 Spots"] += 1
        elif name == user_4th_spot:
            df.loc[df["Name"] == name, "Times in 4th Spot"] += 1
        else:
            df.loc[df["Name"] == name, "Times Not in Any Spot"] += 1

    # Save the updated DataFrame back to the Excel file
    df.to_excel(file_path, index=False)

    # Print the current rotation
    print("=======================")
    print(f"3 Spots: {', '.join(users_3_spots)}")
    print(f"4th Spot: {user_4th_spot if user_4th_spot != 'Empty' else 'None'}")
    print("=======================")

    # Increment the rotation index
    index = (index + 1) % len(active_names) if len(active_names) > 0 else 0

    # Save the current index to file
    with open(index_file_path, "w") as index_file:
        index_file.write(str(index))
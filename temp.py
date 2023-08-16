import pandas as pd

sensor_info_path = r"D:\Zainab\SUB04\left_right_sensors.xlsx"
output_path = r"D:\Zainab\SUB04\final.xlsx"
left_sensor_path = r"D:\Zainab\SUB04\tfr_alpha_left_occipital.csv"
right_sensor_path = r"D:\Zainab\SUB04\tfr_alpha_right_occipital.csv"
print("Read Paths")

# Read Left Right Sensors file
sensor_info = pd.read_excel(sensor_info_path, usecols=[0, 1])  # Read only the first two columns

# Filter out empty rows
sensor_info.dropna(subset=['right sensors', 'left sensors'], inplace=True)

# Sort the DataFrame by 'left sensors' column
sensor_info.sort_values('left sensors', inplace=True)

# Read left occipital file
left_tfr = pd.read_csv(left_sensor_path)

# Read right occipital file
right_tfr = pd.read_csv(right_sensor_path)

# Sorting TFR
left_tfr.sort_values('channel', inplace=True)
right_tfr.sort_values('channel', inplace=True)

# Drop rows where 'ch_type' is not 'grad' from left_tfr DataFrame
left_tfr = left_tfr[left_tfr['ch_type'] == 'grad']

# Drop rows where 'ch_type' is not 'grad' from right_tfr DataFrame
right_tfr = right_tfr[right_tfr['ch_type'] == 'grad']

# Drop all columns except 'channel' and 'value' from left_tfr DataFrame
left_tfr = left_tfr[['channel', 'value']]

# Drop all columns except 'channel' and 'value' from right_tfr DataFrame
right_tfr = right_tfr[['channel', 'value']]

# Remove 'MEG' and convert 'channel' values to float64 in left_tfr DataFrame
left_tfr['channel'] = left_tfr['channel'].str.replace('MEG', '').astype('float64')

# Remove 'MEG' and convert 'channel' values to float64 in right_tfr DataFrame
right_tfr['channel'] = right_tfr['channel'].str.replace('MEG', '').astype('float64')

# Reset the index to have ascending integers
left_tfr.reset_index(drop=True, inplace=True)
right_tfr.reset_index(drop=True, inplace=True)

individual_channels = len(left_tfr) // len(sensor_info)

# Create an empty DataFrame
result_df = pd.DataFrame(columns=['Left Sensor Label', 'Left Value', 'Right Sensor Label', 'Right Value'])

# Set the countdown variable
countdown = len(sensor_info)

# Loop through each row in sensor_info
for index, row in sensor_info.iterrows():
    left_sensor = row['left sensors']
    right_sensor = row['right sensors']

    # Find the index of the matching channel in left_tfr and right_tfr
    left_index = left_tfr[left_tfr['channel'] == left_sensor].index[0]
    right_index = right_tfr[right_tfr['channel'] == right_sensor].index[0]

    # Select channels and values from left_tfr using iloc
    left_selected = left_tfr.iloc[left_index:left_index + individual_channels,
                    [0, 1]]  # Assuming column positions are 0 (channel) and 1 (value)

    # Select channels and values from right_tfr using iloc
    right_selected = right_tfr.iloc[right_index:right_index + individual_channels,
                     [0, 1]]  # Assuming column positions are 0 (channel) and 1 (value)

    # Append to result_df
    for i in range(individual_channels):
        result_df = result_df._append({
            'Left Sensor Label': left_selected.iloc[i]['channel'],
            'Left Value': left_selected.iloc[i]['value'],
            'Right Sensor Label': right_selected.iloc[i]['channel'],
            'Right Value': right_selected.iloc[i]['value']
        }, ignore_index=True)

    # Decrease the countdown
    countdown -= 1

    # Print the countdown value
    print(f"Iteration {countdown} remaining")

# Add a new column to result_df
result_df['Calculation'] = (result_df['Right Value'] - result_df['Left Value']) / (result_df['Right Value'] + result_df['Left Value'])

# Calculate the average of 'Calculation' column
average_calculation = result_df['Calculation'].mean()

# Add a new column with the average value
result_df['Average Calculation'] = average_calculation

# Export the DataFrame to a CSV file
result_df.to_csv('result.csv', index=False)
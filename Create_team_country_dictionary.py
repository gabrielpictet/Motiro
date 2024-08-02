#Created a dictionary from existing Moti file
# in this case: ungrouped_ALL.csv
import pandas as pd

# Load the dataset
file_path = "C:/Users/gabriel.pictet/Documents/Gabriel/REAL/Moti/Moti data/ungrouped_ALL.csv"
df = pd.read_csv(file_path)

# Extract the "Team Name" and "country" columns
team_country_df = df[['Team Name', 'country']].drop_duplicates()

# Group by "Team Name" to ensure each team is represented only once
team_country_df = team_country_df.groupby('Team Name').first().reset_index()

# Save the new data as a text file in the desired format
output_file_path = 'team_country_dictionary.csv'
with open(output_file_path, 'w', encoding='utf-8') as file:
    for index, row in team_country_df.iterrows():
        file.write(f"'{row['Team Name']}':'{row['country']}',\n")

print(f"Dictionary saved to {output_file_path}")
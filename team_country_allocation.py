import pandas as pd

def read_team_country_dict(team_country_dict_file):
	team_country_dict = {}
	try:
		with open(team_country_dict_file, 'r', encoding='utf-8') as file:
			for line in file:
				if line.strip():  # Skip empty lines
					parts = line.strip().split(':')
					if len(parts) == 2:
						team, country = parts
						team_country_dict[team.strip().strip("'")] = country.strip().strip("'")
	except FileNotFoundError:
		pass
	return team_country_dict

def write_team_country_dict(team_country_dict_file, team_country_dict):
	with open(team_country_dict_file, 'w', encoding='utf-8') as file:
		for team, country in team_country_dict.items():
			file.write(f"'{team}':'{country}',\n")

# Path to the team-country dictionary file
team_country_dict_file = 'team_country_dictionary.csv'

# Read the existing team-country dictionary
team_country_dict = read_team_country_dict(team_country_dict_file)

# Load the updated dataset
file_path = 'C:/Users/gabriel.pictet/Documents/Gabriel/REAL/Moti/Moti data/ungrouped_ALL.csv'
df = pd.read_csv(file_path, sep=",", encoding='utf-8')

# Create an empty 'country' column
df['country'] = ''

# Iterate over unique 'Team Name' values and allocate country
for team_name in df['Team Name'].unique():
	# Check if the team name is in the dictionary
	if team_name in team_country_dict:
		country = team_country_dict[team_name]
	else:
		# If the team name is not in the dictionary, prompt for user input
		country = input(f"Enter the country for Team Name '{team_name}': ")
		# Update the dictionary with the new entry
		team_country_dict[team_name] = country
	
	df.loc[df['Team Name'] == team_name, 'country'] = country

# Write the updated dictionary back to the file
write_team_country_dict(team_country_dict_file, team_country_dict)

# Save the updated dataset if needed
updated_file_path = 'C:/Users/gabriel.pictet/Documents/Gabriel/REAL/Moti/Moti data/country_ungrouped_ALL.txt'
df.to_csv(updated_file_path, index=False, encoding='utf-8')

print(f"Updated dictionary saved to {team_country_dict_file}")
print(f"Updated dataset saved to {updated_file_path}")
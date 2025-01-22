import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches

# Load the data
df = pd.read_csv('Individual.csv') 

# Generate a plot
#plt.figure(figsize=(10, 6))
#plt.bar(df['Country'], df['Users'], color='skyblue')
#plt.xlabel('Country')
#plt.ylabel('Number of Users')
#plt.title('Number of Users by Country')
#plt.savefig('users_by_country.png')
#plt.close()

# create date variable from the 'Survey Data' column using the most recent time stamp.
df['Survey Data'] = pd.to_datetime(df['Survey Data'], format='mixed', errors='coerce')
most_recent_date = df['Survey Data'].max().strftime('%d-%B-%Y')

# Calculate statistics
total_volunteers = df['Volunteer'].sum()
total_staff = df['Staff'].sum()
# count unique occurence of teams and countries
total_teams = df['Team Name'].nunique()
total_countries = df['Country'].nunique()
# enter the app's current number of languages
total_languages = 11

# Create a Word Document
doc = Document()
doc.add_heading('Motiro Status Report', level=1)
doc.add_heading('Motiro usage data', level=2)
# Add Text
doc.add_paragraph(f'As of {most_recent_date}, the Motiro app has been used by:')
doc.add_paragraph(f'{total_volunteers} volunteers', style='ListBullet')
doc.add_paragraph(f'{total_staff} staff', style='ListBullet')
doc.add_paragraph(f'from {total_teams} teams', style='ListBullet')
doc.add_paragraph(f'belonging to {total_countries} RCRC entities.', style='ListBullet')
doc.add_paragraph(f'The Motiro app (https://motiro.com) works in {total_languages} languages including the four official IFRC languages.')

# Add Plot
doc.add_picture('RespondentsByCountrySorted.png', width=Inches(6), height=Inches(5))
doc.add_picture('RespondentsByRegionSorted.png', width=Inches(6), height=Inches(5))
doc.add_picture('Moti_by_country.png', width=Inches(6), height=Inches(4))

doc.add_heading('Motiro findings', level=2)
doc.add_paragraph('The app has been used to collect and analyse data on motivation to answer the following questions:')
doc.add_paragraph('What is the quality of volunteer and staff motivation and engagement?', style='ListBullet')
doc.add_paragraph('What are the key drivers of motivation and engagement?', style='ListBullet')
doc.add_paragraph('What motivates volunteers and staff to stay engaged in their team?', style='ListBullet')
doc.add_paragraph('What are the pathways toward improved motivation, engagement and well-being?', style='ListBullet')

doc.add_picture('violin_regions 2x3.png', width=Inches(6), height=Inches(8))
doc.add_picture('Volunteers SDTCorrNetworkGraph.png', width=Inches(6), height=Inches(4))
doc.add_picture('Staff SDTCorrNetworkGraph.png', width=Inches(6), height=Inches(4))
# Save Document
doc.save('MotiroDashboard.docx')


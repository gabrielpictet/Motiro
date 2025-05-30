#make a pie chart comparing number of volunteers with number of paid staff
# import the necessary libraries
import matplotlib
import matplotlib.pyplot as plt
import pandas as pd

# load dataframe individual.csv
df = pd.read_csv("individual.csv")

# count the number of volunteers
volunteers = df["Volunteer"].sum()
# count the number of paid staff
paid_staff = df["Staff"].sum()

# put the data into a list
data = [volunteers, paid_staff]

# make the pie chart
plt.pie(data, labels=["Volunteers", "Paid Staff"], autopct='%1.1f%%')
plt.axis("equal")
plt.title("Volunteers vs Paid Staff")
plt.show()
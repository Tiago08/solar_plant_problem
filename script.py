# Import the needed packages
import time
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

# get the user input of the file name
file_name = input("Enter the file name: ")

# start a timer
start_time = time.time()
# create folders to hold the visualizations and the texts
if os.path.exists("Visualizations") and os.path.exists("Texts"):
    pass
else:
    os.mkdir("Visualizations")
    os.mkdir("Texts")

# load the data with the specified columns
df = pd.read_excel(file_name+".xlsx",header=0, sheet_name=0, usecols=["fecha_im", "active_power_im"])

# clean the active power column
bad_indices = df.query('active_power_im == "data_faltante"').index
df.drop(bad_indices, inplace=True)
df["active_power_im"] = df.active_power_im.astype(float)

# Create the visualization
sns.set_style("darkgrid", {"axes.facecolor": ".9"})
fig, ax = plt.subplots()
sns.lineplot(data=df, x="fecha_im", y="active_power_im", errorbar=None)
ax.set_xlabel("Time of day", fontsize=12)
ax.set_ylabel("Active power", fontsize=12)
plt.title("Active power during the day")
plt.xticks(rotation = 45)
plt.grid(visible=True, which='major', color='k', linestyle='-')
plt.savefig('visualizations/{}.png'.format(file_name), bbox_inches='tight');

# create the active power summary
active_power_sum = df["active_power_im"].sum()
active_power_max = df["active_power_im"].max()
active_power_min = df["active_power_im"].min()

# create the text file for the summary
with open("texts/{}.txt".format(file_name),'w',encoding = 'utf-8') as f:
   f.write("Active power summation: {} \n".format(active_power_sum))
   f.write("Active power Max: {}\n".format(active_power_max))
   f.write("Active power Min: {}\n".format(active_power_min))
   f.write("The viusalization path: Visualizations/{}.png".format(file_name))

# calculate all active power for all plants
total_active_power = 0
for file in os.listdir(os.getcwd()):
    f = os.path.join(os.getcwd(), file)
    # checking if it is a file
    if "xlsx" in f:
        df = pd.read_excel(f,header=0, sheet_name=0, usecols=["fecha_im", "active_power_im"])
        bad_indices = df.query('active_power_im == "data_faltante"').index
        df.drop(bad_indices, inplace=True)
        df["active_power_im"] = df.active_power_im.astype(float)
        active_power = df["active_power_im"].sum()
        total_active_power += active_power
print("The total active power for all plants = {}".format(total_active_power))

# print the time needed for the script
print("Time elapsed = {:.2f}".format(time.time() - start_time))
# Import the needed packages
import time
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import glob
import os

# define a filter function
def get_filters():
    """
    Asks user to specify one file or all files to analyze.

    Returns:
    (str) all_one - "one" for a single excel file, "all" for all excel files inside the directory
    (str) file_name - name of a single file 
    
    """
    cd = os.getcwd()
    all_one = input("Do you want to process all files in directory or just one? input 'all' for all files and 'one' for one file: ")
    while all_one == "one" :
            file_name = input("Enter the file name: ")
            if not(os.path.isfile("{}/{}.xlsx".format(cd, file_name))):
                print("--------------invaild input try again--------------")
                continue
            else:
                break
    else:
        file_name = ""
    return all_one, file_name



# Define a function to create dataframes
def load_dataframe(all_one, file_name):
    """
    Loads data for the specified excel sheet/s

    Args:
        (str) all_one - "one" for a single excel file, "all" for all excel files inside the directory
        (str) file_name - name of a single file
    Returns:
        (pandas dataframe) dfs - Pandas DataFrame/s containing active power and date columns
        (list) filenames - Excel file names
    """
    path = os.getcwd()
    global filenames 
    filenames = glob.glob(path + "\*.xlsx")
    global dfs 
    dfs = []
    if all_one == "all":
        for file in filenames:
            dfs.append(pd.read_excel(file,header=0, sheet_name=0, usecols=["fecha_im", "active_power_im"]))
    else:
        dfs.append(pd.read_excel(file_name+".xlsx",header=0, sheet_name=0, usecols=["fecha_im", "active_power_im"]))
        dfs = dfs[0] 

    return dfs, filenames


# Define a cleaning function
def clean_data(all_one, dfs):
    """
    Removes rows with "data_faltante" values and then converts the active power column to float type

    Args:
        (str) all_one - "one" for a single excel file, "all" for all excel files inside the directory
        (pandas dataframe) dfs - Pandas DataFrame/s containing active power and date columns
    
    Returns:
        (pandas dataframe) dfs - clean Pandas DataFrame/s containing active power and date columns
    """
    if all_one == "all":
        for df in dfs:
            bad_indices = df.query('active_power_im == "data_faltante"').index
            df.drop(bad_indices, inplace=True)
            df["active_power_im"] = df.active_power_im.astype(float)

    else:
        bad_indices = dfs.query('active_power_im == "data_faltante"').index
        dfs.drop(bad_indices, inplace=True)
        dfs["active_power_im"] = dfs.active_power_im.astype(float)
    return dfs

# Define a visualization function
def visualize(all_one, dfs, file_name, filenames):
    """
    Creates and saves a line plot of the active power through the day

    Args:
        (str) all_one - "one" for a single excel file, "all" for all excel files inside the directory
        (pandas dataframe) dfs - Pandas DataFrame/s containing active power and date columns
        (str) file_name - name of a single file
        (list) filenames - Excel file names
    
    """
    if all_one =="all":
        for df, file_name in zip(dfs, filenames):
            sns.set_style("darkgrid", {"axes.facecolor": ".9"})
            fig, ax = plt.subplots()
            sns.lineplot(data=df, x="fecha_im", y="active_power_im", errorbar=None)
            ax.set_xlabel("Time of day", fontsize=12)
            ax.set_ylabel("Active power", fontsize=12)
            plt.title("Active power during the day")
            plt.xticks(rotation = 45)
            plt.grid(visible=True, which='major', color='k', linestyle='-')
            plt.savefig('{}.png'.format(file_name.strip(".xlsx")), bbox_inches='tight');
    else:
        sns.set_style("darkgrid", {"axes.facecolor": ".9"})
        fig, ax = plt.subplots()
        sns.lineplot(data=dfs, x="fecha_im", y="active_power_im", errorbar=None)
        ax.set_xlabel("Time of day", fontsize=12)
        ax.set_ylabel("Active power", fontsize=12)
        plt.title("Active power during the day")
        plt.xticks(rotation = 45)
        plt.grid(visible=True, which='major', color='k', linestyle='-')
        plt.savefig('{}.png'.format(file_name.strip(".xlsx")), bbox_inches='tight');

# Define summary statistics function
def summary(all_one, dfs):
    """
    Calculates summary statistics for the active power

    Args:
        (str) all_one - "one" for a single excel file, "all" for all excel files inside the directory
        (pandas dataframe) dfs - Pandas DataFrame/s containing active power and date columns
    
    Returns:
        (list) active_power_sums - a list of the summations of the acrive power column of the dataframe/s
        (list) active_power_maxs - a list of the maximums of the acrive power column of the dataframe/s
        (list) active_power_mins - a list of the minumums of the acrive power column of the dataframe/s
    """
    global active_power_sums
    global active_power_maxs
    global active_power_mins
    active_power_sums = list()
    active_power_maxs = list()
    active_power_mins = list()
    if all_one =="all":
        for df in dfs:
            active_power_sums.append(df["active_power_im"].sum())
            active_power_maxs.append(df["active_power_im"].max())
            active_power_mins.append(df["active_power_im"].min())
    else:
        active_power_sums.append(dfs["active_power_im"].sum())
        active_power_maxs.append(dfs["active_power_im"].max())
        active_power_mins.append(dfs["active_power_im"].min())
        active_power_sums = active_power_sums[0]
        active_power_maxs = active_power_maxs[0]
        active_power_mins = active_power_mins[0]
    return active_power_sums, active_power_maxs, active_power_mins

# create the text file for the summary
def create_text_file(all_one, file_name, filenames, active_power_sums, active_power_maxs, active_power_mins):
    """
    Creates text file/s for the summary statistics and visualizations paths

    Args:
        (str) all_one - "one" for a single excel file, "all" for all excel files inside the directory
        (pandas dataframe) dfs - Pandas DataFrame/s containing active power and date columns
        (str) file_name - name of a single file
        (list) filenames - Excel file names
        (list) active_power_sums - a list of the summations of the acrive power column of the dataframe/s
        (list) active_power_maxs - a list of the maximums of the acrive power column of the dataframe/s
        (list) active_power_mins - a list of the minumums of the acrive power column of the dataframe/s

    """
    if all_one =="all":
        for file_name in filenames:
            with open("{}.txt".format(file_name.strip(".xlsx")),'w',encoding = 'utf-8') as f:
               f.write("Active power summation: {} \n".format(active_power_sums))
               f.write("Active power Max: {}\n".format(active_power_maxs))
               f.write("Active power Min: {}\n".format(active_power_mins))
               f.write("The viusalization path: {}.png".format(file_name.strip(".xlsx")))
    else:
        with open("{}.txt".format(file_name.strip(".xlsx")),'w',encoding = 'utf-8') as f:
               f.write("Active power summation: {} \n".format(active_power_sums))
               f.write("Active power Max: {}\n".format(active_power_maxs))
               f.write("Active power Min: {}\n".format(active_power_mins))
               f.write("The viusalization path: {}.png".format(file_name.strip(".xlsx")))
# Define a function that prints to the console a summary statistic
def console_power_sum(filenames):
    """
    Prints the summation of active power for every excel file in the directory

    Args:
        (list) filenames - Excel file names

    """
    total_active_power = 0
    for file in filenames:
            df = pd.read_excel(file,header=0, sheet_name=0, usecols=["fecha_im", "active_power_im"])
            bad_indices = df.query('active_power_im == "data_faltante"').index
            df.drop(bad_indices, inplace=True)
            df["active_power_im"] = df.active_power_im.astype(float)
            active_power = df["active_power_im"].sum()
            total_active_power += active_power
    print("The total active power for all plants = {}".format(total_active_power))

# Define a timer function
def timer(start_time):
    """
    Calculates the time taken to run the script

    Args:
        (int) start_time - the starting time of the script
    """
    print("Time elapsed = {:.2f}".format(time.time() - start_time))

def main():
    while True:
        start_time = time.time()
        all_one, file_name = get_filters()
        dfs, filenames = load_dataframe(all_one, file_name)
        dfs = clean_data(all_one, dfs)
        active_power_sums, active_power_maxs, active_power_mins = summary(all_one, dfs)
        visualize(all_one, dfs, file_name, filenames)
        create_text_file(all_one, file_name, filenames, active_power_sums, active_power_maxs, active_power_mins)
        console_power_sum(filenames)
        timer(start_time)
        restart = input('\nWould you like to restart? Enter yes or no.\n')
        if restart.lower() != 'yes':
            break

if __name__ == "__main__":
	main()
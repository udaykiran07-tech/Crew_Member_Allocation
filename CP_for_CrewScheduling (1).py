# Importing all important library

import pandas as pd
import openpyxl
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime, timedelta
import matplotlib.patches as mpatches
import numpy as np
import os


# Output folder name
Output_folder_name = "OUTPUT"
# Path of your excel file
excel_file_path = "cp for crewscheduling.xlsx"


fontfamily = "Times New Roman"

# Some useful function
is_shiffled = True
to_optimize = True



# print(all_sheets.keys())


# Make Distionary for the ploting
crew_members_dict = []
pilot_members_dict = []
copilot_members_dict = []
loadmaster_members_dict = []
flight_engineer_members_dict = []

def shuffle_sheet(sheet):
    # Use the sample method to shuffle the rows randomly
    shuffled_sheet = sheet.sample(
        frac=1).reset_index(drop=True)

    return shuffled_sheet

def plot_bar_graph(data_dict, crewClasses_with_total, day):
    keys = list(data_dict.keys())
    values = list(data_dict.values())
    num_bars = len(values[0])

    plt.figure()

    bar_width = 0.1  # You can adjust the width of the bars

    for i in range(num_bars):
        # if len(values[i]) == 0:
        #     continue
        x_positions = np.arange(len(keys)) + i * bar_width
        plt.bar(x_positions, [value[i] for value in values],
                width=bar_width, label=crewClasses_with_total[i])

    plt.xlabel('Fleet IDs')
    plt.ylabel('Number of Crew Members')
    plt.title('Number of Crew Members for each Fleet ID on Day ' + str(day))
    plt.xticks(np.arange(len(keys)) + (num_bars - 1) * bar_width / 2, keys)
    plt.legend(loc='best')

    plt.savefig('Day' + str(day) + '.png')

def plot_data(data):
    x_values, y_values = zip(*data)
    plt.figure(figsize=(10, 5))
    plt.xticks(np.linspace(min(x_values), max(x_values), max(
        x_values)-min(x_values)+1))  # 10 ticks on the x-axis
    plt.yticks(np.linspace(min(y_values), max(y_values), max(
        y_values)-min(y_values)+1))  # 10 ticks on the y-axis
    plt.plot(x_values, y_values, marker='o')
    plt.title('No. of Crew Members Used')
    x_label = 'Iteration' if not to_optimize else 'Best Solution'
    plt.xlabel(x_label)
    plt.ylabel('No. of Crew Members Used')
    plt.grid(True)
    plt.savefig('NoOfCrewMembersUsed.png')

def plot_data_all(pilot_data, copilot_data, loadmaster_data, flight_engineer_data):

    pilot_data_x_values, pilot_data_y_values = zip(*pilot_data)
    copilot_data_x_values, copilot_data_y_values = zip(*copilot_data)
    loadmaster_data_x_values, loadmaster_data_y_values = zip(*loadmaster_data)
    flight_engineer_data_x_values, flight_engineer_data_y_values = zip(*flight_engineer_data)

    plt.figure(figsize=(10, 5))
    plt.xticks(np.linspace(min(pilot_data_x_values), max(pilot_data_x_values), max(
        pilot_data_x_values)-min(pilot_data_x_values)+1))  # 10 ticks on the x-axis
    # plt.yticks(np.linspace(min(pilot_data_y_values), max(pilot_data_y_values), max(
        # pilot_data_y_values)-min(pilot_data_y_values)+1))  # 10 ticks on the y-axis
    plt.plot(pilot_data_x_values, pilot_data_y_values, marker='o', label='Pilot')
    plt.plot(copilot_data_x_values, copilot_data_y_values, marker='*', label='Co-Pilot')
    plt.plot(loadmaster_data_x_values, loadmaster_data_y_values, marker='+', label='Loadmaster')
    plt.plot(flight_engineer_data_x_values, flight_engineer_data_y_values, marker='x', label='Flight Engineer')
    plt.title('No. of Each Crew Members Used')
    x_label = 'Iteration' if not to_optimize else 'Best Solution'
    plt.xlabel(x_label)
    plt.ylabel('No. of Crew Members Used')
    plt.legend()
    plt.grid(True)
    plt.savefig('NoOfEachCrewMembersUsed.png')

def time_to_minutes(time, day):
    '''This function takes a time in the format of HH:MM and converts it to minutes.
    It also takes a day of the week as a string and converts it to a number.'''
    # Convert time to minutes
    time = time.split(':')
    hours = int(time[0])
    minutes = int(time[1])
    time = hours*60 + minutes
    # Convert day to number
    day_min = (day-1)*24*60
    # Add day to time
    time = time + day_min
    return time

def get_time(minute):
    '''Converts time in minutes to HH:MM format.'''
    hours = minute // 60
    minutes = minute % 60

    date = datetime.now().date()
    date = date + timedelta(days=hours // 24)
    hours = hours % 24

    return pd.to_datetime(f"{date} {hours:02d}:{minutes:02d}")

def get_colour(crew_type):
    # Function to map crew types to colors
    if crew_type == "Pilot":
        return 'tab:blue'
    elif crew_type == "CoPilot":
        return 'tab:orange'
    elif crew_type == "LoadMaster":
        return 'tab:green'
    elif crew_type == "FlightEngineer":
        return 'tab:red'
    else:
        return 'm'


# Input of data, Processing of data, and Solution of the problem
def Allocation_crew_member_in_fights(excel_file_path, Output_folder_name, iteration):

    #  Read all sheets of excel file
    all_sheets = pd.read_excel(excel_file_path, sheet_name=None)
# For the understanding for flight
    LOCATION_f = 0
    FLEET_TYPE_f = 1
    FLEET_IDS_f = 2
    DEP_MIN_f = 3
    ARR_MIN_f = 4
    DAY_f = 5
    FLIGHT_TIME_f = 6

# For the understanding for crew member
    ID_c = 0
    QUALIFICATION_c = 1
    HOME_LOCATION_c = 2
    AVAILABLE_HOURS_c = 3
    FLEET_TYPE_c = 4

# boolean type crew member
    MEMBER_TYPE_cm = 0
    MEMBER_NAME_cm = 1

# Define pilot, copilot , loadmaster, and flight engineer
    crew_members_edit = ["PILOT", "CO_PILOT", "LOAD_MASTER", "FLIGHT_ENGINEER"]
    #   print(crew_members[0])

# set value of qualification
    setOfTrainingLevelQualifications = ["A", "B", "C"]

# Importing data sets

    # Read the sheet with the flight info
    flight_df = all_sheets['FLIGHT_INFO']
    # print(flight_df)

    # Read the sheet with the crew limitation
    crew_df = all_sheets['CREW_LIMITATIONS']
    # print(crew_df)

    # Read the sheets with the pilot
    pilot_df = all_sheets['PILOT']
    # print(pilot_df)

    # Read the sheets with the copilot
    copilot_df = all_sheets['CO_PILOT']
    # print(copilot_df)

    # Read the sheets with the load master
    loadmaster_df = all_sheets['LOAD_MASTER']
    # print(loadmaster_df)

    # Read the sheets with the flight engineer
    flight_engineer_df = all_sheets['FLIGHT_ENGINEER']
    # print(flight_engineer_df)

    # Read the sheet swith the ac
    ac_df = all_sheets['AC']
    # print(ac_df)

    # read the sheets with the output
    output_df = all_sheets['output']
    # print(output_df)

    if is_shiffled:
        # Shuffle the sheets
        flight_df = shuffle_sheet(flight_df)
        pilot_df = shuffle_sheet(pilot_df)
        copilot_df = shuffle_sheet(copilot_df)
        loadmaster_df = shuffle_sheet(loadmaster_df)
        flight_engineer_df = shuffle_sheet(flight_engineer_df)

# Processing of the data

    # Adding the arr time, dep time and time of flight in minute in excel sheets
    flight_df["Dep._Min"] = flight_df.apply(
        lambda row: time_to_minutes(row["Dep."], row["Day"]), axis=1)
    flight_df["Arr._Min"] = flight_df.apply(
        lambda row: time_to_minutes(row["Arr."], row["Day"]), axis=1)
    # Flight time is the difference between arrival and departure
    flight_df["Flight_Time"] = flight_df["Arr._Min"] - flight_df["Dep._Min"]

    # print(flight_df)

    # Formation of sets and excel sheets

    # Sets
    flights = flight_df.index.to_list()
    days = flight_df["Day"].unique().tolist()
    aircraft_variants = ac_df["AircraftVariant1"].unique().tolist()
    flights = {row["Flight"]: (row["Location"], row["FleetType"], row["FleetIDs"], row["Dep._Min"], row["Arr._Min"], row["Day"], row["Flight_Time"])
               for index, row in flight_df.iterrows()}
    # print(flights)
    # print(days)
    # print(aircraft_variants)

    # Excel sheets
    ac_df.columns = ["AircraftVariant1", "Pilot", "QualificationPilot", "CoPilot",
                     "QualificationCoPilot",
                     "LoadMaster", "QualificationLoadMaster", "FlightEngineer",     "QualificationFlightEngineer"]
    # print(ac_df)

    # Flight id define
    flightIds = flights.keys()
    # print(flightIds)

    # Crew class define
    crewClasses = crew_df["crewtype"].unique().tolist()
    # print(crewClasses)

    # Define the max new flying hours
    new_crew_df = crew_df.set_index("crewtype")
    new_crew_df['newDailyMaxFlyingHours'] = new_crew_df.apply(lambda row: min(
        row['DailyMaxFlyingHours'], row['DutyDay'] - row['RestPriod']), axis=1)

    # Define the min sit time
    MinSitTime = new_crew_df["MinSitTime"].to_dict()
    # print(MinSitTime)

    # Define the min layover time
    MinLayoverTime = new_crew_df["MinLayover"].to_dict()
    # print(MinLayoverTime)

    # Define the max new flying hour
    DailyMaxFlyingHours = new_crew_df["newDailyMaxFlyingHours"].to_dict()
    # print(DailyMaxFlyingHours)

    # Define the Daily max landing hour
    DailyMaxLandings = new_crew_df["DailyMaxLandings"].to_dict()
    # print(DailyMaxLandings)

    # Define the number of crew for ac
    numberOfCrewsForAC = ac_df.set_index("AircraftVariant1").to_dict()
    # print(numberOfCrewsForAC)

    # Define Crew Member
    pilots_crews = {("Pilot", row["Name"]): (row["ID"], row["Qualification"], row["HomeLocation"], row["AvailableHours"], row["FleetType"])
                    for index, row in pilot_df.iterrows()}
    copilot_crews = {("CoPilot", row["Name"]): (row["ID"], row["Qualification"], row["HomeLocation"], row["AvailableHours"], row["FleetType"])
                     for index, row in copilot_df.iterrows()}
    loadmaster_crews = {("LoadMaster", row["Name"]): (row["ID"],  row["Qualification"], row["HomeLocation"], row["AvailableHours"])
                        for index, row in loadmaster_df.iterrows()}
    flight_engineer_crews = {("FlightEngineer",  row["Name"]): (row["ID"], row["Qualification"], row["HomeLocation"], row["AvailableHours"])
                             for index, row in flight_engineer_df.iterrows()}
    # print(loadmaster_crews)

    crew_members = {key: value for (key, value) in pilots_crews.items()}
    crew_members.update({key: value for (
        key, value) in copilot_crews.items()})
    crew_members.update({key: value for (
        key, value) in loadmaster_crews.items()})
    crew_members.update({key: value for (
        key, value) in flight_engineer_crews.items()})
    # print(crew_members.keys())

    # Qualification in 0 and 1
    is_crew_qualified = {}
    for cKey, cValue in crew_members.items():
        for qual in setOfTrainingLevelQualifications:
            is_crew_qualified[(cKey, qual)] = 0
    for cKey, cValue in crew_members.items():
        for char_code in range(ord(cValue[QUALIFICATION_c]), ord(setOfTrainingLevelQualifications[-1])+1):
            is_crew_qualified[(cKey, chr(char_code))] = 1
    # print(is_crew_qualified)

    # Convert the dictionary to a 2D array
    crew_matrix = [[is_crew_qualified.get(
        (crew, loc), False) for crew in crew_members.keys()] for loc in setOfTrainingLevelQualifications]

    # Create a DataFrame for the heatmap
    df = pd.DataFrame(crew_matrix, columns=crew_members.keys(),
                      index=setOfTrainingLevelQualifications)

    # Plot the heatmap with labels
    # Set the size of the figure
    plt.figure(figsize=(12, 6))

    # "d" formats the annotation as integers
    sns.heatmap(df, cmap="BuPu", cbar=True)

    # Set labels for x and y axes
    plt.xticks(ticks=range(len(crew_members.keys())),
               labels=crew_members.keys(), rotation=90, ha='left', va='top')
    plt.yticks(ticks=range(len(setOfTrainingLevelQualifications)),
               labels=setOfTrainingLevelQualifications, rotation=0, va='center')
    # plt.savefig("Qualification_of_crew_memebers")
    # plt.show()

    # Define flight days and their code
    flight_for_days = {}
    for day in days:
        flight_for_days[day] = [flight for flight in flights.keys(
        ) if flights[flight][DAY_f] == day]
    # print(flight_for_days)

# Check the compatibility of the two flight with their gap
    def check_not_compatibility(f1, f2, gap):
        # This function checks if two flights are compatible.
        # It takes the first flight, the second flight and the gap between them in minutes.
        # It returns 1 if the flights are compatible and 0 otherwise."""
        if f1[DEP_MIN_f] <= f2[DEP_MIN_f] <= f1[ARR_MIN_f] + gap*60:
            return 1
        elif f2[DEP_MIN_f] <= f1[DEP_MIN_f] <= f2[ARR_MIN_f] + gap*60:
            return 1
        else:
            return 0

# Check the crew is qualified for the flights
    def is_crew_qualified_for_flight(crew, flight):
        # This function checks if a crew is qualified for a flight.
        # It takes a crew and a flight and returns 1 if the crew is qualified for the flight and 0 otherwise.'''
        return is_crew_qualified[(crew,                         numberOfCrewsForAC["Qualification"+crew[MEMBER_TYPE_cm]][flight[FLEET_TYPE_f]])] == 1
# Modeling the crew scheduling problem
    # Import Model
    from docplex.cp.model import CpoModel
    model = CpoModel()

    # Decision variables
    x = model.binary_var_dict(keys=[((fID, flights[fID][DAY_f]), ckey)
                              for fID in flightIds for ckey in crew_members.keys()])

    # Adding variables to the model
    for key, value in x.items():
        model.add(value)

# Adding constraints to the model
    # Pilot and CoPilot can only be assigned to flights of their own fleet type
    for cKey, cValue in pilots_crews.items():
        for fId in flightIds:
            if flights[fId][FLEET_TYPE_f] != cValue[FLEET_TYPE_c]:
                model.add(x[((fId, flights[fId][DAY_f]), cKey)] == 0)
    for cKey, cValue in copilot_crews.items():
        for fId in flightIds:
            if flights[fId][FLEET_TYPE_f] != cValue[FLEET_TYPE_c]:
                model.add(x[((fId, flights[fId][DAY_f]), cKey)] == 0)

    # Every crew must be assigned to exactly one flight at a time
    from itertools import combinations
    for f1, f2 in combinations(flightIds, 2):
        # The gap between two flights is the minimum sit time if they are of the same fleet type, otherwise it is the minimum layover time
        for crew, cValue in crew_members.items():
            gap = MinSitTime[crew[MEMBER_TYPE_cm]
                             ] if flights[f1][FLEET_TYPE_f] == flights[f2][FLEET_TYPE_f] else MinLayoverTime[crew[MEMBER_TYPE_cm]]
            if check_not_compatibility(flights[f1], flights[f2], gap):
                model.add(x[((f1, flights[f1][DAY_f]), crew)] +
                          x[((f2, flights[f2][DAY_f]), crew)] <= 1)

    for day, fIds in flight_for_days.items():
        for crew, cValue in crew_members.items():
            # The total number of landings of a crew in a day must not exceed the daily maximum landings
            model.add(sum(x[((fId, day), crew)] for fId in fIds)
                      <= DailyMaxLandings[crew[MEMBER_TYPE_cm]])
        # The total flying hours of a crew in a day must not exceed the daily maximum flying hours
            model.add(sum(x[((fId, day), crew)]*flights[fId][FLIGHT_TIME_f]
                          for fId in fIds) <= DailyMaxFlyingHours[crew[MEMBER_TYPE_cm]]*60)

        # A crew must be qualified for the flight
            for fId in fIds:
                if not is_crew_qualified_for_flight(crew, flights[fId]):
                    model.add(x[((fId, day), crew)] == 0)

    # Total number of working hour of a crew for whole mission must not exceed the available hours
    for crew, cValue in crew_members.items():
        model.add(sum(x[((fId, flights[fId][DAY_f]), crew)]*flights[fId][FLIGHT_TIME_f]
                      for fId in flightIds) <= cValue[AVAILABLE_HOURS_c]*60)

# Each flight should have required number of crew members
    for fId in flightIds:
        for crewClasse in crewClasses:
            model.add(sum(x[((fId, flights[fId][DAY_f]), crew)] for crew in crew_members.keys(
            ) if crew[MEMBER_TYPE_cm] == crewClasse) == numberOfCrewsForAC[crewClasse][flights[fId][FLEET_TYPE_f]])

    # Objective function
    if to_optimize:
        is_allocated = model.binary_var_dict(keys=[cKey for cKey in crew_members.keys(
        )], name="is_allocated")
        for crew, cValue in crew_members.items():
            model.add((is_allocated[crew] == 0) == (
                sum(x[((fId, flights[fId][DAY_f]), crew)] for fId in flightIds) == 0))
            #
        model.minimize(sum(is_allocated[crew] for crew in crew_members.keys()))
    # Solving the model

    sol = model.solve()


# Check the solution is exist or not

    if not sol:
        print(f'No solution for {iteration}th iteration')
        return

    k = ['FailStatus', 'MemoryUsage', 'NumberOfBranches', 'NumberOfChoicePoints', 'NumberOfConstraints',  'NumberOfFails', 'NumberOfSolutions',
         'NumberOfVariables', 'PeakMemoryUsage', 'PresolveTime', 'SearchStatus', 'SearchStopCause', 'SolveTime', 'TotalTime']
# Extracting the solution
    flight_assignment = {}
    for crew in crew_members.keys():
        flight_assignment[crew] = []
        for fId in flightIds:
            if sol.get_value(x[((fId, flights[fId][DAY_f]), crew)]) == 1:
                flight_assignment[crew].append(fId)
    flight_assignment
    crew_assignment = {}
    for fId in flightIds:

        crew_assignment[fId] = {}

        for crewClasse in crewClasses:

            crew_assignment[fId][crewClasse] = []

            for crew in crew_members.keys():

                if crew[MEMBER_TYPE_cm] == crewClasse and sol.get_value(x[((fId, flights[fId][DAY_f]), crew)]) == 1:
                    crew_assignment[fId][crewClasse].append(
                        crew[MEMBER_NAME_cm])

    sol_matrix = []
    for fId in flightIds:
        sol_matrix.append([sol.get_value(x[((fId, flights[fId][DAY_f]), crew)])
                           for crew in crew_members.keys()])
    for list in sol_matrix:
        print(list)

    sol_info = {key: value for key,
                value in sol.get_solver_infos().items() if key in k}
    for key, value in sol_info.items():
        print({key: value})

    # Creating the output folder name
    Output_folder_name_ = Output_folder_name + str(iteration)
    if not os.path.exists(Output_folder_name_):
        os.makedirs(Output_folder_name_)
    os.chdir(Output_folder_name_)
    output_file_path = "Qualification_of_crew_members.png"
    plt.savefig(output_file_path, dpi=500, bbox_inches='tight')
    # Input Excel file save into output_folder_name
    unique_fleet_id = flight_df["FleetIDs"].unique().tolist()

    data_for_each_day = {}
    for day in days:
        data_for_each_day[day] = {key: [] for key in unique_fleet_id}
    for day in days:
        for fid in flightIds:
            if flights[fid][DAY_f] == day:
                data_for_each_day[day][flights[fid][FLEET_IDS_f]].append(fid)
    data_for_each_day


    for day in days:
        for fleetID, fIds in data_for_each_day[day].items():
            var_pilot = 0
            var_copilot = 0
            var_loadmaster = 0
            var_flightengineer = 0
            for fId in fIds:
                var_pilot += len(crew_assignment[fId]["Pilot"])
                var_copilot += len(crew_assignment[fId]["CoPilot"])
                var_loadmaster += len(crew_assignment[fId]["LoadMaster"])
                var_flightengineer += len(crew_assignment[fId]["FlightEngineer"])

            total = var_pilot + var_copilot + var_loadmaster + var_flightengineer
            if total != 0:
                data_for_each_day[day][fleetID] = {
                    "Pilot": var_pilot, "CoPilot": var_copilot, "LoadMaster": var_loadmaster, "FlightEngineer": var_flightengineer}

    if not os.path.exists("Days"):
        os.mkdir("Days")
    os.chdir("Days")
    # crewClasses_with_total = crewClasses + ["Total"]
    crewClasses_with_total = crewClasses
    for day in days:
        to_plot = {}
        for fleetID , fValue in data_for_each_day[day].items():
            # print(fValue)
            if len(fValue) == 0:
                continue
            # to_plot[fleetID] = (fValue["Pilot"], fValue["CoPilot"], fValue["LoadMaster"], fValue["FlightEngineer"], fValue["Total"])
            to_plot[fleetID] = (fValue["Pilot"], fValue["CoPilot"], fValue["LoadMaster"], fValue["FlightEngineer"])
        plot_bar_graph(to_plot, crewClasses_with_total, day)

    os.chdir("..")


    with pd.ExcelWriter(excel_file_path) as writer:
        for xsheet in all_sheets.keys():
            all_sheets[xsheet].to_excel(writer, sheet_name=xsheet, index=False)

# Fuction for the ploting the graph

    def plot_crew_assignment():
        fig, ax = plt.subplots(
            figsize=(len(flightIds)//2, len(crew_members.keys())//5))

        ax.xaxis.grid(True, which='major', linestyle='-',
                      color='grey', alpha=.25)
        ax.yaxis.grid(True, which='major', linestyle='-',
                      color='grey', alpha=.25)

        for i, fId in enumerate(flightIds):
            for j, (crew, cValue) in enumerate(crew_members.items()):
                if sol.get_value(x[((fId, flights[fId][DAY_f]), crew)]) == 1:
                    clr = get_colour(crew[MEMBER_TYPE_cm])
                    ax.scatter(i, j, c=clr, marker='s', s=50)

        plt.xticks(ticks=range(len(flightIds)), labels=flightIds,
                   va='top', fontname=fontfamily)
        name_list = [crew[MEMBER_NAME_cm] for crew in crew_members.keys()]
        plt.yticks(ticks=range(len(name_list)), labels=name_list,
                   rotation=0, va='center', fontname=fontfamily)

        # Create legend with crew type patches
        crew_type_patch = [mpatches.Patch(color=get_colour(
            type), label=str(type)) for type in crewClasses]
        ax.legend(handles=crew_type_patch, bbox_to_anchor=(1.01, 1), loc='upper left',
                  borderaxespad=0., title="Crew Type", prop={'family': fontfamily})

        plt.title('Crew Assignment', fontname=fontfamily)
        plt.xlabel("Flight ID", fontname=fontfamily)
        plt.ylabel("Crew Members", fontname=fontfamily)

        # Display the plot
        # plt.show()
        plt.savefig("Crew_Assignment.png", dpi=500, bbox_inches='tight')

    def plot_crew_distribution(data):

        # Convert list of crew members to counts

        counts = {role: [len(data[x][role]) for x in data] for role in data[1]}

        # Create DataFrame from the counts

        df = pd.DataFrame(counts, index=data.keys())

        # Plot the data

        fig, ax = plt.subplots()

        bars = df.plot(kind='bar', stacked=True, ax=ax)

        # Add crew member names to the bars

        for height, j in enumerate(df.index):
            for i in range(len(df.columns)):

                for k in range(len(data[j][df.columns[i]])):

                    ax.text(height, sum(df.loc[j][:i]) + 0.5 + k,


                            data[j][df.columns[i]][k], ha='center', va='bottom', fontname=fontfamily)

            ax.text(height, sum(df.loc[j, :]), flights[j][FLEET_TYPE_f],
                    ha='center', va='bottom', fontname=fontfamily)
            # ax.text(height, sum(df.loc[j, :])+0.5, flights[j][FLEET_IDS_f], ha='center', va='bottom', fontname=fontfamily)

        plt.legend(loc='upper left', bbox_to_anchor=(1.01, 1),
                   borderaxespad=0., title="Crew Type", prop={'family': fontfamily})
        plt.xticks(rotation=0, fontname=fontfamily)
        plt.yticks(fontname=fontfamily)

        plt.title('Crew Members Distribution', fontname=fontfamily)

        plt.xlabel('Flight Number', fontname=fontfamily)

        plt.ylabel('Count', fontname=fontfamily)

        plt.savefig("Crew_Distribution.png", dpi=500, bbox_inches='tight')

    def plot_flight_distribution(data):
        # Create a subplot with specified size
        fig, ax = plt.subplots(figsize=(10, len(crew_members.keys())//5))

        assigned_crew = []
        i_ = 0

        # Iterate over crew data and their associated flight IDs
        for i, (crew, flightIDS) in enumerate(data.items()):
            # Get color based on crew type
            color = get_colour(crew[MEMBER_TYPE_cm])

            # Plot horizontal lines for each flight
            for flightID in flightIDS:
                start = mdates.date2num(get_time(flights[flightID][DEP_MIN_f]))
                end = mdates.date2num(get_time(flights[flightID][ARR_MIN_f]))

                ax.hlines(i_, start, end, colors=color, lw=10, label='Flight')
                ax.text((start + end) / 2, i_, flightID, ha='left',
                        va='center', color="white", fontsize=10, fontname=fontfamily)

            if len(flightIDS) != 0:
                assigned_crew.append(crew[MEMBER_NAME_cm])
                i_ += 1

        # Set yticks and labels
        ax.set_yticks(range(len(assigned_crew)))
        # y_index = [crew[MEMBER_NAME_cm] for crew in data.keys()]
        ax.set_yticklabels(assigned_crew, fontname=fontfamily)

        # Set up the formatting for the x-axis to show time
        ax.xaxis_date()
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=max(days)))
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
        plt.xticks(rotation=45., fontname=fontfamily)

        # Add light dotted horizontal lines
        ax.yaxis.grid(color='gray', linestyle='dotted', linewidth=0.5)
        ax.set_facecolor('whitesmoke')

        # Set labels for x and y axes
        plt.xlabel('Time', fontname=fontfamily)
        plt.ylabel('Crew Members', fontname=fontfamily)

        # Create legend with crew type patches
        crew_type_patch = [mpatches.Patch(color=get_colour(
            type), label=str(type)) for type in crewClasses]
        ax.legend(handles=crew_type_patch, bbox_to_anchor=(
            1.01, 1), loc='upper left', borderaxespad=0., title="Crew Type", prop={'family': fontfamily})

        # Add vertical lines at 00:00 of each day
        for day in range(min(days)-1, max(days) + 1):
            day_start = mdates.date2num(get_time(day * 24 * 60))
            ax.vlines(day_start, ymin=-1, ymax=len(assigned_crew),
                      colors='black', linestyles='dotted', linewidth=1)

        # Show the plot
        plt.title('Flight Distribution', fontname=fontfamily)
        plt.savefig("Flight_Distribution.png", dpi=500 , bbox_inches='tight')

    def plot_flight(data):
        # Create a subplot with specified size
        fig, ax = plt.subplots(figsize=(10, len(data.keys())//3))

        # Get unique fleet types from the data
        unique_fleet_types = set([f[FLEET_TYPE_f] for f in data.values()])

        # Generate a color palette map for fleet types
        color_palette = sns.color_palette(
            "hls", len(unique_fleet_types)).as_hex()
        fleet_color_map = dict(zip(unique_fleet_types, color_palette))

        i = 0
        # Iterate over crew data and their associated flight IDs
        for id, f in data.items():

            l1, l2 = f[LOCATION_f][1:-1].split('-')
            start = mdates.date2num(get_time(f[DEP_MIN_f]))
            end = mdates.date2num(get_time(f[ARR_MIN_f]))

            ax.hlines(
                i, start, end, colors=fleet_color_map[f[FLEET_TYPE_f]], lw=10, label='Flight')
            ax.text((start + end) / 2 - len(f[FLEET_IDS_f])/100, i, f[FLEET_IDS_f],
                    va='center', color="black", fontsize=10, fontname=fontfamily)
            ax.text(start - len(l1)/40, i, l1, va='center',
                    color="black", fontsize=10, fontname=fontfamily)
            ax.text(end, i, l2, va='center', color="black",
                    fontsize=10, fontname=fontfamily)
            i += 1

        # Set yticks and labels
        ax.set_yticks(range(len(data)))
        # y_index = [crew[MEMBER_NAME_cm] for crew in data.keys()]
        ax.set_yticklabels(data.keys(), fontname=fontfamily)

        # Set up the formatting for the x-axis to show time
        ax.xaxis_date()
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=max(days)))
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
        plt.xticks(rotation=45., fontname=fontfamily)

        # Add light dotted horizontal lines
        ax.yaxis.grid(color='gray', linestyle='dotted', linewidth=0.5)
        ax.set_facecolor('whitesmoke')

        # Set labels for x and y axes
        plt.xlabel('Time', fontname=fontfamily)
        plt.ylabel('Flight Numbers', fontname=fontfamily)

        # Create legend with crew type patches
        # crew_type_patch = [mpatches.Patch(color=get_colour(
        #     type), label=str(type)) for type in crewClasses]
        # ax.legend(handles=crew_type_patch, bbox_to_anchor=(
        #     1.01, 1), loc='upper left', borderaxespad=0., title="Crew Type", prop={'family': fontfamily})

        # Add vertical lines at 00:00 of each day
        for day in range(min(days)-1, max(days) + 1):
            day_start = mdates.date2num(get_time(day * 24 * 60))
            ax.vlines(day_start, ymin=-1, ymax=len(data),
                      colors='black', linestyles='dotted', linewidth=1)

        # Add a legend
        legend_elements = [mpatches.Patch(
            color=fleet_color_map[type], label=str(type)) for type in unique_fleet_types]
        ax.legend(handles=legend_elements, title='Fleet Type',
                  loc='center left', bbox_to_anchor=(1.01, 0.5))

        # Show the plot
        plt.title('Flight Distribution', fontname=fontfamily)
        plt.savefig("Flight_Distribution_Location.png", dpi=500 , bbox_inches='tight')


# calling the function for the ploting the graph
    plot_crew_distribution(crew_assignment)
    plot_flight_distribution(flight_assignment)
    plot_crew_assignment()
    plot_flight(flights)

    to_keep_in_next_input = []
    for crew, flights_for_this in flight_assignment.items():
        if flights_for_this:
            to_keep_in_next_input.append(crew)
    # to_keep_in_next_input

    sheet_to_edit = ["PILOT", "CO_PILOT", "LOAD_MASTER", "FLIGHT_ENGINEER"]
    sheet_to_edit_type = ["Pilot", "CoPilot", "LoadMaster", "FlightEngineer"]

    new_all_sheets = all_sheets.copy()

    for sheet_name, sheet_type in zip(sheet_to_edit, sheet_to_edit_type):
        this_sheet = all_sheets[sheet_name]
        # this_sheet = this_sheet.set_index("Name")

        # Create a new sheet
        new_sheet = pd.DataFrame(columns=this_sheet.columns)
        for i, row in this_sheet.iterrows():
            temp = (sheet_type, row["Name"])
            if temp in to_keep_in_next_input:
                # Add this row to the new sheet
                # Use loc to add the new row
                new_sheet.loc[len(new_sheet)] = row

        new_all_sheets[sheet_name] = new_sheet

    # Save the new sheets to a new Excel file
    with pd.ExcelWriter('new_crewscheduling.xlsx') as writer:
        for sheet_name, sheet_df in new_all_sheets.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Create a new DataFrame with the solution
    new_flight_df = []
    for fId in flightIds:

        new_flight_df.append([fId, flights[fId][LOCATION_f], flights[fId][FLEET_TYPE_f], flights[fId][FLEET_IDS_f], flights[fId][DEP_MIN_f], flights[fId][ARR_MIN_f], flights[fId][DAY_f],
                              flights[fId][FLIGHT_TIME_f], crew_assignment[fId]["Pilot"], crew_assignment[fId]["CoPilot"], crew_assignment[fId]["LoadMaster"], crew_assignment[fId]["FlightEngineer"]])
    df1 = pd.DataFrame(new_flight_df, columns=["Flight", "Location", "FleetType", "FleetIDs", "Dep_Min",
                                               "Arr_Min", "Day", "Flight_Time", "Pilot", "CoPilot", "LoadMaster", "FlightEngineer"])
    df1

    df2 = pd.DataFrame.from_dict(crew_members, orient="index", columns=[
        "ID", "Qualification", "HomeLocation", "AvailableHours", "FleetType"])
    df2["CrewName"] = df2.apply(
        lambda row: row.name[1], axis=1)
    df2["CrewType"] = df2.apply(
        lambda row: row.name[0], axis=1)
    df2["Flights"] = df2.apply(
        lambda row: flight_assignment[row.name], axis=1)
    df2["FlightCount"] = df2.apply(
        lambda row: len(row["Flights"]), axis=1)
    df2["FlightDuration"] = df2.apply(
        lambda row: "{}hrs {}mins".format(sum(flights[fId][FLIGHT_TIME_f] for fId in row["Flights"])//60, sum(flights[fId][FLIGHT_TIME_f] for fId in row["Flights"]) % 60), axis=1)

    df2 = df2[["CrewName", "CrewType", "ID", "Qualification", "HomeLocation",
               "AvailableHours", "FleetType", "Flights", "FlightCount", "FlightDuration"]]

    df3 = pd.DataFrame.from_dict(sol_info, orient="index", columns=["Value"])
    df3["Key"] = df3.apply(
        lambda row: row.name, axis=1)
    df3 = df3[["Key", "Value"]]

    # Save the DataFrames to Excel file in different sheets
    with pd.ExcelWriter('output.xlsx') as writer:
        df1.to_excel(writer, sheet_name='FLIGHT_INFO', index=False)
        df2.to_excel(writer, sheet_name='CREW_INFO', index=False)
        df3.to_excel(writer, sheet_name='SOLVER_INFO', index=False)

    # no_of_crew_used.append((iteration, len(crew_members)))

# Adding the data in dictionary for the ploting
    crew_members_dict.append((iteration, len(crew_members)))
    pilot_members_dict.append((iteration, len(pilots_crews)))
    copilot_members_dict.append((iteration, len(copilot_crews)))
    loadmaster_members_dict.append((iteration, len(loadmaster_crews)))
    flight_engineer_members_dict.append(
        (iteration, len(flight_engineer_crews)))

    if len(crew_members) == len(to_keep_in_next_input):
        print("All crew members are assigned")
        os.chdir("..")
        return
    else:
        print("Some crew members are not assigned")
        print("Number of crew members not assigned: ", len(
            crew_members) - len(to_keep_in_next_input))
        Allocation_crew_member_in_fights(
            "new_crewscheduling.xlsx", Output_folder_name, iteration + 1)

    os.chdir("..")


if __name__ == "__main__":

    # is_shuflfled = False
    Allocation_crew_member_in_fights(excel_file_path, Output_folder_name, 1)

    # print("Number of crew used in each iteration: ", no_of_crew_used)

    crew_members_dict.append((crew_members_dict[-1][0] + 1, crew_members_dict[-1][1]))
    pilot_members_dict.append((pilot_members_dict[-1][0] + 1, pilot_members_dict[-1][1]))
    copilot_members_dict.append((copilot_members_dict[-1][0] + 1, copilot_members_dict[-1][1]))
    loadmaster_members_dict.append((loadmaster_members_dict[-1][0] + 1, loadmaster_members_dict[-1][1]))
    flight_engineer_members_dict.append((flight_engineer_members_dict[-1][0] + 1, flight_engineer_members_dict[-1][1]))


    print(crew_members_dict)
    plot_data(crew_members_dict)
    plot_data_all(pilot_members_dict, copilot_members_dict, loadmaster_members_dict,
                  flight_engineer_members_dict)
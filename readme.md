
# Crew Scheduling Optimization for Airlines

This project provides a Python-based solution for the airline crew scheduling problem. It utilizes a Constraint Programming (CP) model to assign crew members (Pilots, Co-pilots, Load Masters, Flight Engineers) to a given set of flights. The primary objective is to create a valid and feasible schedule while minimizing the total number of crew members utilized.

The script reads flight and crew data from an Excel file, applies a set of complex operational constraints, and generates a detailed assignment schedule along with insightful visualizations.

***

## 📋 Features

-   **Data-Driven:** All flight schedules, crew details, and operational constraints are imported from a structured Excel file.
-   **Constraint-Based Optimization:** Built using the powerful `docplex` library from IBM to model and solve the complex crew pairing problem.
-   **Comprehensive Constraint Handling:** The model respects a wide range of rules, including:
    -   Crew qualifications and aircraft type compatibility.
    -   Fleet-type restrictions for pilots and co-pilots.
    -   Maximum daily flying hours and landings.
    -   Minimum sit time (between flights in the same duty) and layover time (overnight rest).
    -   Total available work hours for each crew member.
-   **Optimization Goal:** The model is configured to find a solution that minimizes the total number of unique crew members required to cover all flights.
-   **Iterative Solving:** If a solution cannot be found for all flights with the initial crew pool, the script can run iteratively, removing unassigned crew members to solve for a smaller, feasible subset.
-   **Rich Visualization:** Automatically generates a variety of plots and charts for easy analysis of the results:
    -   Gantt charts for flight assignments per crew member.
    -   Stacked bar charts showing crew composition per flight.
    -   Heatmaps for crew qualifications.
    -   Summary plots tracking optimization progress.
-   **Detailed Reporting:** Exports the final schedule, crew utilization stats, and solver information into a clean Excel report.

***

## ⚙️ How It Works

The script follows a systematic process to achieve the optimal crew schedule:

1.  **Data Loading:** It begins by loading all data from the specified Excel file (`cp for crewscheduling.xlsx`) into pandas DataFrames. This includes flight schedules, crew rosters for different roles, aircraft requirements, and operational rules.
2.  **Data Preprocessing:** Timestamps are converted into minutes for easier calculations. Data is structured into dictionaries and sets to be fed into the optimization model.
3.  **Model Formulation:** A Constraint Programming model is built using `docplex.cp.model`.
    -   **Decision Variables:** A binary variable `x` is created for each possible pairing of a flight and a crew member. `x(flight, crew) = 1` if the crew member is assigned to the flight, and `0` otherwise.
    -   **Constraints:** The core operational rules (e.g., a crew member cannot be on two overlapping flights, daily flying hour limits must be respected) are added to the model as mathematical constraints.
    -   **Objective Function:** The model is instructed to minimize the sum of all unique crew members assigned to at least one flight.
4.  **Solving:** The CP-Optimizer solver is invoked to find a feasible and optimal solution that satisfies all constraints.
5.  **Output Generation:** If a solution is found, the script processes the results and generates:
    -   An `OUTPUT` folder containing all visualizations as `.png` files.
    -   An `output.xlsx` file with three sheets:
        -   `FLIGHT_INFO`: The flight schedule with assigned crew names.
        -   `CREW_INFO`: Detailed statistics for each crew member (flights assigned, total duration).
        -   `SOLVER_INFO`: Performance metrics from the solver.
    -   A `new_crewscheduling.xlsx` file, which serves as the input for the next iteration if needed.

***

## 🛠️ Prerequisites

Before you run the script, ensure you have Python 3.x installed. Then, install the required libraries using pip:

```bash
pip install pandas openpyxl seaborn matplotlib numpy docplex
````

-----

## 🚀 How to Run

1.  **Clone the repository:**

    ```bash
    git clone <your-repository-url>
    cd <repository-folder>
    ```

2.  **Prepare the Input Data:**

      - Create an Excel file named `cp for crewscheduling.xlsx` in the same directory.
      - Populate it with the required sheets and data as described in the **Input File Structure** section below.

3.  **Configure the Script (Optional):**

      - Open the Python script and modify the global variables at the top if needed:
          - `excel_file_path`: Change if your Excel file has a different name.
          - `is_shiffled`: Set to `True` to randomize the order of flights/crew, which can help find solutions faster in some cases.
          - `to_optimize`: Set to `True` to minimize crew count; `False` just finds any feasible solution.

4.  **Execute the Script:**

      - Run the script from your terminal:

    <!-- end list -->

    ```bash
    python your_script_name.py
    ```

5.  **Check the Results:**

      - A new folder named `OUTPUT1` will be created. If the script runs iteratively, subsequent folders `OUTPUT2`, etc., will be created for each run. Inside, you will find all the generated plots and the `output.xlsx` report.

-----

## 📄 Input File Structure

The `cp for crewscheduling.xlsx` file must contain the following sheets with the specified columns:

  - **`FLIGHT_INFO`**: Details of each flight.

      - `Flight`: Unique identifier for the flight (e.g., 1, 2, 3).
      - `Location`: Origin-Destination pair (e.g., `(A-B)`).
      - `FleetType`: Type of aircraft (e.g., `TypeA`).
      - `FleetIDs`: Specific aircraft tail number (e.g., `AC1`).
      - `Dep.`: Departure time in `HH:MM` format.
      - `Arr.`: Arrival time in `HH:MM` format.
      - `Day`: Day of the operation (e.g., 1, 2).

  - **`CREW_LIMITATIONS`**: Operational rules for each crew role.

      - `crewtype`: Role (e.g., `Pilot`, `CoPilot`).
      - `DailyMaxFlyingHours`, `DutyDay`, `RestPriod`, `MinSitTime`, `MinLayover`, `DailyMaxLandings`.

  - **`PILOT`**, **`CO_PILOT`**, **`LOAD_MASTER`**, **`FLIGHT_ENGINEER`**: Separate sheets for each crew role.

      - `ID`: Unique ID for the crew member.
      - `Name`: Name of the crew member.
      - `Qualification`: Skill level (e.g., `A`, `B`, `C`).
      - `HomeLocation`: Base station.
      - `AvailableHours`: Total hours available for the schedule period.
      - `FleetType`: (For Pilot/Co-Pilot) The specific fleet type they are qualified for.

  - **`AC`**: Aircraft specifications.

      - `AircraftVariant1`: Fleet type name matching `FLIGHT_INFO`.
      - `Pilot`, `CoPilot`, ...: Number of crew members of each type required.
      - `QualificationPilot`, ...: Minimum qualification level required for each role on this aircraft.

-----

## 📊 Output Description

  - **Folders (`OUTPUT1`, `OUTPUT2`, ...):** Each folder represents one successful run of the optimizer.
  - **Plots (.png):**
      - `Crew_Assignment.png`: A grid showing which crew is on which flight.
      - `Crew_Distribution.png`: A bar chart showing the number and names of crew on each flight.
      - `Flight_Distribution.png`: A Gantt chart visualizing each crew member's schedule.
      - `Flight_Distribution_Location.png`: A Gantt chart visualizing the schedule of all flights.
      - `DayX.png`: Bar charts summarizing daily crew usage by fleet.
  - **Excel Reports:**
      - `output.xlsx`: The final, human-readable schedule and summary.
      - `new_crewscheduling.xlsx`: An intermediate file containing only the assigned crew, used for subsequent iterations.

<!-- end list -->


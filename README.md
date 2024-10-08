# UFC Data Visualization Project (EXCEL)

## Overview
The **UFC Data Visualization Project** is an analytical dashboard created in Microsoft Excel that visualizes various statistics for UFC fighters. This project leverages PivotTables, interactive charts, and advanced Excel functionalities to present fighter data in an intuitive and engaging manner. The dashboard allows users to explore and analyze different fighter attributes, win/loss statistics, striking accuracy, and other key performance metrics, providing insights into trends and patterns across various weight classes.

The project is fully documented and organized into a GitHub repository, complete with VBA scripts, Excel files, and supplementary images. It is designed to be an interactive tool for UFC enthusiasts, data analysts, or anyone interested in exploring fighter statistics in a visual format.

![0E9B5F19-1F3E-481B-A1A8-4C4E790FE699](https://github.com/user-attachments/assets/da17ac8b-6ebf-42e3-9e58-4896c6394e30)

## Project Features

### 1. Interactive Dashboards
The dashboards are created using Excel's interactive elements, such as slicers, to filter data dynamically by weight class, gender, and fighter attributes. Key statistics such as `Total Wins`, `KO Wins`, and `Win/Loss Ratios` are presented in a visual format with bar charts, radar charts, and scatterplots.

- **Top 10 and Bottom 10 Fighters**: The dashboard allows users to view the top 10 and bottom 10 fighters by rank within each weight class.
- **Detailed Fighter Profiles**: Each fighter's profile can be explored with information on weight, height, reach, stance, and performance metrics.

![87091A63-BDBB-4EAF-AB09-9C922898DB1C](https://github.com/user-attachments/assets/be936e78-42b6-41e0-9b9d-3792ffede47d)

### 2. Data Visualizations
Several visualizations are used to highlight fighter data:

- **Bar Charts**: Show the number of wins, KO wins, and total matches by weight class.
- **Radar (Spider) Charts**: Visualize fighter stances and attributes like reach and height.
- **Scatterplots**: Compare variables such as reach vs. wins and age vs. performance metrics.
- **Heat Maps**: Present reach, weight, and other attributes as heat maps for easy comparison.

### 3. Data Analysis Using PivotTables
PivotTables are used extensively to analyze fighter statistics:

- **Win/Loss Analysis**: Breakdown of wins and losses across different weight classes and gender.
- **Weight Class vs. Performance Metrics**: Visualize performance metrics like `Striking Accuracy` and `Defense` based on different weight classes.

### 4. Fighter Images Integration
Each fighter's image is displayed on the dashboard, changing dynamically based on the slicer selection. This enhances the visual appeal and provides a quick reference for identifying fighters.

![B4C9799F-A85F-4729-837F-5E43F61157B0](https://github.com/user-attachments/assets/aebdc6dd-3862-41f9-9698-6bd4eff881c5)

### 5. VBA Scripts for Advanced Functionality
The project includes VBA scripts for automating certain tasks, such as dynamically updating charts or images based on selections. These scripts are included in the repository for advanced users who want to explore or modify the code.

## Project Structure
The repository is organized as follows:

- **`Fighter_Attributes.xlsx`**: Contains detailed attributes for each fighter, including weight, height, reach, stance, and performance metrics like wins, losses, KO wins, and submission wins.
- **`UFC_Fighter_Stats.xlsx`**: Aggregated data on fighter statistics, categorized by weight class and gender.
- **`UFC_Dashboard_v1.xlsx`**: Main dashboard file with slicers, dynamic charts, and interactive elements.
- **`UFC_Dashboard_v2.xlsx`**: Enhanced version of the dashboard with additional functionality and visualizations.
- **`Data_Processing_Scripts.ipynb`**: Python Jupyter notebook for processing the raw data before importing into Excel.
- **`Excel_Chart_Macros.bas`**: VBA script file to automate chart updates and dynamic interactions.

## Usage Instructions

### Step 1: Open the Dashboard
1. Download and open `UFC_Dashboard_v1.xlsx` in Microsoft Excel.
2. Ensure that macros are enabled to use advanced functionalities like dynamic chart updates.

### Step 2: Interact with the Dashboard
1. Use the **weight class slicers** on the left to filter fighters by weight class.
2. View detailed fighter attributes and metrics in the main dashboard area.
3. Click on any fighter’s name or weight class to see their profile and key statistics.

### Step 3: Explore Visualizations
1. Navigate through different sheets to explore scatterplots, heat maps, and bar charts.
2. Use slicers to switch between different metrics and weight classes.

### Step 4: Modify or Extend the Project
1. Use the VBA scripts provided in the `Scripts` folder to extend the functionality.
2. Modify the data in `Fighter_Attributes.xlsx` to include new fighters or update statistics.

## Contributing
Contributions to this project are welcome. Please feel free to open issues or submit pull requests to add new features, fix bugs, or improve the documentation.

## Screenshots
### Main Dashboard Overview
![0E9B5F19-1F3E-481B-A1A8-4C4E790FE699](https://github.com/user-attachments/assets/fbcf14f4-6ab2-4dbb-b179-434bfa81124c)

### Fighter Attributes Dashboard
![1ADB0CC9-CBE5-4F1F-B335-95E6C23FC813](https://github.com/user-attachments/assets/a519dc5b-7484-4ecb-a2ba-a215236d60a2)

### Fighter Images
![B4C9799F-A85F-4729-837F-5E43F61157B0](https://github.com/user-attachments/assets/83d47f43-6004-4a57-988b-9c5104efafb2)

## License
This project is licensed under the MIT License. See the `LICENSE` file for more information.

## Contact
For questions or feedback, please reach out through the Issues section or submit a pull request for contributions.

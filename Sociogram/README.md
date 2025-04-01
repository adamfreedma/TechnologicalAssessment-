
# Sociogram 

## Overview
This project analyzes social structures based on directed graphs constructed from survey data. The goal is to extract meaningful insights into social connections and group dynamics among cadets using sociometric and demographic analysis tools.

## Features
- **Graph-Based Sociometric Analysis**: Constructs a directed graph from cadet survey responses to identify friendships and social connections.
- **Clique Detection and Visualization**: Identifies social cliques using graph theory methods and generates visualizations.
- **Demographic Group Analysis**: Examines connections within specific demographic groups.
- **Popularity Metrics**: Tracks individual popularity over time based on sociometric data.
- **Personalized Sociometric Reports**: Generates individualized reports for cadets, summarizing their social standing.
- **Sociogram Utility Functions**: Provides helper methods for sorting, analyzing, and visualizing social data.

## Project Structure
```
Sociogram/
│── main_sociogram.py        # Entry point for the analysis pipeline
│── sociogram.py             # Core module for sociometric analysis
│── sociogram_utils.py       # Utility functions for data processing
│── clique_slides.py        # Clique detection and visualization
│── demographic_maps.py      # Generates demographic-based social maps
│── private_slides.py       # Personalized sociometric slides for individuals
│── popularity.py           # Tracks popularity metrics over time
```

## Dependencies
The project requires the following Python libraries:
- `pandas`
- `networkx`
- `matplotlib`
- `numpy`
- `plotly`
- `openpyxl`

## Running the Project
1. **Prepare Data**: Ensure that you have all required input files - i.e the sociogram of current semester and the last on on your computer.
2. **Data Validation** Ensure manually that the input data is in the correct format (for example that no one filled a name instead of a number as one of their choices)
2. **Update Constants** Update the following constant :
   * `FILE_PATHS`- Place the paths of the input files in the  constant in cronological order
   * `FILE_OUT`- The path to the desired output directory
   * `SEMESTER_INDEX` - The index of the current semester of the data in the FILE_PATHS contstant
   * `IDENTIFIER` name of the column containing the ID of the cadets
   * `NAME` name of the column containing the name of the cadets
   * `NUM` name of the column containing the corresponding number of the cadets
3. **Run Main Analysis**:
   ```sh
   python main_sociogram.py
   ```

## Output
- **Graph Visualizations**: Group structures, social cliques, and demographic maps.
- **Individual Reports**: Personalized analysis for each cadet.
- **Excel Reports**: Summary statistics of social connections.

## Contributors
- Developed as part of the Talpiot evaluation system for cadet assessment.
- First developed by Iftah Farkash from מ"ד
- Further developement was made by Ben Hillel from מ"ד and Ohad Levin from ל"ט


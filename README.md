# Monte-Carlo-Risk
This project makes use of Monte Carlo Simulation with Triangular or PERT distribution to give better estimates on risks

Information about PERT/Triangular distribution:
- https://www.riskamp.com/beta-pert

Information about Monte Carlo Simulation:
- https://www.ibm.com/topics/monte-carlo-simulation


# Objective:
- This project uses monte carlo simulation with Triangular or PERT distribution to produce estimates according to percentage confidence levels
- Value range can be currency or hours
- Once the application runs an Excel file is created with the estimates + charts.
  - Additionally, the image files found in the Excel file are also create

# TO-DO list
- (Optional) Images are created and stored at the folder. May need to remove that depending on what people think of
- Create VBA code that calls the python code (https://stackoverflow.com/questions/68312525/run-python-script-through-excel-vba)
- Once VBA code finishes, open the copy

# Problems: 
- None, so far

# Limitations:
- Openpyxl cannot read values from macros, so avoid using it when manipulating values. Otherwise an error message will be thrown


# Libs: 

This project uses the following libs: openpyxl, numpy, pyplot, seaborn, pert

For PERT import, run in project terminal (PyCharm):
- pip3 install pertdist 
  - Package information: https://pypi.org/project/pertdist/
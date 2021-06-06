# Gantt-Chart-generator
Given a custom table, this file is meant to generate:
- A Gantt Chart
- A completed the table with the length of the tasks (in months)

This script is meant to be used with Python3

### Input files required:
#### 1. GanttChart_input
csv format

An example input table is provided

Date format: mm/dd/yyyy

As convention, it is recommended to insert the dates as last day of the month


#### 2. Run python script:
Two parameters are required: 
- the file that contains the input table 
- the output file name

For instance:
```
python3 GanttChart_Github.py -f /path/input.csv -o Gantt.xlxs
```

Please cite if used and get in contact with comments and suggestions
Cheers!

# Nextech Stats Showdown 2022

## Instructions
The algorithm is implemented with Python and the openpyxl library. In order to run the algorithm, replace <PLACEHOLDER_WORKBOOK_NAME> and <PLACEHOLDER_WORKSHEET_NAME> in the script with references to the appropriate workbook name and worksheet name. Run the script (Python and openpyxl have to be installed) and the program will print out the top ten teams on the console. 

## Explanation
1. Data for each team is taken into a nested list, which includes values from each column listed as a “factor” in the table below
2. Each value, if not already between 0 and 1, is normalized with the formula (val - min) / (min - max), with the min and max representing the minimum and maximum values in the column 
3. Values from each column of a team is multiplied by a corresponding predetermined weight, as listed in the table below
4. Each team is assigned a score, which is the sum of weighted values in all its columns
5. Teams are sorted in descending order according to their score and the top ten teams are printed

| Factor                        | Weight  |
|-------------------------------|---------|
|Conference Tournament Champion |-0.17    |
|Made Tournament Previous Year  |0.3      |
|Wins                           |0.37     |
|Losses                         |-0.33    |
|Assists                        |0.22     |
|Assist to Turnover Ratio       |0.28     |
|ESPN Strength of Schedule      |-0.38    |
|Quad 1 Wins                    |0.45     |
|Total Scoring Differential     |0.44     |
|Scoring Differential Per Game  |0.43     |
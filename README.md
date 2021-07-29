# Tables
working with xlsx files from PyCharm

In Tables.py there is a code which imports openpyxl directory in order to process the file Population_in_millions.xlsx.
In Population_in_millions.xlsx there is the table which contains number of people (in millions) of particular religion on all continents.
Program takes that file and converts numbers of people in millions to percentages.
At the end, program saves that percentages to the new file titled Population_per_percentage.xlsx.

In Charts1.py program imports openpyxl again, in order to continue working with file Population_per_percentage.xlsx.
It uses data from the table (percentage of followers of each religion by continent) to form a chart for each continent.
It also adds a title (name of the continent) to each chart.
Charts can be seen in Charts.xlsx in which I saved it.
Labels (Christians, Muslims, etc.) are added in xlsx file.
If someone knows how to add it from PyCharm, fell free to fulfill this code with that feature :)

In File.py the codes are united inside of the method create_charts, so any file which contains a table can be processed.
We only have to call the method create_charts(file), where file is a name of an xlsx file.
Program converts numerical data from any table in xlsx file to percentages, than plots the charts for every column. 

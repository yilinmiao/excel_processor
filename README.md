# Excel Processor
Develop a script (Python preferred) to convert the data in the attached spreadsheet.

Input is the data in the first tab, and produce the output in the second tab based on the example format.

For example, in the first tab, we have the following:

|Kevin, A	|CMPE	|110	|1	   MW	1200	1250
|---------|-----|-----|:-----------------
||||2	   T	1330	1620
||||3	   T	1630	1920
||||200	1	   W	1500	1745
||||240	1	   M	1500	1745

The output of his classes to look like the following:

|Kevin, A	|CMPE 110-01, MW, 12:00-12:50 	|CMPE 110-02, T, 13:30-16:20	|CMPE 110-03, T, 16:30-19:20	|CMPE 200-01, W, 15:00-17:45
|---------|-------------------------------|-----------------------------|-----------------------------|----------------------------

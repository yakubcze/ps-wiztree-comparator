# wiztree-comparator
- A very simple script to automatically scan a selected folder, export to .csv and compare it to previous version of the scan

- This script utilizes WizTree and [WizTreeCompare](https://github.com/AlphaDelta/WizTreeCompare). Script scans specified folder using WizTree, exports the result to .csv file to a specific folder. Then WizTreeCompare is launched, which takes the newest .csv file + the 2nd newest from the folder and compares them. The result is again saved in a .csv file in a different specified folder. Folder scan can be interrupted using Ctrl+C. 

- Script is launched using the supplied "Run_WizTreeComparator.bat" file, which ensures the script is launched with admin rights.

- !WizTree must be set to english language, or else the columns in .csv file will be incorrectly named.!
- !No other instance of WizTree can be launched at the same time, because the script waits for WizTree to exit so it can start WizTreeCompare!

## Parameters
- $wtPath = path to the WizTree .exe
- $wtcPath = path to the WizTreeCompare .exe
- $scanPath = path to scan

- $csvFilePath_wt = path to which .csv files from WizTree are exported
- $csvFilePath_wtc = path to which .csv files from WizTreeCompare are exported

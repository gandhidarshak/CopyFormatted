## Copy-Formatted

At work, I often find myself copying excel data to textual table/csv/yaml files. Earlier, my approach was to [1] Copy the data of my interest to a new excel file [2] Saving that excel as csv [3] Using gvim to edit/align the data to get what I want. One day, I got bored and wrote few VBA macros to do these in a single click, with better formatting and accuracy.

## Demo

![Demo CopyFormatted](https://github.com/gandhidarshak/CopyFormatted/blob/master/Demo/CopyFormattedDemo.gif)

## How to install?

### Download the plugin file
Download or git-clone the CopyFormatted.xlam file. Keep it in a location which would not change that often.

### Add Plugin to Excel 
1. Make sure you have Developer's tab visible in Excel. If not visible, On the File tab, go to Options > Customize Ribbon. Under Customize the Ribbon and under Main Tabs, select the Developer check box.
2. Open an Excel workbook and Go to Developer –> Add-ins –> Excel Add-ins
3. In the Add-ins dialogue box, browse and locate the CopyFormatted.xlam file that you saved, and click OK.
At this point CopyAsTable, CopyAsCSV, CopyAsYaml Macros are added to your excel.

### Add Macros to Ribbon
1. Right-click on any of the ribbon tabs and select Customize the Ribbon
2. Create a New Tab (MyMacros in the Demo above) or use some existing tab depending on where you want the buttons to show up.
3. Create a New Group (Copy Formatted Text in the Demo above). Make sure the tab and groups are visible by selecting their check boxes.
4. In Choose commands from dropdown list, pick Macros. 
5. Three CopyAs* macros should be visible there. Add them to the group created above. Rename and change icons if preferred. 
At this point, you should have the three ribbon buttons set-up and ready to use. Enjoy!

### User-defined type error pop-up
I am using MS Forms Object library to copy data to Windows clipboard. If you get a user-defined type missing error pop-up, you may need to enable that library in your VBA editor (Alt+F11). Once in VBA editor, click Tools > References and check the box next to “Microsoft Forms 2.0 Object Library.”


## Want something more?

If you want the excel data to be copied to any other format than table/csv/yaml, please shoot me with specifics of the format/use-case and I would be happy to incorporate them or guide you with pointers on how to do it on your own. 

## License

CopyFormatted uses the MIT license. See [LICENSE](https://github.com/gandhidarshak/CopyFormatted/blob/master/LICENSE) for more details.

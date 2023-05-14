# RPA challenge in Power Automate For Desktop

Solution for the RPA challenge created in MS Power Automate for Desktop.
I am sharing only the screenshots of the flow created for the attended solution.
___
The structure of the repository includes only the .png files (and that one readme). Here is a little explanation what is done where:

  * _main.png_ - the main flow of the project. The _TimeoutVar_ is used in error handling as a timeout value (expressed in seconds).
  ![main](https://github.com/CzadowyChlopak/RPAchallengeInPowerAutomateForDesktop/blob/main/main.png)
  
  * _kill_excel.png_ - kills the whole Excel process. Mainly used to close the Excel instance if I forgot to do it while running the whole flow. If you want to recreate my solution please keep in mind to save every Excel Workbook before runing this subflow.
  ![main](https://github.com/CzadowyChlopak/RPAchallengeInPowerAutomateForDesktop/blob/main/kill_excel.png)
  
  * _GetLocalDirectory.png_ - checks and creates the folder path where the project files will be deployed.
  ![main](https://github.com/CzadowyChlopak/RPAchallengeInPowerAutomateForDesktop/blob/main/GetLocalDirectory.png)
  
  * _DownloadAndLoadFile.png_ - downloads the RPA challenge Excel file and loads it into the memory.
  ![main](https://github.com/CzadowyChlopak/RPAchallengeInPowerAutomateForDesktop/blob/main/DownloadAndLoadFile.png)
  
  * _ExcelCheckEmpty.png_ - runs in the DowloadAndLoadFile subflow. It checks in Excel file one cell after the another if it is empty. If the cell is empty it deletes the whole row.
  ![main](https://github.com/CzadowyChlopak/RPAchallengeInPowerAutomateForDesktop/blob/main/ExcelCheckEmptyCells.png)
  
  * _FillTheWebpage.png_ - browser automation for filling the text boxes and clicking the buttons during the RPA challenge
  ![main](https://github.com/CzadowyChlopak/RPAchallengeInPowerAutomateForDesktop/blob/main/FillTheWebpage.png)
  
  * _UI_elements.png_ - list of the elements stored in memory  
  * _Variables.png_ - list of variables used in the flow  
  * _Output.png_ - final screenshot of the web browser after completing the challenge.

___
I would also like to describe how to properly refer to the web elements. In the RPA challenge the id's are generated randomly for each refreshment. As default Power Automat for Desktop refers to the textboxes by the property. To solve that issue I would like to recommend search for the input in the parent div.

To do this after recording the text field as usual we have to edit its value. Go to the *Default Selector* of the recorded element and switch into the text editor mode. Here is the prompt for the *Address* field:

`div:contains('Address') > input`


___
The other thing that I want to discuss is the subflow called *ExcelCheckEmpty*.

It is created in the purpose of checking the downloaded file in case of blank cells. There are sort of actions that can handle that kind of error - in my case the row with the blank cell is deleted. This feature is not necessary in the RPA chellenge soltuion, but in my opinion it is important error handling part during the bot development. 

The algorithm to check those blank is quite simple - it is iterating thorugh each column from the last row till the 1st row of the data (which is 2nd of the sheet - 1st row of the sheet includes the header). Iterating through the last row allows the algorithm to delete a row without missing the data. If the row is deleted also the value which stores the last row index is decreased.

Checking those blanks by the Excel is quite time cosuming - the faster way to do it would be load a whole data table into the memory and computer those operations in the memory. However during development I was focused on showcasing sort of the software technical knowledge and my thought was that in that exactly case it would be better to show that Power Automate for Desktop can also do some *Excel-VBA-Macro* like stuff.

___

As final result the bot filled the form with 100% accuracy with a little bit more than 2 minutes time.
![main](https://github.com/CzadowyChlopak/RPAchallengeInPowerAutomateForDesktop/blob/main/Output.png)

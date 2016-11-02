cd\
cd TomH\Code\FDLD
:: -----------------------------------------------------------------------------------------------
:: Note - Please use :: to comment any line. To uncomment, please remove :: from start of the line
:: -----------------------------------------------------------------------------------------------

::Set Preferences Full RF=Y
::FD-Regression.txt Scenario 42a: Set Specific Preferences - Full RF=Y
Call pybot -i TC42a --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 34 Validate UI Leads Summary Apps
Call pybot -i TC34 --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

PAUSE

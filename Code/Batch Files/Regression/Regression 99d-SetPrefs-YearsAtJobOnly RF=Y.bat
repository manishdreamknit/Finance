cd\
cd TomH\Code\FDLD
:: -----------------------------------------------------------------------------------------------
:: Note - Please use :: to comment any line. To uncomment, please remove :: from start of the line
:: -----------------------------------------------------------------------------------------------

::Set Preference YearsAtJob Only
::FD-Regression.txt Scenario 42c: Set Specific Preferences - YearsAtJob Only - RF=Y
Call pybot -i TC42c --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

Exit

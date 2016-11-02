cd\
cd TomH\Code\FDLD
:: -----------------------------------------------------------------------------------------------
:: Note - Please use :: to comment any line. To uncomment, please remove :: from start of the line
:: -----------------------------------------------------------------------------------------------

::Set Preference YearsAtJob Only - RF=N
::FD-Regression.txt Scenario 42d: Set Specific Preferences - YearsAtJob Only - RF=N
Call pybot -i TC42d --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt TC06 GetLead Audit
Call pybot -i TC06 --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

PAUSE

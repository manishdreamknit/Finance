cd\
cd TomH\Code\FDLD
:: -----------------------------------------------------------------------------------------------
:: Note - Please use :: to comment any line. To uncomment, please remove :: from start of the line
:: -----------------------------------------------------------------------------------------------

::::::::::::::::::::::::::::::::::::::::::::::::::::
::Remove any Routing Preferences prior to executing
::::::::::::::::::::::::::::::::::::::::::::::::::::

::Set Preference YearsAtJob Only - RF=N
::FD-Regression.txt Scenario 42d: Set Specific Preferences - YearsAtJob Only - RF=N
Call pybot -i TC42d --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 43b Unset CRM Preferences
Call pybot -i TC43b --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 17: Global Update
Call pybot -i TC17 --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

PAUSE

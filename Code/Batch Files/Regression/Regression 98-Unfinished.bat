cd\
cd TomH\Code\FDLD
:: -----------------------------------------------------------------------------------------------
:: Note - Please use :: to comment any line. To uncomment, please remove :: from start of the line
:: -----------------------------------------------------------------------------------------------

::Set Preference YearsAtJob Only
::FD-Regression.txt Scenario 42d: Set Specific Preferences - YearsAtJob Only - RF=N
Call pybot -i TC42d --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 09: Decisions
Call pybot -i TC09 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 15: Service Errors
Call pybot -i TC15 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::Restore Full Preferences
::FD-Regression.txt Scenario 42a: Set Specific Preferences - Full RF=Y
Call pybot -i TC42a --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

Exit

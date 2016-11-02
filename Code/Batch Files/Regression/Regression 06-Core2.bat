cd\
cd TomH\Code\FDLD
:: -----------------------------------------------------------------------------------------------
:: Note - Please use :: to comment any line. To uncomment, please remove :: from start of the line
:: -----------------------------------------------------------------------------------------------

:: See note for Scenario 36

::Set Preference YearsAtJob Only
::FD-Regression.txt Scenario 42d: Set Specific Preferences - YearsAtJob Only - RF=N
Call pybot -i TC42d --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 11: Community Property
Call pybot -i TC11 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::Set Notifications & Messaging\CRM Routing so that: AJK, DXA, DXB, DXC can all receive leads created by DWR.
::FD-Regression.txt Scenario 36: Routed Leads Iterations
Call pybot -i TC36 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::Restore Full Preferences
::FD-Regression.txt Scenario 42a: Set Specific Preferences - Full RF=Y
Call pybot -i TC42a --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

Exit

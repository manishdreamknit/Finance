cd\
cd TomH\Code\FDLD
:: -----------------------------------------------------------------------------------------------
:: Note - Please use :: to comment any line. To uncomment, please remove :: from start of the line
:: -----------------------------------------------------------------------------------------------

::Set Preference YearsAtJob Only
::FD-Regression.txt Scenario 42d: Set Specific Preferences - YearsAtJob Only - RF=N
Call pybot -i TC42d --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 13: Enum Iterations - Employment, Housing, Suffix
Call pybot -i TC13 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 14: Zip Code Ranges - App
Call pybot -i TC14 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 18: Relationship
Call pybot -i TC18 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 19: Email Address Validation
Call pybot -i TC19 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 20: Phone 800-Series
Call pybot -i TC20 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 26: Endpoints D2I
Call pybot -i TC26 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::These cause unhandled 'Unauthorized Access' errors and can not presently be handled in Automation.
::FD-Regression.txt Scenario 16: Authentication
::Call pybot -i TC16 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::Restore Full Preferences
::FD-Regression.txt Scenario 42a: Set Specific Preferences - Full RF=Y
Call pybot -i TC42a --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

Exit

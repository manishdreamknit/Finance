cd\
cd TomH\Code\FDLD
:: -----------------------------------------------------------------------------------------------
:: Note - Please use :: to comment any line. To uncomment, please remove :: from start of the line
:: -----------------------------------------------------------------------------------------------

::Set Preferences Full
::FD-Regression.txt Scenario 42a: Set Specific Preferences - Full RF=Y
Call pybot -i TC42a --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 02: Check Services
Call pybot -i TC02 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 03: Random Data
Call pybot -i TC03 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 40: GetLead Service
Call pybot -i TC40 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 41: DDF Service
Call pybot -i TC41 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 07: Validate the FD Process
Call pybot -i TC07 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

Exit

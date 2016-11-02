cd\
cd TomH\Code\FDLD
:: -----------------------------------------------------------------------------------------------
:: Note - Please use :: to comment any line. To uncomment, please remove :: from start of the line
:: -----------------------------------------------------------------------------------------------

::FD-Regression.txt Scenario 37a: UI Preferences Audit Pass
Call pybot -i TC37a --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 37b: UI Preferences Audit Fail
Call pybot -i TC37b --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

Exit

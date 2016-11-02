cd\
cd TomH\Code\FDLD
:: -----------------------------------------------------------------------------------------------
:: Note - Please use :: to comment any line. To uncomment, please remove :: from start of the line
:: -----------------------------------------------------------------------------------------------

::FD-Regression.txt Scenario 43a: Set CRM Preferences - Set DXA, DXB, DXC to Receive DWR
Call pybot -i TC43a --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt TC38 Routed Leads DDF Audit
Call pybot -i TC38 --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 43b Unset CRM Preferences
Call pybot -i TC43b --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

PAUSE
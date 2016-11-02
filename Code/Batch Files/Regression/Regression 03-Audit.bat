cd\
cd TomH\Code\FDLD
:: -----------------------------------------------------------------------------------------------
:: Note - Please use :: to comment any line. To uncomment, please remove :: from start of the line
:: -----------------------------------------------------------------------------------------------

::See note for Scenario 38

::Set Preferences Full
::FD-Regression.txt Scenario 42a: Set Specific Preferences - Full RF=Y
Call pybot -i TC42a --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 05: DDF Audit
Call pybot -i TC05 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 06: GetLead Audit
Call pybot -i TC06 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt Scenario 24: DDF Versions
Call pybot -i TC24 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::Set Preferences for Routed Leads
::FD-Regression.txt Scenario 38: Routed Leads DDF Audit
Call pybot -i TC38 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

::No point in doing this after an FD LeadUpdate. We would need to map a native TD Lead to DDF.
::FD-Regression.txt Scenario 28: TradeDriver DDF Audit
::Call pybot -i TC28 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-Regression.txt

Exit

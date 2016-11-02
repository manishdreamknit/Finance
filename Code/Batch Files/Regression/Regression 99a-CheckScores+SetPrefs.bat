cd\
cd TomH\Code\FDLD
:: -----------------------------------------------------------------------------------------------
:: Note - Please use :: to comment any line. To uncomment, please remove :: from start of the line
:: -----------------------------------------------------------------------------------------------

::FD-HealthCheckup.txt TC01 Setup Test Using CBWS Request
Call pybot -i TC01 --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_ FD-HealthCheckup.txt

Exit

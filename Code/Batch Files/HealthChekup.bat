cd\
cd Project\FDLD\testcases\FDLD
:: -----------------------------------------------------------------------------------------------
:: Note - Please use :: to comment any line. To uncomment, please remove :: from start of the line
:: -----------------------------------------------------------------------------------------------

:: Below command will execute complete Credit App test suite which contains 7 scenarios.
Call pybot --timestampoutputs --log ExecutionLog\HealthCheckup_Log --output ExecutionLog\HealthCheckup_Output --report ExecutionLog\HealthCheckup_report -i TC03 --variable FOLDER_PATH:Input\\Health_Checkup --escape space:_ --variable REPLACE_PARTNER:y --variable HCScenarioFile:ScenarioHealthCheckup.txt --escape space:_  FD-HealthCheckup.txt

PAUSE
Exit

cd\
cd TomH\Code\FDLD
:: -----------------------------------------------------------------------------------------------
:: Note - Please use :: to comment any line. To uncomment, please remove :: from start of the line
:: -----------------------------------------------------------------------------------------------

::FD-Regression.txt TC02 Check Services
Call pybot -i TC02 --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

PAUSE

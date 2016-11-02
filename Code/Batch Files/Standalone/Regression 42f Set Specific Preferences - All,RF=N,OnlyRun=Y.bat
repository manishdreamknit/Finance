cd\
cd TomH\Code\FDLD
:: -----------------------------------------------------------------------------------------------
:: Note - Please use :: to comment any line. To uncomment, please remove :: from start of the line
:: -----------------------------------------------------------------------------------------------

::Set Preferences All,RF=Y,OnlyRun=Y
::FD-Regression.txt Scenario 42f: Set Specific Preferences - All,RF=N,OnlyRun=Y
Call pybot -i TC42f --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

PAUSE

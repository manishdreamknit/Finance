cd\
cd TomH\Code\FDLD
:: -----------------------------------------------------------------------------------------------
:: Note - Please use :: to comment any line. To uncomment, please remove :: from start of the line
:: -----------------------------------------------------------------------------------------------

::Set Preference All,RF=Y,OnlyRun=Y
::FD-Regression.txt Scenario 42e: Set Specific Preferences - All,RF=Y,OnlyRun=Y
Call pybot -i TC42e --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

Exit

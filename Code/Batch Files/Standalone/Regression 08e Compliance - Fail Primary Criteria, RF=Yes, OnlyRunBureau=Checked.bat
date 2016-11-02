cd\
cd TomH\Code\FDLD
:: -----------------------------------------------------------------------------------------------
:: Note - Please use :: to comment any line. To uncomment, please remove :: from start of the line
:: -----------------------------------------------------------------------------------------------

::Set Preferences All,RF=Y,OnlyRun=Y
::FD-Regression.txt Scenario 42e: Set Specific Preferences - All,RF=Y,OnlyRun=Y
Call pybot -i TC42e --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt TC08e: Compliance - Fail Primary Criteria, RF=Yes, OnlyRunBureau=Checked
Call pybot -i TC08e --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

PAUSE

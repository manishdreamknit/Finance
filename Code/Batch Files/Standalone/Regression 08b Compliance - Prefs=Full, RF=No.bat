cd\
cd TomH\Code\FDLD
:: -----------------------------------------------------------------------------------------------
:: Note - Please use :: to comment any line. To uncomment, please remove :: from start of the line
:: -----------------------------------------------------------------------------------------------

::Set Preferences Full RF=N
::FD-Regression.txt Scenario 42b: Set Specific Preferences - Full RF=N
Call pybot -i TC42b --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt TC08b: Compliance - Prefs=Full, RF=No
Call pybot -i TC08b --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

PAUSE

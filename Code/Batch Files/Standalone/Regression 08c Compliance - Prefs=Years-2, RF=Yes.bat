cd\
cd TomH\Code\FDLD
:: -----------------------------------------------------------------------------------------------
:: Note - Please use :: to comment any line. To uncomment, please remove :: from start of the line
:: -----------------------------------------------------------------------------------------------

::Set Preference YearsAtJob Only - RF=Y
::FD-Regression.txt Scenario 42c: Set Specific Preferences - YearsAtJob Only - RF=Y
Call pybot -i TC42c --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt TC08c: Compliance - Prefs=Years-2, RF=Yes
Call pybot -i TC08c --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

PAUSE

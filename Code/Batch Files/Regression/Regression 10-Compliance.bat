cd\
cd TomH\Code\FDLD
:: -----------------------------------------------------------------------------------------------
:: Note - Please use :: to comment any line. To uncomment, please remove :: from start of the line
:: -----------------------------------------------------------------------------------------------

::Set Preferences Full RF=Y
::FD-Regression.txt Scenario 42a: Set Specific Preferences - Full RF=Y
Call pybot -i TC42a --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt TC08a: Compliance - Prefs=Full, RF=Yes
Call pybot -i TC08a --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

::

::Set Preferences Full RF=N
::FD-Regression.txt Scenario 42b: Set Specific Preferences - Full RF=N
Call pybot -i TC42b --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt TC08b: Compliance - Prefs=Full, RF=No
Call pybot -i TC08b --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

::

::Set Preference YearsAtJob Only - RF=Y
::FD-Regression.txt Scenario 42c: Set Specific Preferences - YearsAtJob Only - RF=Y
Call pybot -i TC42c --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

--FD-Regression.txt TC08c: Compliance - Prefs=Years-2, RF=Yes
Call pybot -i TC08c --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

::

::Set Preference YearsAtJob Only - RF=N
::FD-Regression.txt Scenario 42d: Set Specific Preferences - YearsAtJob Only - RF=N
Call pybot -i TC42d --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

--FD-Regression.txt TC08d: Compliance - Prefs=Years-2, RF=No
Call pybot -i TC08d --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

::

::Set Preferences All,RF=Y,OnlyRun=Y
::FD-Regression.txt Scenario 42e: Set Specific Preferences - All,RF=Y,OnlyRun=Y
Call pybot -i TC42e --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

--FD-Regression.txt TC08e: Compliance - Fail Primary Criteria, RF=Yes, OnlyRunBureau=Checked
Call pybot -i TC08e --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

::

::Set Preferences All,RF=Y,OnlyRun=Y
::FD-Regression.txt Scenario 42f: Set Specific Preferences - All,RF=N,OnlyRun=Y
Call pybot -i TC42f --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

::FD-Regression.txt TC08f: Compliance - Fail Primary Criteria, RF=No, OnlyRunBureau=Checked
Call pybot -i TC08f --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

::

::FD-Regression.txt TC08g: Compliance - CBWS, RF=Yes
Call pybot -i TC08g --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

::

::FD-Regression.txt TC08h: Compliance - CBWS, RF=No
Call pybot -i TC08h --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

::

::Restore Full Preferences
::FD-Regression.txt Scenario 42a: Set Specific Preferences - Full RF=Y
Call pybot -i TC42a --variable HCScenarioFile:ScenarioQA.txt --escape space:_ FD-Regression.txt

Pause

Exit

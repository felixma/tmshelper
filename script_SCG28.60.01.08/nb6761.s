Test I.D.: nb6761
Title: PreSU - Software version check
Owner: C.YU
Originator: maqi
Script Status: I
Requirement(s): Implicit
Feature(s): Implicit
Reference(s):
Functional Area(s): 90
Test Level: G
Original Target Application(s): IMS
Parent Test(s):
Execution Mode: MAN
Estimated Execution Time: 30

Description:
PreSU - Software version check

Issue Notes:
none

Resources/Configuration:
none

Initial Conditions/System Setup:

Test Procedure:
1. Log on SCG as lss.
2. Run below commands:
mi_adm -v -a version
lss_adm -v -a version

Verify:
The load number is R28.60.01.02.

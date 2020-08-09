
# Contribution
Any contribution is welcome, last but not least concerning the wording in the documentation which may be sub-optimal since I am not a native English man.

The Excel Workbook [Msg.xlsm](#Msg.xlsm) is for development and test. The module _mTest_ provides all means for testing of the individual features, specifically the exceptions. A proper regression test combining all individual test is available as well, all accessible through the Worksheet _wsMsgTest_.
The test procedures in the _mTest_ module focus on a well designed  tests and in that are not necessarily usefully usage examples. For those see [README](#README.md).

When providing changes or amendments to the implementation, performing a regression test should be obligatory, if not covered yet by the test, test procedures should be added or modified accordingly.

Please note: The Workbook has my "automated update of common components" implemented which will be ignored when the conditional compile argument _CompMan_ is set to 0.

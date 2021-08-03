set expdate=(Date_(MM_DD_YYYY)_%date:~0,2%-%date:~3,2%-%date:~6,4%)-(Time_%time:~0,2%%time:~3,2%)

expdp bankadmin/9122335311 full=Y directory=TEST_DIR dumpfile=bdbacku.dmp logfile=dblog.log
exit
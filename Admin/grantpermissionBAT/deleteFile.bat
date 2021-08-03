set expdate=(Date_(MM_DD_YYYY)_%date:~0,2%-%date:~3,2%-%date:~6,4%)-(Time_%time:~0,2%%time:~3,2%)

cd \oraclexe\app\oracle\oradata/
del BDBACKU.DMP

cd \oraclexe\app\oracle\oradata/
del dblog.log
exit

copy a:\rpms.exe c:\rpigeons\rpms.exe
c:
echo Backup old files
cd \rpigeons
md v32y
cd v32y
copy ..\*.rpt
copy ..\rp.exe
copy ..\rp.mdb
cd..
echo Upgrade reports
rpms
echo Update Database
dbupg
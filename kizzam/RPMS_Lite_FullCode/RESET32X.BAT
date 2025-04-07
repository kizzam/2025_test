echo off
h:
cd \
cd dev_vb
cd rp
if %1==/U goto upgDisk
if %1==/LHA goto upgCOMPRESS
if %1==/upgDISK goto upgDISK
del all.rp_
del bloodl.rp_
del bloodlin.rp_
del breeder.rp_
del breederl.rp_
del breederA.rp_
del brnotes.rp_
del cage.rp_
del colour.rp_
del con50916.rp_
del conshort.rp_
del distance.rp_
del distancg.rp_
del dist_yds.rp_
del dist_ydg.rp_
del eyesign.rp_
del fstruc1.rp_
del genrace.rp_
del genstock.rp_
del kpi.rp_
del main.rp_
del master.rp_
del pedg.rp_
del pedg4.rp_
del point.rp_
del racelst1.rp_
del races1.rp_
del ratings.rp_
del results1.rp_
del rp.ex_
del rp.md_
del rpextra.md_
del salebp.rp_
del salefp.rp_
del salerp.rp_
del team.rp_
del teamcode.rp_
del teaml.rp_
del teaml2.rp_
del teaml3.rp_
del teaml4.rp_
del teaml5.rp_
del teamp.rp_
del toss1.rp_
d:
cd \vb3
cd setupkit
cd kitfiles
h:
d:compress -r All.rpt
d:compress -r bloodl.rpt
d:compress -r bloodlin.rpt
d:compress -r breeder.rpt
d:compress -r breederA.rpt
d:compress -r breederl.rpt
d:compress -r brnotes.rpt
d:compress -r cage.rpt
d:compress -r colour.rpt
d:compress -r genrace.rpt
d:compress -r con50916.rpt
d:compress -r conshort.rpt
d:compress -r distance.rpt
d:compress -r distancg.rpt
d:compress -r dist_yds.rpt
d:compress -r dist_ydg.rpt
d:compress -r eyesign.rpt
d:compress -r fstruc1.rpt
d:compress -r genrace.rpt
d:compress -r genstock.rpt
d:compress -r kpi.rpt
d:compress -r main.rpt
d:compress -r master.rpt
d:compress -r pedg.rpt
d:compress -r pedg4.rpt
d:compress -r point.rpt
d:compress -r racelst1.rpt
d:compress -r races1.rpt
d:compress -r ratings.rpt
d:compress -r results1.rpt
d:compress -r rp.exe
d:compress -r empty\rpempty.mdb
d:compress -r empty\rpextra.mdb
d:compress -r salerp.rpt
d:compress -r salebp.rpt
d:compress -r salefp.rpt
d:compress -r team.rpt
d:compress -r teamcode.rpt
d:compress -r teaml.rpt
d:compress -r teaml2.rpt
d:compress -r teaml3.rpt
d:compress -r teaml4.rpt
d:compress -r teaml5.rpt
d:compress -r teamp.rpt
d:compress -r toss1.rpt
copy all.rp_ setuprp\disk3d\all.rp_ /y
copy bloodl.rp_ setuprp\disk2d\bloodl.rp_ /y
copy bloodlin.rp_ setuprp\disk2d\bloodlin.rp_ /y
copy breeder.rp_ setuprp\disk2d\breeder.rp_ /y
copy breederA.rp_ setuprp\disk3d\breederA.rp_ /y
copy breederl.rp_ setuprp\disk3d\breederl.rp_ /y
copy brnotes.rp_ setuprp\disk3d\brnotes.rp_ /y
copy cage.rp_ setuprp\disk2d\cage.rp_ /y
copy colour.rp_ setuprp\disk2d\colour.rp_ /y
copy con50916.rp_ setuprp\disk2d\con50916.rp_ /y
copy conshort.rp_ setuprp\disk2d\conshort.rp_ /y
copy distance.rp_ setuprp\disk2d\distance.rp_ /y
copy distancg.rp_ setuprp\disk3d\distancg.rp_ /y
copy dist_yds.rp_ setuprp\disk3d\dist_yds.rp_ /y
copy dist_ydg.rp_ setuprp\disk3d\dist_ydg.rp_ /y
copy eyesign.rp_ setuprp\disk2d\eyesign.rp_ /y
copy fstruc1.rp_ setuprp\disk3d\fstruc1.rp_ /y
copy genstock.rp_ setuprp\disk2d\genstock.rp_ /y
copy genrace.rp_ setuprp\disk2d\genrace.rp_ /y
copy kpi.rp_ setuprp\disk3d\kpi.rp_ /y
copy main.rp_ setuprp\disk2d\main.rp_ /y
copy master.rp_ setuprp\disk2d\master.rp_ /y
copy pedg.rp_ setuprp\disk3d\pedg.rp_ /y
copy pedg4.rp_ setuprp\disk2d\pedg4.rp_ /y
copy point.rp_ setuprp\disk2d\point.rp_ /y
copy races1.rp_ setuprp\disk2d\races1.rp_ /y
copy racelst1.rp_ setuprp\disk2d\racelst1.rp_ /y
copy ratings.rp_ setuprp\disk2d\ratings.rp_ /y
copy results1.rp_ setuprp\disk2d\results1.rp_ /y
copy rp.ex_ setuprp\disk1d\rp.ex_ /y
copy empty\rpempty.md_ setuprp\disk3d\rp.md_ /y
copy empty\rpextra.md_ setuprp\disk3d\rpextra.md_ /y
copy salerp.rp_ setuprp\disk3d\salerp.rp_ /y
copy salebp.rp_ setuprp\disk3d\salebp.rp_ /y
copy salefp.rp_ setuprp\disk3d\salefp.rp_ /y
copy team.rp_ setuprp\disk2d\team.rp_ /y
copy teamcode.rp_ setuprp\disk2d\teamcode.rp_ /y
copy teaml.rp_ setuprp\disk2d\teaml.rp_ /y
copy teamp.rp_ setuprp\disk2d\teamp.rp_ /y
rem copy teams.rp_ setuprp\disk2d\teams.rp_ /y
copy teaml2.rp_ setuprp\disk2d\teaml2.rp_ /y
copy teaml3.rp_ setuprp\disk2d\teaml3.rp_ /y
copy teaml4.rp_ setuprp\disk3d\teaml4.rp_ /y
copy teaml5.rp_ setuprp\disk3d\teaml5.rp_ /y
copy toss1.rp_ setuprp\disk2d\toss1.rp_ /y
copy all.rp_ setupcd\cd\all.rp_ /y
copy bloodl.rp_ setupcd\cd\bloodl.rp_ /y
copy bloodlin.rp_ setupcd\cd\bloodlin.rp_ /y
copy breeder.rp_ setupcd\cd\breeder.rp_ /y
copy breederA.rp_ setupcd\cd\breederA.rp_ /y
copy breederl.rp_ setupcd\cd\breederl.rp_ /y
copy brnotes.rp_ setupcd\cd\brnotes.rp_ /y
copy cage.rp_ setupcd\cd\cage.rp_ /y
copy colour.rp_ setupcd\cd\colour.rp_ /y
copy con50916.rp_ setupcd\cd\con50916.rp_ /y
copy conshort.rp_ setupcd\cd\conshort.rp_ /y
copy distance.rp_ setupcd\cd\distance.rp_ /y
copy distancg.rp_ setupcd\cd\distancg.rp_ /y
copy dist_yds.rp_ setupcd\cd\dist_yds.rp_ /y
copy dist_ydg.rp_ setupcd\cd\dist_ydg.rp_ /y
copy eyesign.rp_ setupcd\cd\eyesign.rp_ /y
copy fstruc1.rp_ setupcd\cd\fstruc1.rp_ /y
copy genstock.rp_ setupcd\cd\genstock.rp_ /y
copy genrace.rp_ setupcd\cd\genrace.rp_ /y
copy kpi.rp_ setupcd\cd\kpi.rp_ /y
copy main.rp_ setupcd\cd\main.rp_ /y
copy master.rp_ setupcd\cd\master.rp_ /y
copy pedg.rp_ setupcd\cd\pedg.rp_ /y
copy pedg4.rp_ setupcd\cd\pedg4.rp_ /y
copy point.rp_ setupcd\cd\point.rp_ /y
copy races1.rp_ setupcd\cd\races1.rp_ /y
copy racelst1.rp_ setupcd\cd\racelst1.rp_ /y
copy ratings.rp_ setupcd\cd\ratings.rp_ /y
copy results1.rp_ setupcd\cd\results1.rp_ /y
copy rp.ex_ setupcd\cd\rp.ex_ /y
copy empty\rpempty.md_ setupcd\cd\rp.md_ /y
copy empty\rpextra.md_ setupcd\cd\rpextra.md_ /y
copy salerp.rp_ setupcd\cd\salerp.rp_ /y
copy salebp.rp_ setupcd\cd\salebp.rp_ /y
copy salefp.rp_ setupcd\cd\salefp.rp_ /y
copy team.rp_ setupcd\cd\team.rp_ /y
copy teamcode.rp_ setupcd\cd\teamcode.rp_ /y
copy teaml.rp_ setupcd\cd\teaml.rp_ /y
copy teamp.rp_ setupcd\cd\teamp.rp_ /y
rem copy teams.rp_ setupcd\cd\teams.rp_ /y
copy teaml2.rp_ setupcd\cd\teaml2.rp_ /y
copy teaml3.rp_ setupcd\cd\teaml3.rp_ /y
copy teaml4.rp_ setupcd\cd\teaml4.rp_ /y
copy teaml5.rp_ setupcd\cd\teaml5.rp_ /y
copy toss1.rp_ setupcd\cd\toss1.rp_ /y
copy rp.exe setupcd\cd\suppdisk /y
copy empty\rpempty.mdb setupcd\cd\suppdisk /y
copy empty\rpextra.mdb setupcd\cd\suppdisk /y
pause
copy rp.exe suppdisk /y
copy empty\rpempty.mdb suppdisk /y
copy empty\rpextra.mdb suppdisk /y
pause
cd setuprp
d:compress -r setup1.exe
copy setup1.ex_ disk1d\setup1.ex_ /y
cd..
cd setupcd
d:compress -r setup1.exe
copy setup1.ex_ cd\setup1.ex_
cd ..
pause
:upgDisk
echo upgDisk
pause
copy all.rpt upgrade1\all.rpt /y
copy bloodl.rpt upgrade1\bloodl.rpt /y
copy bloodlin.rpt upgrade1\bloodlin.rpt /y
copy breeder.rpt upgrade1\breeder.rpt /y
copy breederA.rpt upgrade1\breederA.rpt /y
copy breederl.rpt upgrade1\breederl.rpt /y
copy brnotes.rpt upgrade1\brnotes.rpt /y
copy cage.rpt upgrade1\cage.rpt /y
copy colour.rpt upgrade1\colour.rpt /y
copy con50916.rpt upgrade1\con50916.rpt /y
copy conshort.rpt upgrade1\conshort.rpt /y
copy distance.rpt upgrade1\distance.rpt /y
copy distancg.rpt upgrade1\distancg.rpt /y
copy dist_yds.rpt upgrade1\dist_yds.rpt /y
copy dist_ydg.rpt upgrade1\dist_ydg.rpt /y
copy eyesign.rpt upgrade1\eyesign.rpt /y
copy fstruc1.rpt upgrade1\fstruc1.rpt /y
copy genstock.rpt upgrade1\genstock.rpt /y
copy genrace.rpt upgrade1\genrace.rpt /y
copy kpi.rpt upgrade1\kpi.rpt /y
copy main.rpt upgrade1\main.rpt /y
copy master.rpt upgrade1\master.rpt /y
copy pedg.rpt upgrade1\pedg.rpt /y
copy pedg4.rpt upgrade1\pedg4.rpt /y
copy point.rpt upgrade1\point.rpt /y
copy races1.rpt upgrade1\races1.rpt /y
copy racelst1.rpt upgrade1\racelst1.rpt /y
copy ratings.rpt upgrade1\ratings.rpt /y
copy results1.rpt upgrade1\results1.rpt /y
copy rp.exe upgrade1\rp.exe /y
rem copy empty\rpempty.mdb upgrade1\rp.mdb /y
rem copy empty\rpextra.mdb upgrade1\rpextra.mdb /y
copy salerp.rpt upgrade1\salerp.rpt /y
copy salebp.rpt upgrade1\salebp.rpt /y
copy salefp.rpt upgrade1\salefp.rpt /y
copy team.rpt upgrade1\team.rpt /y
copy teamcode.rpt upgrade1\teamcode.rpt /y
copy teaml.rpt upgrade1\teaml.rpt /y
copy teamp.rpt upgrade1\teamp.rpt /y
rem copy teams.rpt upgrade1\teams.rpt /y
copy teaml2.rpt upgrade1\teaml2.rpt /y
copy teaml3.rpt upgrade1\teaml3.rpt /y
copy teaml4.rpt upgrade1\teaml4.rpt /y
copy teaml5.rpt upgrade1\teaml5.rpt /y
copy toss1.rpt upgrade1\toss1.rpt /y
:upgCOMPRESS
echo on
cd upgrade1
d:
cd\
cd dosprogs\compress\lha
h:
cd\dev_vb\rp\upgrade1
del rpms.exe
del rpms.lzh
del rpms.32y
d:lha a rpms ..\*.rpt ..\rp.exe ..\empty\rpextra.mdb /m+
pause
d:lha s rpms /m-
ren rpms.exe rpms.32y
:upgDisk
cd\dev_vb\rp\upgrade1
copy rpms.32y disk1\rpms.32y /y
copy upgrade.exe disk1\upgrade.exe /y
del *.rp_
del rp.ex_



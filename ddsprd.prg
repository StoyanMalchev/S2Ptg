*!*	От файла ПРОДАЖБИ на ДНЕВНИЦИ създава данни за импорт в ПИТАГОР

SET EXCLUSIVE off
set dele on
set cent on
set date dmy
SET SAFETY off
SET TALK off
CLOSE TABLES all
use prmw
SELECT * from ktrg WHERE .f. into CURSOR ktrg__ readw 

SELECT 0

set date dmy

SELECT * from prodagbi WHERE .f. into CURSOR prd_ readw 
APPEND FROM prodagbi.txt sdf

SELECT docno as fan,CTOD(ddate)as dt, ;
    SPACE(10)as ktrg, knum as id_dds, SPACE(13)as efn, kname, ;  
    STR(1,10)as artkl, 1 as q, 20 as pdds, ;
    ss10 as StFar,ss10 as cn, ss10+ss20 as St2, "703" as ptdr ;
  FROM prd_ into CURSOR prd_ readw

REPLACE pdds WITH 0 FOR StFar=St2
REPLACE pdds WITH 9 FOR STR(StFar*1.09,15,2) = STR(St2,15,2)

REPLACE ALL efn WITH IIF(id_dds="BG", SUBSTR(id_dds,3,13), id_dds)

SELECT * from prd_ where NOT (bstat(efn) OR egn(efn)) INTO CURSOR xxx
IF RECCOUNT("xxx") > 0
   BROWSE TITLE "Записи с некоректен булстат:"
ENDIF 
COPY TO Problemi.log sdf
use

SELECT prd_
GO top
  
SCAN
  SELECT ktrg
  LOCATE FOR efn=prd_.efn
  IF EOF("ktrg")
     INSERT INTO ktrg__ (efn,name) VALUES (prd_.efn, prd_.kname)
     DO case
     CASE egn(prd_.efn)
        REPLACE ktrg WITH "p"+SUBSTR(prd_.efn,1,9)in ktrg__
     CASE bstat(prd_.efn)
        REPLACE ktrg WITH SUBSTR(prd_.efn,1,8)in ktrg__
     OTHERWISE 
        LOCATE FOR ktrg="_000099999"
        IF EOF("ktrg")
           INSERT INTO ktrg__ (ktrg,name) VALUES ("_000099999", "???")
        ENDIF         
     ENDCASE 
     lcKtrg=ktrg__.ktrg
  ELSE
     lcKtrg=ktrg.ktrg  
  ENDIF 
  REPLACE ktrg WITH lcKtrg IN prd_
ENDSCAN

SELECT ktrg__
COPY TO _ktrg.xxx  

SELECT * from prd_ where YEAR(dt)>2011 INTO CURSOR fa__

SELECT * from fa WHERE .f. into CURSOR _fa readw
APPEND FROM DBF("fa__")
COPY TO _fa.xxx
GO top
lnGG = YEAR(dt)

SELECT * from artkl WHERE .f. into CURSOR _artkl
COPY TO _artkl.xxx 

SELECT * from fa2 WHERE .f. into CURSOR _fa2 readw
APPEND FROM DBF("fa__")
REPLACE ALL gg WITH lnGG IN _fa2
copy to _fa2.xxx
SUM st2 TO lnTtl

SELECT prmw.efn as efn, MIN(Dt)as dt1, MAX(Dt)as dt2, "01" as User,;
    MIN(fan)as fan1, MAX(fan)as fan2, ;
    sum(st2)as ttl, 0 as ttlDkm, RECC("fa__")as brF ;
  FROM fa__ INTO curs _exprt

COPY to _exprt.xxx
    
SELECT fas,fan, "pp"as td, fan as NDok, "1"as nr, dt,;
    "501  "as Debit, "411"+SPACE(7)+ktrg+STR(fan,10)as kredit, ttl as StDkm ;
  FROM _fa WHERE skaw="411" GROUP BY fas,fan      INTO CURSOR dkm_
  
COPY TO _dkm.xxx  
lcTmpDir = ADDBS(GETDIR("c:\temp","Изберете папка"))
oZip = CreateObject("XStandard.Zip")
ozip.Pack("_fa.xxx",  lcTmpDir+"_exprt.zip", .f., "", 9)
ozip.Pack("_fa.fpt",  lcTmpDir+"_exprt.zip", .f., "", 9)
ozip.Pack("_fa2.xxx", lcTmpDir+"_exprt.zip", .f., "", 9)
ozip.Pack("_fa2.fpt", lcTmpDir+"_exprt.zip", .f., "", 9)
ozip.Pack("_ktrg.xxx", lcTmpDir+"_exprt.zip", .f., "", 9)
ozip.Pack("_artkl.xxx",lcTmpDir+"_exprt.zip", .f., "", 9)
ozip.Pack("_dkm.xxx",  lcTmpDir+"_exprt.zip", .f., "", 9)
ozip.Pack("_dkm.fpt",  lcTmpDir+"_exprt.zip", .f., "", 9)
ozip.Pack("_exprt.xxx",lcTmpDir+"_exprt.zip", .f., "", 9) 
    

MESSAGEBOX(TRAN(RECC("fa__") )+" документа на стойност"+;
    str(lnTtl,13,2)+" лв. "+CHR(13),48, "Трансфер на продажби")

CLOSE TABLES ALL

**********
func bstat(cEfn)

priv lcEFN, lcVal, j
lcEFN = substr(alltrim(cEFN),1,9)
lcVal = 0
for j=1 to 8 
  lcVal = lcVal +  val(substr(lcEFN,j,1))*j 
  next
lcVal = mod(lcVal, 11)
if lcVal=10  
   lcVal=0
   for j=1 to 8 
     lcVal = lcVal + val(substr(lcEFN,j,1))*(j+2) 
     next
   lcVal = mod(lcVal,11)
   if lcVal=10 
      lcVal=0
   endif
endif
retu substr(lcEFN,9,1)=str(lcVal,1)


********
func egn(s)
IF len(alltrim(s)) <> 10
	RETURN .f.
EndIf

PRIVATE I, v, teq
DIMENSION teq(9)
ALINES(teq, "2,4,8,5,10,9,7,3,6", .F., ",")

IF not (between(val(substr(s,3,2)),1,12) or ;
        between(val(substr(s,3,2)),41,52) )
	RETURN .f.
ENDIF
IF ((substr(s,3,2) $ "01;03;05;07;08;10;12") or ;
    (substr(s,3,2) $ "41;43;45;47;48;50;52")) 	and ;
    ( val(substr(s,5,2))>31 or val(substr(s,5,2))<1 )
	RETURN .f.
ENDIF
IF ((substr(s,3,2) $ "04;06;09;11") or (substr(s,3,2) $ "44;46;49;51")) ;
		and ( !between(val(substr(s,5,2)),1,30) )
	RETURN .f.
ENDIF
IF ((substr(s,3,2) = "02") or (substr(s,3,2) = "42")) and mod(val(substr(s,1,2)),4)<>0 and;
		( !between(val(substr(s,5,2)),1,28) )
	RETURN .f.
ENDIF
IF (substr(s,3,2)="02" or substr(s,3,2)="42") and mod(val(substr(s,1,2)),4)=0 and;
		( !between(val(substr(s,5,2)),1,29) )
	RETURN .f.
ENDIF

v=0
FOR I=1 to 9
	v = v + val(substr(s,I,1))*Val(teq(I))
NEXT
v = mod(v,11)
IF v=10
	v=0
ENDIF
IF substr(s,10,1) = str(v,1)
	RETURN .T.
ELSE
	RETURN .F.
ENDIF

**********
FUNC CpCnv(pStr)
private v,vv,i
v=""
for i=1 to len(pStr)
   vv = substr(pStr, i, 1)
   vvv = asc(vv)
   do case
*   case inlist(vvv,201,205,192,193,194,196)
*      vv= "="
*   case inlist(vvv,192,195,197,198,199,200,211,218)
*      vv="|"
   case vvv>127 and vvv<192
      vv=chr(vvv+64)
   endcase
   v=v+vv
endfor
return v   

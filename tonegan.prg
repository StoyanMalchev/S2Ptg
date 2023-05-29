** Артикулите са с идентификаторите от ПИТАГОР  !

PUBLIC gcTmpDir
gcTmpDir = addbs(SYS(2023)) + "_exprt\"
IF NOT DIRECTORY(gcTmpDir)
   MD (gcTmpDir)
ENDIF 
ERASE ( gcTmpDir+"*.*" )

PUBLIC _path1, _path, _eik, _usr
_Path1=SYS(5)+curdir()
SET PATH TO .

PUBLIC gcTmpDir
lcFN = GETFILE("xls", "Файл:","Избери",0, "Документи за продажби:")

IMPORT FROM (lcFN) XL5 
*IMPORT FROM lznc_agro.xls xl5
SELECT ROUND(VAL(p),0)as fan,CTOD(f)as dt, ;
   SUBSTR(h,1,15)as id_dds,SUBSTR(k,1,13)as efn, m as kname, ;
    STR(VAL(b),10)as art, d as AName, ;
    ROUND(VAL(ab),3)as q, ;
    ROUND(VAL(t),2)as StFar, ;
    ROUND(VAL(z),2)as St2, ad as cPmnt ;   
  FROM (ALIAS())  ;
  having fan > 0 INTO CURSOR fa_ readw

*!*	SELECT dist art, aname FROM fa_
*!*	retu
lcID = "BG;BE;GB;DE;EL;DK;EE;IE;ES;IT;CY;LV;LT;LU;MT;PL;PT;RO;SK;SI;HU;FI;FR;NL;CZ;SE;HR;ATU;"
REPLACE id_dds WITH "" ;
  FOR NOT SUBSTR(id_dds,1,2)+";" $ lcID

CALCULATE MIN(dt), MAX(dt) TO gDt1, gDt2
  
REPLACE efn WITH PADL(ALLTRIM(efn),9,"0") FOR LEN(TRIM(efn))<9  
*SELECT dist efn,kname FROM fa_ where LEN(TRIM(efn))>9 and bstat(efn)
repl efn with SUBSTR(efn,1,9)for LEN(TRIM(efn))>9 and bstat(efn)

SELECT fa_.*, 00000 as NDok, {} as dtz,;
  " "as fas,SPACE(10)as ktrg ;
  FROM fa_ into CURSOR fa_ readw

REPLACE NDok WITH 1, dtz WITH {1.1.2011} for StFar < 0
  
SELECT efn,kname,fan,dt FROM fa_   ;
  where NOT (bstat(efn)or egn(efn)) ;
  GROUP BY efn,kname ;
  INTO curs tst_
IF RECCOUNT("tst_") > 0
   BROWSE TITLE "Некоректни EIK:"
ENDIF    

SELECT SPACE(10)as ktrg,KName as name, efn, id_dds ;
      FROM fa_  ;
  WHERE NOT EMPTY(efn) ;
  GROUP BY efn ;
  INTO CURSOR ktrg_ readw

DO _ktrg 

* SELECT * from fa_ group BY ktrg,fan

*!*	SELECT NVL(k_.ktrg,SPACE(10))as ktrg, fa_.* from fa_ ;
*!*	   LEFT JOIN ktrg k_ ON k_.efn=fa_.efn

********
IF NOT FILE("art_.dat")
   SELECT 0
   CREATE TABLE art_.dat FREE (art c(20), name c(40), mka c(5), artkl c(10))
ENDIF 

SELECT art as artkl, aname as name,20 as pdds,;
  "304"+STR(1,17)+art as PtdSP, "702"as ptd ;
  from fa_ group BY art into CURSOR art_ readw

* BROWSE
*COPY TO _artkl.dat  
COPY TO (gcTmpDir+"_artkl.dat")


*!*	SELECT artkl_.*, NVL(artkl,SPACE(10))as artkl, NVL(mka, SPACE(5))as mka ;
*!*	  FROM artkl_ left JOIN artkl ON artkl.artkl=artkl_.art ;
*!*	  ORDER BY artkl INTO CURSOR artkl_ readw

*!*	SELECT artkl_.* from artkl_ ;
*!*	  WHERE NOT exists (sele art from art_.dat a_ where a_.art=artkl_.art) ;
*!*	     INTO CURSOR a__
*!*	SELECT art_
*!*	APPEND FROM DBF("a__")     
*!*	USE IN artkl_  

lcArtkl=SPACE(10)
SELECT YEAR(dt)as gg," "as fas,fan, 001 as far,;
    NVL(art_.artkl, lcArtkl)as artkl, q, ROUND(StFar/q,4)as cn,StFar,st2,"11"as ws, ;
    SUBSTR(id_DDS,1,2)as id_ ;
  FROM fa_  ;
  LEFT JOIN art_ on art_.artkl=fa_.art  ;
  WHERE YEAR(dt)=YEAR(gDt1) ;
  ORDER BY fas,fan    INTO CURSOR fa2_ readw

SELECT fa2_.*, NVL(a_.pdds,20)as pdds,;
    NVL(a_.PtdSP, "304"+STR(1,17)+fa2_.artkl) as PtdSP, ;
    NVL(a_.Ptd, "702")as PtdR ;
  FROM fa2_ LEFT JOIN artkl a_ on a_.artkl=fa2_.artkl ;
  ORDER BY fas,fan ;
  INTO CURSOR fa2_ readw 

DO ws_art IN s2ptg

SUM st2 TO lnTtl

GO top
DO WHILE NOT EOF()
   lnFan=fan
   lnRow=1
   DO WHILE fan=lnFan AND NOT EOF()
      REPLACE far WITH lnRow
      lnRow = lnRow + 1 
      skip
   ENDDO    
ENDDO 

* brow
*COPY TO _fa2.dat
COPY TO (gcTmpDir+"_fa2.dat")

SELECT fas,fan,dt, ktrg, sum(st2)as ttl, cPmnt FROM fa_ ;
  group BY fan INTO CURSOR fa1_

SELECT * from fa WHERE .f. into CURSOR _fa readw
APPEND FROM DBF("fa1_")
*COPY TO _fa.dat
COPY TO (gcTmpDir+"_fa.dat")

***********
lcPtd501 = "501"+SPACE(37)
SELECT fas,fan, "pp"as td, 000000000 as NDok,000 as nr, dt, ;
    "411"+SPACE(7)+ktrg+STR(fan,10)as kredit, lcPtd501 as debit,;
    ttl as StDkm ;
  FROM fa1_ where EMPTY(fas) AND "в брой" $ LOWER(cPmnt) ;
  GROUP BY fas,fan     INTO CURSOR dkmp_ readw

REPLACE debit WITH kredit FOR StDkm < 0
REPLACE kredit WITH lcPtd501, StDkm WITH -StDkm FOR StDkm < 0

REPLACE ALL NDok WITH 100000+(dt-DATE(YEAR(Dt),1,1)+1)*100
  
SELECT * from dkmp_ order BY NDok INTO CURSOR dkmp_ readw 
DO WHILE NOT EOF()
   lnRow=1
   lnDkm=NDok
   DO WHILE NDok=lnDkm AND NOT EOF()
      REPLACE nr WITH lnRow
      lnRow = lnRow + 1 
      SKIP
   ENDDO
ENDDO 
* brow
COPY TO (gcTmpDir+"_dkm.dat")

USE prmw IN 0
SELECT prmw.efn as efn, gDt1 as dt1, gDt2 as dt2, "01"as User,;
    MIN(fan)as fan1, MAX(fan)as fan2, ;
    m.lnTtl as ttl, 0 as ttlDkm, RECC("fa_")as brF, 0 as brDkm ;
  FROM fa_ INTO curs _exprt

COPY TO (gcTmpDir+"_exprt.dat")

lcPath=GETDIR(SYS(5)+curdir(), "Изберете папка:", "Експорт на данни за продажбите:")
ERASE ( lcPath+"_exprt.zip" )
oZip = CreateObject("XStandard.Zip")
ozip.Pack(gcTmpDir+"_fa.dat",  lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_fa.fpt",  lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_fa2.dat", lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_ktrg.dat", lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_artkl.dat", lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_dkm.dat",  lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_exprt.dat",lcPath+"_exprt.zip", .f., "", 9)

  

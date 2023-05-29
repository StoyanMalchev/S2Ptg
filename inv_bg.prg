SET DATE DMY 

SELECT 0
IF NOT USED("artkl")
   USE artkl
ELSE
   SELECT artkl
ENDIF       
LOCATE FOR VAL(artkl)=11
IF EOF("artkl")
   MESSAGEBOX("Дефинирайте артикул 11 ",16, "Прекъсване:")
   retu
ENDIF 
lcPtd70 = artkl.ptd

lcFN = GETFILE("txt", "Файл:","Избери",0, "Документи във формат INV.BG:")
CREATE CURSOR s_ (f1 c(10), Dt d, fan n(10), f2 c(10),st2 n(15,2), f9 c(10),;
    KName c(50), f3 c(1),f4 c(10), f5 c(10),;
    id_dds c(15), efn c(13), f6 c(10), f7 c(10), f8 c(10),;
    nVAT n(15,2))

*APPEND FROM (lcFN) DELIMITED WITH CHARACTER  "|"
APPEND FROM (lcFN) DELIMITED WITH _ with CHARACTER  "|"

SELECT " "as fas,fan,dt, efn, SPACE(10)as ktrg, KName, ID_DDS, ;
    "411" as skaw,  st2 as ttl, st2 - nVAT as StFar, St2 FROM s_ ;
  into CURSOR fa_ readw

SELECT SPACE(10)as ktrg,KName as name, efn, id_dds  FROM s_  ;
  WHERE NOT EMPTY(efn) ;
  GROUP BY efn ;
  INTO CURSOR ktrg_ readw

DO _ktrg  IN s2ptg

SELECT fa_
SCAN
  lnSZ=LEN(ALLTRIM(efn))
  IF BETWEEN(lnSZ, 5, 8)
     REPLACE efn WITH REPLICATE("0", 9-lnSZ)+ALLTRIM(efn)
  ENDIF 
ENDSCAN
*BROWSE FOR NOT EMPTY(efn)

REPLACE fas WITH "к", fan WITH dt - DATE(YEAR(dt),1,1) + 1 FOR EMPTY(efn)
CALCULATE MIN(dt), MAX(dt) TO gDt1, gDt2

SELECT SPACE(10)as ktrg, KName as name, efn, id_dds  FROM fa_  ;
  WHERE NOT EMPTY(efn) ;
  GROUP BY efn ;
  INTO CURSOR ktrg_ readw

DO _ktrg  

SELECT fa_
GO top
SCAN
   IF fas=" "and EMPTY(ktrg)
      SELECT ktrg
      LOCATE FOR efn=SUBSTR(fa_.efn,1,13)
      IF eof() OR EMPTY(efn)
         REPLACE ktrg WITH "_000000999" IN fa_
      ELSE
         REPLACE ktrg WITH ktrg.ktrg IN fa_      
      ENDIF    
   ENDIF 
ENDSCAN 

SELECT fa_
COPY TO (gcTmpDir+"_fa.dat")

SELECT YEAR(dt)as gg, " "as fas, fan, 1 as far,;
    STR(11,10)as artkl, 1 as q, StFar as cn, StFar, st2, ;
    20 as pdds, m.lcPtd70 as ptdr, SPACE(40)as PtdSP, "11"as ws, ;
    SUBSTR(id_DDS,1,2)as id_ ;
  FROM fa_  ;
  ORDER BY fas,fan    INTO CURSOR fa2_ readw

REPLACE pdds WITH 9, ws WITH "13" FOR ROUND(StFar * 1.09, 2)= st2
REPLACE pdds WITH 0, ws WITH "19" FOR ROUND(StFar, 2)= st2
REPLACE ws WITH "15" FOR pdds=0 AND id_+";" $ "ATU;BE;GB;DE;EL;DK;EE;IE;ES;IT;CY;LV;LT;LU;MT;PL;PT;RO;SK;SI;HU;FI;FR;NL;CZ;SE;HR;"

*UPDATE fa2_ set pdds=9, ws="13" WHERE ROUND(StFar * 1.09, 2)= st2
SELECT * from fa2_  ;
  ORDER BY fas,fan ;
  INTO CURSOR fa2_ readw 

CALCULATE sum(1), sum(st2) TO lnTtlDkm, lnTtl  
  
*!*	GO top
*!*	DO WHILE NOT EOF()
*!*	   lnFan=fan
*!*	   lnFar=1
*!*	   DO WHILE fan=lnFan AND NOT EOF()
*!*	      REPLACE far WITH lnFar
*!*	      lnFar = lnFar + 1 
*!*	      SKIP
*!*	   ENDDO 
*!*	ENDDO    

COPY TO (gcTmpDir+"_fa2.dat")

***********
SELECT fas,fan, "pp"as td, 000000000 as NDok,000 as nr, dt, ;
    "411"+SPACE(7)+ktrg+STR(fan,10)as kredit, "501"as debit,;
    nPmnt as StDkm ;
  FROM fa_ where EMPTY(fas) ;
  GROUP BY fas,fan     INTO CURSOR dkmp_ readw

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

COPY TO (gcTmpDir+"_dkm.dat")

SELECT prmw.efn as efn, gDt1 as dt1, gDt2 as dt2, "01"as User,;
    MIN(fan)as fan1, MAX(fan)as fan2, ;
    m.lnTtl as ttl, lnTtlDkm as ttlDkm, RECC("fa_")as brF, 0 as brDkm ;
  FROM fa_ INTO curs _exprt

COPY TO (gcTmpDir+"_exprt.dat")

*!*	** тест
*!*	SELECT fan, sum(st2)as st2 FROM fa2 ;
*!*	  GROUP BY fan INTO CURSOR _1

*!*	SELECT fan, sum(st2)as st2 FROM fa2_ ;
*!*	  GROUP BY fan INTO CURSOR _2

*!*	SELECT _1.*, _2.st2 as st2_ from _1 ;
*!*	  FULL JOIN _2 ON _1.fan=_2.fan ;
*!*	  having _1.st2 # _2.st2
*!*	**

*lnCount = lnCount - val( loRoot.getElementsByTagName("Records").item[0].nodeTypedValue )
* WAIT TRANSFORM(lnCount)+" записа" WINDOW AT 0,0

IF NOT tstLcnz(_Usr, _eik, gDt2) 
   retu
ENDIF 

lcPath=GETDIR(SYS(5)+curdir(), "Изберете папка:", "Експорт на данни за продажбите:")
ERASE ( lcPath+"_exprt.zip" )
oZip = CreateObject("XStandard.Zip")
ozip.Pack(gcTmpDir+"_fa.dat",  lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_fa2.dat", lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_ktrg.dat", lcPath+"_exprt.zip", .f., "", 9)
*ozip.Pack(gcTmpDir+"_artkl.dat", lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_dkm.dat",  lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_exprt.dat",lcPath+"_exprt.zip", .f., "", 9)

MESSAGEBOX(TRANSFORM(RECC("fa_"))+" записа"+CHR(13)+;
          TRANSFORM(lnTtl)+" лв. с ДДС",64, "Данни за импортиране:")

lnCount = RECC("fa2_")+RECCOUNT("dkmp_")          
lcXml=chkUsg(_Usr, _eik, gDt1, lnCount) 
          
CLEAR events          
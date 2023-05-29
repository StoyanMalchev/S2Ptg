SET DATE DMY 
*CLOSE DATABASES all

SELECT 0
IF NOT USED("artkl")
   USE artkl
ELSE
   SELECT artkl
ENDIF       
LOCATE FOR EMPTY(PtdSP) AND ptd="702"
IF EOF("artkl")
   MESSAGEBOX("Дефинирайте артикул 11 ",16, "Прекъсване:")
   retu
ENDIF 
lcArtkl = artkl.artkl
lcPtd70 = artkl.ptd
WAIT "Артикул:"+artkl.artkl+" "+artkl.name WINDOW AT 0,0 nowa


*!*	CREATE CURSOR s_ (Dt d, f1 c(1), fan n(10), f2 c(1),st2 n(15,2), f9 c(1),;
*!*	    KName c(50), f3 c(1),f4 c(1), f5 c(1),;
*!*	    id_dds c(15), efn c(13), f6 c(1), f7 c(1), f8 c(1),;
*!*	    nVAT n(15,2))
*!*	    
lcFN = GETFILE("xls", "Файл:","Избери",0, "Документи за продажби:")
IMPORT FROM (lcFN) XL5 
SELECT TTOD(CTOt(a))as dt, INT(VAL(c))as fan, D as KName, LOWER(e) as cTtl, ;
    ROUND(VAL(f),2)as StFar, ROUND(VAL(f)+VAL(g),2)as st2, VAL(g)as nVat, LOWER(h)as cPmnt, ;
    i as efn, j as id_dds ;
  FROM (ALIAS()) INTO CURSOR s_ readw

REPLACE all cTtl WITH STRTRAN(cTtl,"лв.", "") 

SELECT " "as fas,fan,dt, 00000 as NDok, {} as dtz, ;
    PADR(allTRIM(efn),13," ")as efn, SPACE(10)as ktrg, KName, ;
    PADR(ALLTRIM(id_dds),15," ")as ID_DDS, ;
    IIF("в брой" $ cPmnt, VAL(cTtl), 000000.00)as nPmnt, "411" as skaw, ;
    VAL(cTtl)as ttl,  ROUND(VAL(cTtl)-nVat,2)as StFar, VAL(cTtl)as St2 FROM s_ ;
  WHERE YEAR(dt) > 2020 AND fan > 0 ;
  into CURSOR fa_ readw

REPLACE NDok WITH 1, dtz WITH {1.1.2011} for StFar < 0

REPLACE efn WITH "0"+efn FOR LEN(TRIM(efn))=8 AND TRIM(efn) $ id_dds
REPLACE efn WITH "00"+efn FOR LEN(TRIM(efn))=7 AND TRIM(efn) $ id_dds
  
REPLACE id_dds WITH "BG"+id_dds FOR id_dds=efn AND LEN(TRIM(id_dds))=9
  
SELECT SPACE(10)as ktrg,KName as name, efn, id_dds ;
      FROM fa_  ;
  WHERE NOT EMPTY(efn) ;
  GROUP BY efn ;
  INTO CURSOR ktrg_ readw

DO _ktrg  IN s2ptg

SELECT fa_
*!*	SCAN
*!*	  lnSZ=LEN(ALLTRIM(efn))
*!*	  IF BETWEEN(lnSZ, 5, 8)
*!*	     REPLACE efn WITH REPLICATE("0", 9-lnSZ)+ALLTRIM(efn)
*!*	  ENDIF 
*!*	ENDSCAN
*BROWSE FOR NOT EMPTY(efn)

*!*	REPLACE fas WITH "к", fan WITH dt - DATE(YEAR(dt),1,1) + 1 FOR EMPTY(efn)

CALCULATE MIN(dt), MAX(dt) TO gDt1, gDt2

*!*	SELECT SPACE(10)as ktrg, KName as name, efn, id_dds  FROM fa_  ;
*!*	  WHERE NOT EMPTY(efn) ;
*!*	  GROUP BY efn ;
*!*	  INTO CURSOR ktrg_ readw

*!*	DO _ktrg  

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
    lcArtkl as artkl, 1 as q, StFar as cn, StFar, st2, ;
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
lcPtd501 = "501"+SPACE(37)
SELECT fas,fan, "pp"as td, 000000000 as NDok,000 as nr, dt, ;
    "411"+SPACE(7)+ktrg+STR(fan,10)as kredit, lcPtd501 as debit,;
    nPmnt as StDkm ;
  FROM fa_ where EMPTY(fas) AND nPmnt # 0 ;
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

COPY TO (gcTmpDir+"_dkm.dat")

SELECT prmw.efn as efn, gDt1 as dt1, gDt2 as dt2, "01"as User,;
    MIN(fan)as fan1, MAX(fan)as fan2, ;
    m.lnTtl as ttl, lnTtlDkm as ttlDkm, RECC("fa_")as brF, 0 as brDkm ;
  FROM fa_ INTO curs _exprt

COPY TO (gcTmpDir+"_exprt.dat")

IF NOT tstLcnz(_Usr, _eik, gDt2) 
   retu
ENDIF 


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

lcPath=GETDIR(SYS(5)+curdir(), "Изберете папка:", "Експорт на данни за продажбите:")
ERASE ( lcPath+"_exprt.zip" )
oZip = CreateObject("XStandard.Zip")
ozip.Pack(gcTmpDir+"_fa.dat",  lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_fa2.dat", lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_ktrg.dat", lcPath+"_exprt.zip", .f., "", 9)
*ozip.Pack(gcTmpDir+"_artkl.dat", lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_dkm.dat",  lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_exprt.dat",lcPath+"_exprt.zip", .f., "", 9)

MESSAGEBOX(TRAN(RECC("fa_"))+" записа"+CHR(13)+;
       "Период от:"+TRAN(gDt1)+" до:"+TRAN(gDt2)+CHR(13)+;
           TRAN(lnTtl)+" лв.вкл. ДДС"+CHR(13)+;
           TRAN(RECC("dkmp_"))+" документа платени в брой" ;
           ,64, "Данни за импортиране:")

lnCount = RECC("fa2_")+RECCOUNT("dkmp_")          
lcXml=chkUsg(_Usr, _eik, gDt1, lnCount) 
                    
CLEAR events          
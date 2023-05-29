*!*	Приложенията СКЛАД и СЧЕТОВОДСТВО използват обща картотека АРТИКУЛИ.
*!*	Ако има новите артикули импорта се прекъсва
*!*	Трябва да се въведат преди импорта на данните


CLOSE DATABASES all
*SET DEFAULT TO ("d:\temp\rines_gensoft2\")

SELECT 0
*lcFN = GETFILE("xls", "Файл:","Избери",0, "Документи за продажби:")
lcFN = LOCFILE("InvoicedSales-*.xls","xls","Документи - файл 1:")
IMPORT FROM (lcFN) XL5 
SELECT ROUND(VAL(J),0)as fan, CTOD(k)as dt,  G as KName,  ;
   ROUND(VAL(i),2)as ttl, e as efn, f as id_dds ;
  FROM (ALIAS()) INTO CURSOR s_
SELECT * from s_ where fan>0 AND NOT EMPTY(dt) ;
  INTO CURSOR s_

*USE IN (ALIAS())

SELECT 0
lcFN = LOCFILE("ItemsByInvoices-*.xls","xls","Документи - файл 2:")
IMPORT FROM (lcFN) XL5 

SELECT ROUND(VAL(p),0)as fan, CTOD(q)as dt, ;
    STR(VAL(g),10)as art, h as nameArt, SUBSTR(i,1,5)as mka, ROUND(VAL(j),2)as q, ;
    ROUND(VAL(n),2) as st2   ;
  FROM ( ALIAS()) INTO CURSOR s2_

SELECT * from s2_ where fan>0 AND NOT EMPTY(dt) ;
  INTO CURSOR s2_


SELECT s_.*, art, nameART,q,mka, st2 FROM s_ ;
  LEFT JOIN s2_ on s2_.fan=s_.fan ;
  INTO CURSOR s_


SELECT  art, nameArt from s_ ;
  GROUP BY art   into CURSOR artkl_

SELECT artkl_.* from artkl_ ;
  WHERE NOT exists ;
    (sele artkl from artkl a_ where a_.artkl=artkl_.art AND NOT EMPTY(ptd)) ;
  INTO CURSOR _a

IF RECCOUNT("_a") > 0
   COPY TO "Недефинирани артикули"  DELIMITED WITH tab
   BROWSE TITLE "Недефинирани артикули:"

	SELECT art as artkl,nameArt as name from artkl_ ;
	  WHERE NOT exists ;
	    (sele artkl from artkl a_ where a_.artkl=artkl_.art )  ;
	  INTO CURSOR __a
	  
  SELECT artkl
  APPEND FROM DBF("__a")
  RETURN  
ENDIF    

SELECT s_.*, artkl.pdds, artkl.ptd as PtdR, artkl.PtdSP, ;
    ROUND(st2/(1+pdds/100),2)as StFar ;
  FROM s_ inner JOIN artkl ON artkl.artkl=art.artkl ;
  INTO CURSOR s_
  
SELECT " "as fas,fan,dt, 00000 as NDok, {} as dtz,;
    PADR(allTRIM(efn),13," ")as efn, SPACE(10)as ktrg, KName, ;
    PADR(ALLTRIM(id_dds),15," ")as ID_DDS, "411" as skaw, ;
    sum(st2)as ttl, 000000000.00 as nPmnt FROM s_ ;
  WHERE YEAR(dt) > 2020 AND fan > 0 ;
  GROUP BY fan   into CURSOR fa_ readw

REPLACE NDok WITH 1, dtz WITH {1.1.2011} for ttl < 0

REPLACE efn WITH "0"+efn FOR LEN(TRIM(efn))=8 AND TRIM(efn) $ id_dds
REPLACE efn WITH "00"+efn FOR LEN(TRIM(efn))=7 AND TRIM(efn) $ id_dds
*REPLACE nPmnt WITH ttl FOR "в брой" $ cPmnt AND NOT "с карта" $ cPmnt
  
REPLACE id_dds WITH "BG"+id_dds FOR id_dds=efn AND LEN(TRIM(id_dds))=9
  
SELECT SPACE(10)as ktrg,KName as name, efn, id_dds ;
      FROM fa_  ;
  WHERE NOT EMPTY(efn) ;
  GROUP BY efn ;
  INTO CURSOR ktrg_ readw

DO _ktrg  IN s2ptg

SELECT fa_

CALCULATE MIN(dt), MAX(dt) TO gDt1, gDt2

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

SELECT YEAR(dt)as gg," "as fas,fan, 001 as far,;
    art as artkl, q, IIF(q=0, 0, StFar/q)as cn, StFar, st2, "11"as ws, ;
    pdds, PtdR, PtdSP,    SUBSTR(id_DDS,1,2)as id_ ;
  FROM s_  ;
  WHERE YEAR(dt)=YEAR(gDt1) ;
  ORDER BY fas,fan    INTO CURSOR fa2_ readw


REPLACE pdds WITH 9, ws WITH "13" FOR ROUND(StFar * 1.09, 2)= st2
REPLACE pdds WITH 0, ws WITH "19" FOR ROUND(StFar, 2)= st2
REPLACE ws WITH "15" FOR pdds=0 AND id_+";" $ "ATU;BE;GB;DE;EL;DK;EE;IE;ES;IT;CY;LV;LT;LU;MT;PL;PT;RO;SK;SI;HU;FI;FR;NL;CZ;SE;HR;"

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

*UPDATE fa2_ set pdds=9, ws="13" WHERE ROUND(StFar * 1.09, 2)= st2
SELECT * from fa2_  ;
  ORDER BY fas,fan ;
  INTO CURSOR fa2_ readw 

CALCULATE sum(1), sum(st2) TO lnTtlDkm, lnTtl  
  
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

lnTtl = ROUND(lnTtl, 2)

lcPath=GETDIR(SYS(5)+curdir(), "Изберете папка:", "Експорт на данни за продажбите:")
ERASE ( lcPath+"_exprt.zip" )
oZip = CreateObject("XStandard.Zip")
ozip.Pack(gcTmpDir+"_fa.dat",  lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_fa2.dat", lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_ktrg.dat", lcPath+"_exprt.zip", .f., "", 9)
*ozip.Pack(gcTmpDir+"_artkl.dat", lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_dkm.dat",  lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_exprt.dat",lcPath+"_exprt.zip", .f., "", 9)

MESSAGEBOX(TRAN(RECC("fa_"))+" документа, "+;
           TRAN(RECC("fa2_"))+" документореда"+CHR(13)+;
           TRAN(lnTtl)+" лв.вкл. ДДС"+CHR(13)+;
           TRAN(RECC("dkmp_"))+" документа платени в брой" ;
           ,64, "Данни за импортиране:")

lnCount = RECC("fa2_")+RECCOUNT("dkmp_")          
lcXml=chkUsg(_Usr, _eik, gDt1, lnCount) 
                    
CLEAR events          
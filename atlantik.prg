
lcFN = GETFILE("csv", "*.csv","Избери",0, "Документи за продажби:")
CREATE CURSOR s_ (cType n(1), efn c(13), fan n(10), Dt d, ;
   art c(10), q n(9,3), cn c(1), StFar n(12,2), vat c(1), st2 n(12,2))

SET DATE ymd
APPEND FROM (lcFN) DELIMITED WITH _ with CHARACTER ";"
SET DATE dmy
SELECT cType,efn, " "as fas, fan, dt, art, q, StFar, St2 FROM s_ ;
  WHERE NOT EMPTY(dt) into CURSOR s_ readw
REPLACE art with "Винетки" FOR art="0386"
REPLACE art with "0% ддс" FOR StFar=St2
REPLACE art with "1000" FOR BETWEEN(art, "0000", "9999")

SCAN
  lnSZ=LEN(ALLTRIM(efn))
  IF BETWEEN(lnSZ, 5, 8)
     REPLACE efn WITH REPLICATE("0", 9-lnSZ)+ALLTRIM(efn)
  ENDIF 
ENDSCAN
BROWSE FOR NOT EMPTY(efn)

*REPLACE fas WITH "к", fan WITH dt - DATE(YEAR(dt),1,1)+1 FOR EMPTY(efn)
CALCULATE MIN(dt), MAX(dt) TO gDt1, gDt2
REPLACE fas WITH "к", dt WITH gDt2, fan WITH gDt2 - DATE(YEAR(gDt2),1,1)+1 FOR EMPTY(efn)

SELECT ktrg
CALCULATE MAX(ktrg) TO lcID FOR ktrg="_01"
lnID_ = VAL(SUBSTR(lcID,4,7))+1

SELECT s_
CALCULATE MIN(dt),MAX(dt) to gDt1,gDt2

SELECT fas, fan, dt, efn, SPACE(10)as ktrg, ""as KName, ;
    art, sum(q)as q, sum(StFar)as StFar, sum(st2)as st2 ;
  FROM s_  ;
  group by fas,fan, art ;
  into CURSOR s_ readw

SELECT fas,fan,dt, efn, ktrg, ;
    IIF(fas="к", "501", "411")as skaw,sum(st2)as ttl FROM s_ ;
  GROUP BY fas,fan ;
  into CURSOR fa_ readw

*!*	SELECT 0
*!*	CREATE CURSOR k_ (f1 c(5), name c(40), mol c(20), id_dds c(15),efn c(13), f2 c(1), adr c(40))
*!*	APPEND FROM Clients.txt DELIMITED WITH tab
*!*	SELECT dist efn, SPACE(15)as id_DDS, SPACE(40)as name from s_ ;
*!*	  INTO CURSOR k_ readw

SELECT SPACE(10)as ktrg,SPACE(40)as name, efn, SPACE(15)as id_dds  FROM s_  ;
  WHERE NOT EMPTY(efn) ;
  GROUP BY efn ;
  INTO CURSOR ktrg_ readw

DO _ktrg

*!*	SCAN
*!*	  lcID=""
*!*	  SELECT ktrg
*!*	  IF NOT EMPTY(ktrg_.efn)
*!*	     LOCATE FOR efn=ktrg_.efn AND NOT EMPTY(ktrg)
*!*	     IF NOT EOF()
*!*	        lcID = ktrg.ktrg
*!*	        REPLACE name WITH ktrg.name,id_dds WITH ktrg.id_dds IN ktrg_
*!*	     ELSE
*!*	        DO case
*!*	        CASE LEN(TRIM(ktrg_.efn))=10 AND egn(ktrg_.efn)
*!*	          lcID="p"+SUBSTR(ktrg_.efn,1,9)
*!*	        CASE bstat(ktrg_.efn)
*!*	          lcID=SUBSTR(ktrg_.efn,1,8)
*!*	        ENDCASE 
*!*	     ENDIF 
*!*	  ENDIF 

*!*	** булстата е празен или невалиден
*!*	  IF NOT EMPTY(ktrg_.id_dds)
*!*	     LOCATE for id_dds = ktrg_.id_dds  AND NOT EMPTY(ktrg)
*!*	     IF NOT EOF()
*!*	        lcID=ktrg.ktrg
*!*	     ENDIF 
*!*	  ENDIF 
*!*	    
*!*	  IF EMPTY(lcID)  
*!*	      LOCATE FOR LOWER(name)=LOWER(SUBSTR(ktrg_.name,1,40)) ;
*!*	         AND EMPTY(efn)  AND NOT EMPTY(ktrg)
*!*	      IF NOT EOF()
*!*	         lcID=ktrg.ktrg
*!*	      ENDIF 
*!*	  ENDIF       
*!*	  IF EMPTY(lcID)
*!*	      lcID = "_01" + TRAN(lnID_,"@L 9999999")
*!*	      lnID_ = lnID_ + 1          
*!*	  ENDIF 

*!*	  REPLACE ktrg WITH lcID IN ktrg_           
*!*	  REPLACE ktrg WITH lcID for efn=ktrg_.efn AND fas=" "  IN fa_
*!*	ENDSCAN 

*!*	SELECT  ktrg_.efn, s_.fan, s_.dt FROM ktrg_ ;
*!*	  INNER JOIN s_ on s_.efn=ktrg_.efn ;
*!*	  where EMPTY(ktrg_.name) ;
*!*	  GROUP BY ktrg_.efn INTO CURSOR ktrg__
*!*	IF recc("ktrg__") > 0
*!*	  BROWSE TITLE "Нови клиенти:"
*!*	  COPY TO "Нови клиенти.txt" sdf
*!*	ENDIF 

*!*	SELECT ktrg_
*!*	COPY TO (gcTmpDir+"_ktrg.xxx")

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
COPY TO (gcTmpDir+"_fa.xxx")
*BROWSE

*IF thisform.chkDtjl.Value=1
IF .t.
	SELECT s_.art, SPACE(40)as name, SPACE(5)as mka, SPACE(10)as artkl ;
	  FROM s_  GROUP BY s_.art   ;
	  INTO CURSOR artkl_ 
	IF NOT FILE("art_.dat")
	   COPY TO art_.dat
	ELSE
	   USE art_.dat IN 0
	   SELECT artkl_.* from artkl_ ;
	     WHERE NOT exists (sele art from art_.dat a_ where a_.art=artkl_.art) ;
	     INTO CURSOR a__

	   SELECT art_
	   APPEND FROM DBF("a__")     
	ENDIF 
	USE IN artkl_
	   
	SELECT s_.art, NVL(artkl,STR(2,10))as artkl ;
	  FROM s_ ;
	  left join art_.dat on VAL(s_.art)=VAL(art_.art) ;
	  GROUP BY s_.art   ;
	  INTO CURSOR a_ readw 

*!*		LOCATE FOR EMPTY(artkl)
*!*		IF NOT EOF()
	   DO FORM art_
*!*		ENDIF    
	SELECT s_.*,NVL(artkl,STR(2,10))as artkl, ;
	    IIF(q=0, 000000.000, StFar / q)as cn ;
	  FROM s_ ;
	  left join art_.dat on s_.art=art_.art ;
	  INTO CURSOR s_
	  
	SELECT YEAR(dt)as gg, fas, fan, 000 as far,;
	  s_.artkl, q, cn, stFar, st2, 20 as pdds,  ;
	       NVL(artkl.ptd, "702")as ptdr, NVL(artkl.PtdSP, SPACE(40))as PtdSP ;
	  FROM s_  ;
	  left join artkl on s_.artkl=artkl.artkl ;
	  ORDER BY fas,fan    INTO CURSOR fa2_ readw

    REPLACE pdds WITH 0 FOR St2=StFar
ENDIF 

SELECT fa2_ 
CALCULATE sum(1), sum(st2) TO lnTtlDkm, lnTtl  
  
GO top
DO WHILE NOT EOF()
   lnFan=fan
   lnFar=1
   DO WHILE fan=lnFan AND NOT EOF()
      REPLACE far WITH lnFar
      lnFar = lnFar + 1 
      SKIP
   ENDDO 
ENDDO    

COPY TO (gcTmpDir+"_fa2.xxx")

***********
SELECT fas,fan, "pp"as td, 000000000 as NDok,000 as nr, dt, ;
    "411"+SPACE(7)+ktrg+STR(fan,10)as kredit, "501"as debit,;
    ttl as StDkm ;
  FROM fa_ where EMPTY(fas) ;
  GROUP BY fas,fan     INTO CURSOR dkmp_ readw

REPLACE ALL NDok WITH (dt-DATE(YEAR(Dt),1,1)+1)*10
  
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

COPY TO (gcTmpDir+"_dkm.xxx")

SELECT prmw.efn as efn, gDt1 as dt1, gDt2 as dt2, "01"as User,;
    MIN(fan)as fan1, MAX(fan)as fan2, ;
    m.lnTtl as ttl, lnTtlDkm as ttlDkm, RECC("fa_")as brF, 0 as brDkm ;
  FROM fa_ INTO curs _exprt

COPY TO (gcTmpDir+"_exprt.xxx")

*!*	** тест
*!*	SELECT fan, sum(st2)as st2 FROM fa2 ;
*!*	  GROUP BY fan INTO CURSOR _1

*!*	SELECT fan, sum(st2)as st2 FROM fa2_ ;
*!*	  GROUP BY fan INTO CURSOR _2

*!*	SELECT _1.*, _2.st2 as st2_ from _1 ;
*!*	  FULL JOIN _2 ON _1.fan=_2.fan ;
*!*	  having _1.st2 # _2.st2
*!*	**

* IF .f.
*!*		IF NOT tstLcnz()
*!*		   CLEAR EVENTS
*!*		   RETURN
*!*		ENDIF    

*	lcUsr = MLINE(FILETOSTR("c:\ptg\ptg-xp.lcnz"), 1)
    IF NOT tstLcnz(lcUsr, prmw.efn, gDt1)
       retu
    ENDIF 
    
    
*	lcXml=chkLLimit(lcUsr, prmw.efn, gDt1, RECC("fa2_")+RECCOUNT("dkmp_")) 
    lnCount = RECC("fa2_")+RECCOUNT("dkmp_")
	lcXml=chkUsg(_Usr, _eik, gDt1, lnCount) 

lcPath=GETDIR(SYS(5)+curdir(), "Изберете папка:", "Експорт на данни за продажбите:")

oZip = CreateObject("XStandard.Zip")
ozip.Pack(gcTmpDir+"_fa.xxx",  lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_fa2.xxx", lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_ktrg.xxx", lcPath+"_exprt.zip", .f., "", 9)
*ozip.Pack(gcTmpDir+"_artkl.xxx", lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_dkm.xxx",  lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_exprt.xxx",lcPath+"_exprt.zip", .f., "", 9)

MESSAGEBOX(TRANSFORM(RECC("fa_"))+" записа"+CHR(13)+;
          TRANSFORM(lnTtl)+" лв. с ДДС",64, "Данни за импортиране:")
          
CLEAR events          
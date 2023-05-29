SET DATE DMY 
lcFN = GETFILE("txt", "Файл:","Избери",0, "Документи за продажби:")
CREATE CURSOR s_ (fan n(10), Dt d, efn c(13), id_dds c(15), KName c(50),;
  st2 n(15,2),StFar n(15,2), St703 n(12,2), nPmnt n(12,2))

APPEND FROM (lcFN) DELIMITED WITH tab
*BROWSE

SELECT " "as fas,fan,dt, efn, SPACE(10)as ktrg, KName, ID_DDS, ;
    "411" as skaw,  st2 as ttl, StFar, St2, St703, St2*0 as St2703, nPmnt FROM s_ ;
  into CURSOR fa_ readw

REPLACE St2703 WITH St2 - ROUND(St2*(StFar - St703)/StFar,2) FOR St703 > 0
REPLACE St2 WITH St2 - St2703    FOR St703 > 0
REPLACE StFar WITH StFar - St703 FOR St703 > 0

SELECT SPACE(10)as ktrg,KName as name, efn, id_dds  FROM s_  ;
  WHERE NOT EMPTY(efn) ;
  GROUP BY efn ;
  INTO CURSOR ktrg_ readw

DO _ktrg  

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

SELECT YEAR(dt)as gg, " "as fas, fan, 000 as far,;
  STR(11,10)as artkl, 1 as q, St703 as cn, st703 as StFar, st2703 as st2, ;
  20 as pdds, "703  "as ptdr, SPACE(40)as PtdSP ;
  FROM fa_  ;
  WHERE St703 # 0 ;
  ORDER BY fas,fan    INTO CURSOR fa2__ readw

SELECT YEAR(dt)as gg, " "as fas, fan, 000 as far,;
  STR(10,10)as artkl, 1 as q, StFar as cn, stFar, st2, 20 as pdds,  ;
       "702  "as ptdr, SPACE(40)as PtdSP ;
  FROM fa_  ;
  WHERE StFar # 0 ;
  INTO CURSOR fa2_ readw

APPEND FROM DBF("fa2__") for st2#0

SELECT * from fa2_  ;
  ORDER BY fas,fan ;
  INTO CURSOR fa2_ readw 

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

*!*	IF .f.
*!*	IF NOT tstLcnz()
*!*	   CLEAR EVENTS
*!*	   RETURN
*!*	ENDIF    

lcUsr = MLINE(FILETOSTR("c:\ptg\ptg-xp.lcnz"), 1)
lcXml=chkUsg(lcUsr, prmw.efn, gDt1, RECC("fa2_")+RECCOUNT("dkmp_")) 

*WAIT TRANS(RECC("fa2_")+RECC("dkmp_"))

*!*	STRTOFILE(lcXml, "xml.xml")
*!*	loXML = CREATEOBJECT('MSXML2.DomDocument')  
*!*	loXML.ASYNC = .f.
*!*	loXML.LOADXML(lcXml)
*!*	loRoot = loXML.documentElement

*!*	lnCount = val( loRoot.getElementsByTagName("LimitRecords").item[0].nodeTypedValue )

*!*	IF lnCount < 1
*!*	   MESSAGEBOX("Няма лиценз за "+prmw.efn,48, "Проверка на лиценза:")
*!*	   CLEAR EVENTS 
*!*	   RETURN
*!*	ENDIF    
*!*	ENDIF 
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

MESSAGEBOX(TRANSFORM(RECC("fa_"))+" записа"+CHR(13)+;
          TRANSFORM(lnTtl)+" лв. с ДДС",64, "Данни за импортиране:")
          
CLEAR events          
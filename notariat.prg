
SELECT 0
IF NOT USED("artkl")
   USE artkl
ELSE
   SELECT artkl
ENDIF       
LOCATE FOR EMPTY(PtdSP) AND ptd="70"
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
CREATE CURSOR fa_ (cFAN c(10), Dt d, cTxt c(30), ;
   cTtl c(12), cVAT c(12),;
   efn c(14), id_dds c(15), KName c(40))

lcXML = GETFILE("xml", "Файл:","Избери",0, "Документи за продажби:")
lcXml=FILETOSTR(lcXML)

loXML = CREATEOBJECT('MSXML2.DomDocument')  
loXML.ASYNC = .f.
loXML.LOADXML(lcXml)
CLEAR
SET DATE YMD 
loRoot = loXML.documentElement
loRoot_ = loRoot.GetElementsByTagName("Operations").item[0]

FOR EACH loNode IN loRoot_.childNodes
  APPEND BLANK IN fa_
*  REPLACE Dt WITH loNode.getAttribute("Date")
  REPLACE cTxt WITH STRCONV( NVL(loNode.getAttribute("Term"),""), 11)  
  REPLACE cTtl WITH loNode.getAttribute("DocumentSumVATIncl")
  REPLACE cVAT WITH loNode.getAttribute("DocumentVAT")

  lo = loNode.GetElementsByTagName("Company").item[0]
  repl KName with  STRCONV( lo.getAttribute("Name"), 11)  
  REPLACE efn    WITH NVL(lo.getAttribute("Bulstat"),"")
  REPLACE ID_DDS WITH NVL( lo.getAttribute("VatNumber"),"")
  
  lo = loNode.GetElementsByTagName("Document").item[0]
  REPLACE Dt WITH CTOD( lo.getAttribute("Date") )
  REPLACE cFan WITH lo.getAttribute("Number")
    
Next

SET DATE DMY

SELECT VAL(cfan)as fan,dt,efn,KName,id_dds, ;
    VAL(cTtl )as st2, VAL(cTtl )-VAL(cVAT)as StFar, 20 as pdds ;
  from fa_ into curs fa_ readw

SELECT " "as fas,fan,dt, 00000 as NDok, {} as dtz, ;
    PADR(allTRIM(efn),13," ")as efn, SPACE(10)as ktrg, KName, ;
    PADR(ALLTRIM(id_dds),15," ")as ID_DDS, ;
    0 * St2 as nPmnt, "411" as skaw, ;
    st2 as ttl,  StFar, St2 FROM fa_ ;
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

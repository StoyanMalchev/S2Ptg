CLOSE DATABASES all
*SET DEFAULT TO ("d:\temp\rb_clock\")
USE prmw

SELECT 0
lcFN = GETFILE("csv", "*.csv","Избери",0, "Документи за продажби:")
CREATE CURSOR s_ (f1 c(1), t1 c(10),t2 c(2), t3 c(20), fan n(10),;
   f2 c(1), Dt d,  efn c(13), KName c(40), id_dds c(15),;
   f3 c(1), f4 c(1), f5 c(1), f6 c(1), ;
   id_setvice c(10), art c(10), nameArt c(40), ;
   pdds n(2), ww c(3), q n(9,3), StFar n(12,2), nVAT n(12,2),;
   f9 c(1), ptype c(20), ska50 c(10), pww c(5), nPmnt n(12,2))

*APPEND FROM (lcFN) csv
APPEND FROM (lcFN) DELIMITED WITH _ with CHARACTER ","
*!*	BROWSE FIELDS fan,dt,KName, art,nameART
*!*	retu

SELECT " "as fas,fan,dt,  nPmnt, ska50  FROM s_ ;
  where t1="doc_payment" AND ;
       (ska50="501"or (ska50="503" AND ptype="card") );
  into CURSOR p_ readw

SELECT s_
REPLACE ALL art WITH LTRIM(art)
  
SELECT " "as fas,fan,dt, ;
    efn, SPACE(10)as ktrg, KName, ID_DDS, ;
    "411" as skaw, art, sum(q)as q, ;
    sum(StFar)as StFar, sum(StFar+nVAT)as st2, pdds ;
  FROM s_ where t1="doc_line" ;
  GROUP BY fan, art ;
  into CURSOR fa_ readw

REPLACE efn WITH "" FOR NOT (bstat(efn)or egn(efn)) 
REPLACE efn WITH SUBSTR(id_dds,3,9), KName WITH "" ; 
  FOR bstat( SUBSTR(id_dds,3,9)) AND EMPTY(efn)

SELECT dist KName as name FROM fa_ where EMPTY(efn) INTO CURSOR _ktrg
    
SELECT SPACE(10)as ktrg,KName as name, efn, id_dds ;
      FROM fa_  ;
  WHERE NOT EMPTY(efn) ;
  GROUP BY efn ;
  INTO CURSOR ktrg_ readw

APPEND FROM DBF("_ktrg")

REPLACE id_dds WITH "BG"+id_dds FOR id_dds=efn AND LEN(TRIM(id_dds))=9

SELECT FA_
REPLACE id_dds WITH "" FOR NOT SUBSTR(id_dds,1,2)+";" $ ;
  "BG;AT;BE;GB;DE;EL;DK;EE;IE;ES;IT;CY;LV;LT;LU;MT;PL;PT;RO;SK;SI;HU;FI;FR;NL;CZ;SE;HR;"

SELECT KTRG_
REPLACE id_dds WITH "" FOR NOT SUBSTR(id_dds,1,2)+";" $ ;
  "BG;AT;BE;GB;DE;EL;DK;EE;IE;ES;IT;CY;LV;LT;LU;MT;PL;PT;RO;SK;SI;HU;FI;FR;NL;CZ;SE;HR;"

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

SELECT * from fa_ group BY fas, fan INTO CURSOR fa1_
COPY TO (gcTmpDir+"_fa.dat")

********
IF NOT FILE("art_.dat")
   SELECT 0
   CREATE TABLE art_.dat FREE ;
     (art c(20), name c(40), nameArt c(40), mka c(5), artkl c(10))
ENDIF 

SELECT  art, nameArt as name, nameArt from s_ ;
  GROUP BY art   into CURSOR artkl_

SELECT artkl_.*, NVL(artkl,SPACE(10))as artkl, NVL(mka, SPACE(5))as mka ;
  FROM artkl_ left JOIN artkl ON artkl.artkl=artkl_.art ;
  ORDER BY artkl INTO CURSOR artkl_ readw

SELECT artkl_.* from artkl_ ;
  WHERE NOT exists (sele art from art_.dat a_ where a_.art=artkl_.art) ;
     INTO CURSOR a__
SELECT art_
APPEND FROM DBF("a__")     
USE IN artkl_  
      
DO FORM art_

SELECT art_
LOCATE FOR art=" " AND NOT EMPTY(artkl)
lcArtkl=art_.artkl
SELECT YEAR(dt)as gg," "as fas,fan, 001 as far,;
    NVL(art_.artkl, lcArtkl)as artkl, q, IIF(q=0, 0,StFar) as cn,StFar,st2,"11"as ws, ;
    SUBSTR(id_DDS,1,2)as id_ ;
  FROM fa_  ;
  LEFT JOIN art_ on art_.art=fa_.art AND NOT EMPTY(art_.artkl) ;
  WHERE YEAR(dt)=YEAR(gDt1) ;
  ORDER BY fas,fan    INTO CURSOR fa2_ readw

SELECT fa2_.*, a_.pdds, a_.PtdSP, a_.Ptd as PtdR ;
  FROM fa2_ ;
  LEFT JOIN artkl a_ on a_.artkl=fa2_.artkl ;
  ORDER BY fas,fan ;
  INTO CURSOR fa2_ readw 

REPLACE pdds WITH 9, ws WITH "13" FOR BETWEEN( ROUND(StFar * 1.09, 2) - st2, -0.02, 0.02)
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
*lcPtd501 = "501"+SPACE(37)

SELECT p_.*, fa1_.ktrg FROM p_ ;
  INNER JOIN fa1_ on fa1_.fan=p_.fan ;
  INTO CURSOR p_

SELECT fas,fan, "pp"as td, 000000000 as NDok,000 as nr, dt, ;
    "411"+SPACE(7)+ktrg+STR(fan,10)as kredit, ska50 as debit,;
    nPmnt as StDkm, ska50 ;
  FROM p_ where EMPTY(fas) AND nPmnt # 0 ;
  GROUP BY fas,fan     INTO CURSOR dkmp_ readw

REPLACE debit WITH kredit FOR StDkm < 0
REPLACE kredit WITH ska50, StDkm WITH -StDkm FOR StDkm < 0

REPLACE ALL NDok WITH 100000+(dt-DATE(YEAR(Dt),1,1)+1)*100 FOR ska50="501"
REPLACE ALL NDok WITH 200000+(dt-DATE(YEAR(Dt),1,1)+1)*100 FOR ska50="503"
  
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
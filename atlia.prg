CLOSE DATABASES all
*SET DEFAULT TO d:\temp\acserv-atlia\

lcFN=GETFILE("xls","Фактури")
IMPORT FROM (lcFN) xl5
lcAlias=ALIAS()
REPLACE p WITH "9990000000" FOR STR(VAL(p),10) # PADL(ALLTRIM(p),10," ")

*BROWSE FIELDS n,o,p,q,u, y, aa FOR STR(VAL(y)*1.2,12,2) # STR(VAL(aa),12,2)

REPLACE all y WITH STRTRAN(y," ",""), aa WITH STRTRAN(aa," ","")

SELECT SPACE(10)as ktrg, c as KName, d as efn, e as id_dds, ;
  " "as fas, ROUND(VAL(r),0)as fan, ;
  CTOD(s) as dt, STR(VAL(t),10)as art_, u as aname, ;
  VAL(y) as q, VAL(ac)as stfar, VAL(ae)as st2  ;
  FROM (lcAlias) INTO CURSOR s_
    
SELECT * from s_ where  BETWEEN(YEAR(dt), 2021, 2025) ;
  INTO CURSOR s_ readw
  
CALCULATE MIN(dt), MAX(dt) TO gDt1, gDt2
IF MONTH(gDt1)#MONTH(gDt2) OR YEAR(gDt1)#YEAR(gDt2)
   gdt1 = CTOD(INPUTBOX("От дата:","Период:",TRANSFORM(gDt1)))
   gDt2 = CTOD(INPUTBOX("От дата:","Период:",TRANSFORM(gDt2)))
   SELECT * from s_ where BETWEEN(dt, gDt1, gDt2)  ;
     INTO CURSOR s_ readw
ENDIF

SELECT 0
lcFN=GETFILE("xls", "Плащания:")
IMPORT FROM (lcFN) xl5
lcAlias=ALIAS()
SELECT val(c) as fan, " "as fas, VAL(j)as nPMNT, LOWER(s)as cPmnt, t as VidDkm ;
  FROM (lcAlias) INTO CURSOR p_
*SELECT * from p_ where LOWER(cPmnt)="в брой" AND LOWER(vidDkm)="фактура" INTO CURSOR p_

SELECT s_.*,"411"as skaw, vidDKM, nPMNT, cPmnt FROM s_ ;
  left JOIN p_ on p_.fan=s_.fan ;
  INTO CURSOR s_ readw

*SELECT fan,stfar, st2  FROM s_ where STR(fan,10)="21" ORDER BY fan
*COPY TO d:\temp\21 xl5

REPLACE fas WITH "к", efn WITH "", dt WITH gDt2 FOR STR(fan,10)="21" 

REPLACE fan WITH 1000000000+MONTH(gDt1), ktrg WITH "_000000100" ;
   FOR STR(fan,10)="21" AND NOT ("наложен"$ cPMNT OR "в брой"$ cPMNT)
REPLACE fan WITH 2000000000+MONTH(gDt1), skaw WITH "422 " ;
   FOR STR(fan,10)="21" AND "наложен"$ cPMNT
REPLACE fan WITH 3000000000+MONTH(gDt1), skaw WITH "501" ;
   FOR STR(fan,10)="21" AND "в брой"$ cPMNT

REPLACE fan WITH 4000000000+MONTH(gDt1)   FOR STR(fan,10)="21" 

*BROWSE FIELDS fan,dt,efn,st2,vidDKM, cPmnt
*retu

* SELECT * from s_ where fas="к"

SELECT fas,fan,dt, PADR(LTRIM(efn),13," ")as efn,ktrg,skaw,sum(st2)as ttl FROM s_ ;
  GROUP BY fas,fan   into CURSOR fa_ readw

*!*	SELECT efn,kname FROM fa_ group BY efn INTO CURSOR k_ readw
*!*	SELECT * from k_ ;
*!*	  where NOT EMPTY(efn)and NOT (bstat(efn)or egn(efn)) ;
*!*	  INTO CURSOR tst_
*!*	IF RECCOUNT("tst_") > 0
*!*	   BROWSE TITLE "Некоректен ЕИК:"
*!*	ENDIF 

SELECT fa_
REPLACE efn WITH "" FOR NOT (bstat(efn)or egn(efn)) 

SELECT SPACE(10)as ktrg,SPACE(40)as name, efn, SPACE(15)as id_dds  FROM s_  ;
  WHERE NOT EMPTY(efn) ;
  GROUP BY efn ;
  INTO CURSOR ktrg_ readw

USE ktrg ORDER ktrg IN 0  

DO _ktrg

SELECT fa_
GO top
SCAN
   IF fas=" "and EMPTY(ktrg)
      SELECT ktrg
      LOCATE FOR efn=SUBSTR(fa_.efn,1,13)
      IF eof() OR EMPTY(efn)
         REPLACE ktrg WITH "_000000100" IN fa_
      ELSE
         REPLACE ktrg WITH ktrg.ktrg IN fa_      
      ENDIF    
   ENDIF 
ENDSCAN 

SELECT fa_
COPY TO (gcTmpDir+"_fa.dat") 

SELECT * from p_ where LOWER(cPmnt)="в брой" AND LOWER(vidDkm)="фактура" INTO CURSOR p_
SELECT fa_.*, nPmnt  FROM fa_ inner JOIN p_ on p_.fan=fa_.fan ;
  INTO CURSOR p__

SELECT fas,fan, "pp"as td, 000000000 as NDok,000 as nr, dt, ;
    "411"+SPACE(7)+ktrg+STR(fan,10)as kredit, "501"as debit, nPmnt as StDkm ;
  FROM p__ ;
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
SUM StDkm TO lnTTL501

SELECT fa_
CALCULATE MIN(dt), MAX(dt) TO gDt1, gDt2

*!*	SELECT * from fa_
*!*	SELECT * from s_
* retu

*SELECT * from fa_ order BY art_ desc
  
* _path="d:\temp\acserv-atlia\"
*!*	SELECT NVL(artkl, SPACE(10))as artkl, ;
*!*	       NVL(name, SPACE(40))as name, ;
*!*	    fa_.art_,aname from fa_  ;
*!*	  left JOIN d:\temp\acserv-atlia\artkl on VAL(fa_.art_)=VAL(artkl)  ;
*!*	  INTO CURSOR art_
  
SELECT s_.art_ as art, AName as name ;
  FROM s_  GROUP BY art   ;
  INTO CURSOR artkl_ 

*!*	SELECT artkl_.*, NVL(artkl,SPACE(10))as artkl, NVL(mka, SPACE(5))as mka ;
*!*	  FROM artkl_ left JOIN artkl ON artkl.artkl=artkl_.art ;
*!*	  ORDER BY artkl INTO CURSOR artkl_ readw

*!*	IF NOT FILE("art_.dat")
*!*	   COPY TO art_.dat
*!*	   USE art_.dat IN 0	   
*!*	ELSE
*!*	   USE art_.dat IN 0
*!*	   SELECT artkl_.* from artkl_ ;
*!*	     WHERE NOT exists (sele art from art_.dat a_ where a_.art=artkl_.art) ;
*!*	     INTO CURSOR a__

*!*	   SELECT art_
*!*	   APPEND FROM DBF("a__")     
*!*	ENDIF 
*!*	USE IN artkl_  
      
*** DO FORM art_

SELECT artkl_.art as artkl, name, 20 as pdds, ;
    "304"+STR(1,17)+art as PtdSP, "702"+STR(1,17) as ptd     from artkl_ ;
  WHERE NOT exists ;
    (sele artkl from artkl a_ where a_.artkl=artkl_.art) ;
  INTO CURSOR a__

BROWSE FIELDS artkl,name  TITLE "НОВИ АРТИКУЛИ:"  
  
COPY TO (gcTmpDir+"_artkl.dat") 


*!*	SELECT artkl
*!*	LOCATE FOR artkl="9990000000"
*!*	IF EOF()
*!*	   INSERT INTO artkl (artkl,Name,pdds,Ptd) ;
*!*	     VALUES ("9990000000", "Други", 20, "702")
*!*	ENDIF 
*!*	COPY TO (gcTmpDir+"_artkl.dat") FOR .f.
     
*!*	   SELECT s_.*,NVL(artkl,"9990000000")as artkl, ;
*!*		  IIF(q=0, 000000.000, StFar / q)as cn ;
*!*		  FROM s_ ;
*!*		  left join art_.dat on s_.art_=art_.art AND NOT EMPTY(art_.art) ;
*!*		  INTO CURSOR s_
   SELECT s_.*, art_ as artkl, ;
	  IIF(q=0, 000000.000, StFar / q)as cn ;
	  FROM s_ 	  INTO CURSOR s_
	
	lcPtd702="702"+STR(1,17)
	SELECT YEAR(dt)as gg, fas, fan, 000 as far,;
	  s_.artkl, q, cn, stFar, st2, 20 as pdds,  ;
	       NVL(artkl.ptd, lcPtd702)as ptdr, NVL(artkl.PtdSP, SPACE(40))as PtdSP, efn ;
	  FROM s_  ;
	  left join artkl on s_.artkl=artkl.artkl ;
	  ORDER BY fas,fan    INTO CURSOR fa2_ readw

    REPLACE ptdR WITH "702"+STR(2,17) FOR STR(fan,10)="20" OR EMPTY(efn)

    REPLACE pdds WITH 0 FOR St2=StFar


*!*	SELECT fan,artkl,q,stfar,st2, ROUND(st2 - stFar*1.2,2) ;
*!*	  from fa2_ where q=0 OR ROUND(stFar*1.2,2) # ROUND(st2,2) OR pdds#20

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

COPY TO (gcTmpDir+"_fa2.dat")

USE prmw IN 0

SELECT prmw.efn as efn, gDt1 as dt1, gDt2 as dt2, "01"as User,;
    MIN(fan)as fan1, MAX(fan)as fan2, ;
    m.lnTtl as ttl, lnTtlDkm as ttlDkm, RECC("fa_")as brF, 0 as brDkm ;
  FROM fa_ INTO curs _exprt

COPY TO (gcTmpDir+"_exprt.dat")

*SET DEFAULT TO d:\vfp\s2ptg

lcPath=GETDIR(SYS(5)+curdir(), "Изберете папка:", "Експорт на данни за продажбите:")

oZip = CreateObject("XStandard.Zip")
ozip.Pack(gcTmpDir+"_fa.dat",  lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_fa2.dat", lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_ktrg.dat", lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_artkl.dat", lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_dkm.dat",  lcPath+"_exprt.zip", .f., "", 9)
ozip.Pack(gcTmpDir+"_exprt.dat",lcPath+"_exprt.zip", .f., "", 9)

MESSAGEBOX("Период:"+TRAN(gDt1)+" до "+TRAN(gDt2)+CHR(13)+;
    TRAN(RECC("fa_"))+" документа"+CHR(13)+;
    TRAN(lnTtl)+" лв. с ДДС"+CHR(13)+;
    TRAN(RECC("dkmp_"))+" платени в брой"+STR(lnTTL501,12,2)+" лв.",;
           64, "Данни за импортиране:")

retu

*!*	lcFN=GETFILE("xls", "Данни за плащанията")
*!*	IMPORT FROM (lcFN) xl5
*!*	lcAlias=ALIAS()
*!*	SELECT c as fan, i as stst, s as cPmnt, t as VidDkm FROM (lcAlias) INTO CURSOR p_
*!*	SELECT * from p_ where LOWER(cPmnt)="в брой" AND LOWER(vidDkm)="фактура" INTO CURSOR p_

*!*	retu

CLOSE DATABASES all
SET DELETED ON
SET EXCLUSIVE OFF
SET TALK OFF
SET SAFETY OFF
SET DATE DMY
SET CENTURY on
SET POINT TO "."
*ON SHUTDOWN CLEAR EVENTS
ON ERROR DO ErrLog WITH ERROR(),PROG(),LINE(), MESSAGE(),MESSAGE(1), SYS(2018)

PUBLIC gcTmpDir
gcTmpDir = addbs(SYS(2023)) + "_exprt\"
IF NOT DIRECTORY(gcTmpDir)
   MD (gcTmpDir)
ENDIF 
ERASE ( gcTmpDir+"*.*" )

PUBLIC _path1, _path, _eik, _usr
_Path1=SYS(5)+curdir()
_usr = MLINE(FILETOSTR("c:\ptg\ptg-xp.lcnz"), 1)

select name, path, dt_tm from (_Path1+"db") ;
  order by dt_tm desc into cursor db_ 
_Path = path
_Path = INPUTBOX("Папка с данни ПП ПИТАГОР:", db_.name, _Path)
SET DEFAULT TO ( _path )

IF NOT FILE("ImportS.lcnz")
   MESSAGEBOX("Липсва ImportS.lcnz", 16, "Прекъсване!")
   QUIT 
ENDIF    

USE prmw
_eik = prmw.efn

*cMode = "-Gensoft"
cMode = MLINE(FILETOSTR("ImportS.lcnz"), 2)

USE artkl in 0
USE smpl IN 0
USE ktrg ORDER ktrg IN 0

*!*	DO Notarius
*!*	retu


DO case
CASE cMode=="-GenSoft"
   DO Gensoft
CASE cMode=="-GenSoft2"
   DO Gensoft2
CASE cMode="-InvBg" 
**   OR  INLIST(prmw.efn, "203782156", "205313481", "206641417")	&& inv.bg
   DO Inv_bg
CASE  cMode="-Atlantik"
**   or  prmw.efn="125038159"
   DO Atlantik
CASE  cMode="-Mega"
**   or prmw.efn="835037338"
   DO Mega
CASE  cMode="-Atlia"
**   or prmw.efn="131051990"
   DO Atlia
CASE cMode=="-NKSoft"
   DO NKSoft
CASE cMode=="-Clock"
   DO Clock
CASE cMode=="-Tonegan"
   DO Tonegan
CASE cMode=="-Notariat"
   DO Notariat
CASE cMode=="-MicroBg2"
   DO MicroBg2
OTHERWISE 
   MESSAGEBOX(prmw.efn+" "+prmw.name,16,"Няма алгоритъм &cMode за тази фирма")
ENDCASE    

*READ events
retu

****************
PROCEDURE ws_art
IF NOT USED("artkl")
   USE artkl IN 0
ENDIF 
IF TYPE("artkl.ws") # "C"
   retu
ENDIF  

SELECT fa2_.*, NVL(artkl.ws, "  ")as _ws ;
  FROM fa2_ left JOIN artkl ON artkl.artkl=fa2_.artkl ;
  INTO CURSOR fa2_ readw

REPLACE ws WITH _ws FOR pdds= 0 and NOT EMPTY(_ws)
retu
  

***************
PROCEDURE _ktrg
IF NOT USED("ktrg")
   USE ktrg order ktrg IN 0
ENDIF    

SELECT ktrg
CALCULATE MAX(ktrg) TO lcID FOR ktrg="_01"
lnID_ = VAL(SUBSTR(lcID,4,7))+1

SELECT ktrg_
SCAN
  lcID=""
  SELECT ktrg
  IF NOT EMPTY(ktrg_.efn)
     LOCATE FOR efn=ktrg_.efn AND NOT EMPTY(ktrg)
     IF NOT EOF()
        lcID = ktrg.ktrg
        REPLACE name WITH ktrg.name,id_dds WITH ktrg.id_dds IN ktrg_
     ELSE	&& Нов клиент:
        DO case
        CASE LEN(TRIM(ktrg_.efn))=10 AND egn(ktrg_.efn)
          lcID="p"+SUBSTR(ktrg_.efn,1,9)
        CASE bstat(ktrg_.efn)
**          lcID=SUBSTR(ktrg_.efn,1,8)
**
	        lcIdx=Space(2)
	        DO while .t.
	           lcID=Substr(ktrg_.efn,1,8)+lcIdx
	           Seek lcID
	           If Eof()
				  Insert into ktrg ( ktrg ) values ( lcID )
	 	          exit              
	           EndIf 
	           lcIdx = Padr(Tran(Val(lcIdx)+1),2," ")
	        ENDDO
	    OTHERWISE 
**  ???
        ENDCASE 
     ENDIF 
  ENDIF 

** булстата е празен или невалиден
  IF EMPTY(lcID)and NOT EMPTY(ktrg_.id_dds)
     LOCATE for id_dds = ktrg_.id_dds  AND NOT EMPTY(ktrg)
     IF NOT EOF()
        lcID=ktrg.ktrg
     ENDIF 
  ENDIF 
    
  IF EMPTY(lcID)  
      LOCATE FOR LOWER(name)=LOWER(SUBSTR(ktrg_.name,1,40)) ;
         AND EMPTY(efn)  AND NOT EMPTY(ktrg)
      IF NOT EOF()
         lcID=ktrg.ktrg
      ENDIF 
  ENDIF       
  IF EMPTY(lcID)
      lcID = "_01" + TRAN(lnID_,"@L 9999999")
      lnID_ = lnID_ + 1          
  ENDIF 

  REPLACE ktrg WITH lcID IN ktrg_   
  
  IF EMPTY(ktrg_.efn)and EMPTY(ktrg_.id_dds) 
     REPLACE ktrg WITH lcID for KName=ktrg_.Name AND fas=" "  IN fa_  
  ELSE 
     REPLACE ktrg WITH lcID ;
       for efn=ktrg_.efn AND id_dds=ktrg_.id_dds AND fas=" "  IN fa_

     IF NOT EMPTY(ktrg_.efn)
** заради Тонеган - ако няма ДДС номер в данните
	     REPLACE ktrg WITH lcID for efn=ktrg_.efn AND fas=" "  IN fa_
     ENDIF 
  ENDIF 
  _vfp.StatusBar = lcID
ENDSCAN 

*!*	SELECT  ktrg_.efn, s_.fan, s_.dt, s_.KName FROM ktrg_ ;
*!*	  INNER JOIN s_ on s_.efn=ktrg_.efn ;
*!*	  where EMPTY(ktrg_.name) ;
*!*	  GROUP BY ktrg_.efn INTO CURSOR ktrg__

SELECT ktrg_.*, NVL(ktrg.name, "")as _KName  from ktrg_ ;
  left JOIN ktrg ON ktrg.ktrg=ktrg_.ktrg ;
  HAVING EMPTY(_KName) ;
  INTO CURSOR ktrg__
  
*!*	SELECT  ktrg_.*, NVL(ktrg.name, "") as _KName FROM ktrg_ ;
*!*	  INNER JOIN fa_ ;
*!*	    on fa_.efn=ktrg_.efn AND fa_.id_dds=ktrg_.id_dds ;
*!*	      AND fa_.KName=SUBSTR(ktrg_.name,1,30) ;
*!*	  jeft JOIN ktrg ON ktrg.ktrg=ktrg_.ktrg ;
*!*	  GROUP BY ktrg_.efn,ktrg_.id_dds,ktrg_.name ;
*!*	  INTO CURSOR ktrg__  
    
IF recc("ktrg__") > 0
  BROWSE TITLE "Нови клиенти:"
  COPY TO "Нови клиенти.txt" sdf
ENDIF 

SELECT ktrg_
COPY TO (gcTmpDir+"_ktrg.dat")
COPY TO (gcTmpDir+"_ktrg.xxx")

RETURN


**************  
Function BStat  
Parameters pEFN
priv lcEFN, lcVal, j
lcEFN = substr(alltrim(pEFN),1,9)
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


***************
FUNCTION ErrLog
PARAMETER M.ERRNUM, M.PROGRAM, M.LINE, M.MESS, M.MESS1, M.PARAM

_SCREEN.VISIBLE = .T.
SET ESCAPE OFF
DO CASE
CASE  Wexist("Trace")   
   SUSPEND
CASE M.ERRNUM = 1707	&& * нет индексного файла
  RETRY	&& * тогда мы попробуем еще раз
CASE M.ERRNUM=109	&& * запись заблокирована другим пользователем
  IF MESSAGEBOX("Блокиран запис от друг потребител", 48 + 1, WTITLE("")) = 1
     RETRY
  ELSE
     RETURN
  ENDIF
ENDCASE

PRIVATE K, I
SET safety OFF 

= MESSAGEBOX("Metod:"+m.program+" Ред "+STR(M.Line)+" "+M.MESS+'"', 48, "Грешка")

QUIT
CLEAR EVENTS 

****************
FUNCTION tstLcnz (cUsr, cEFN, _dt)
lcLcnz =FILETOSTR("ImportS.lcnz")
lcStr = alltrim(MLINE(lcLcnz,1))+alltrim(MLINE(lcLcnz,2))+;
    alltrim(MLINE(lcLcnz,3))+ alltrim(MLINE(lcLcnz,4))

lcSign = SYS(2007, lcStr)

IF alltrim(MLINE(lcLcnz,1))== cUsr and ;
    alltrim(MLINE(lcLcnz,3))== TRIM(cEFN) and ;
     alltrim(MLINE(lcLcnz,5))== lcSign     
ELSE      
   MESSAGEBOX("Некоректен ImportS.lcnz", 48, "Проблем:")
   RETURN .f.
ENDIF 

_dtL = ctod(MLINE(lcLcnz,4))
IF _dtL < _dt     
   MESSAGEBOX("Бевалиден лиценз след "+TRANSFORM(_dtL), 48, "Проблем:")
   RETURN .f.
ENDIF    

RETURN .t.


******************
*FUNCTION chkLLimit
FUNCTION chkUsg
PARAMETERS cUSER, cEIK, Dt1, nRecords
PRIVATE xh, lcDt
clear

lcDt=DTOS(m.dt1)
lcDt=LEFT(lcDt,4)+"-"+SUBSTR(lcDt,5,2)+'-'+RIGHT(lcDt,2)

* POST - data 
SET MEMOWIDTH TO 160
SET TEXTMERGE on
TEXT TO lcEnv NOSHOW 
<?xml version="1.0" encoding="UTF-8" standalone="no"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
            xmlns:xsd="http://www.w3.org/2001/XMLSchema" 
           xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
<soap:Body>
<CheckLicenseLimit xmlns="http://pitagor-bg.com/">
<aUserId><<cUSER>></aUserId>
<aPartnerEIK><<cEIK>></aPartnerEIK>
<aDate><<lcDt>></aDate>
<aRecords><<TRANSFORM(nRecords)>></aRecords>
</CheckLicenseLimit>
</soap:Body>
</soap:Envelope>
ENDTEXT
*STRTOFILE(lcEnv, "d:\temp\soap.xml")
lcOnErr = ON("error")
ok=.t.
ON ERROR ok=.f.
xh = CREATEOBJECT("Msxml2.XMLHTTP.6.0")
xh.Open("POST", "http://pitagor-bg.com/service.asmx","False")
IF xh.ReadyState # 1 OR NOT ok
	ON ERROR &lcOnErr
    WAIT "I can't open web services" WINDOW AT 0,0 
    xh=.null.
    RELEASE xh
    RETURN ""
ENDIF

xh.setRequestHeader ("Content-Type", "text/xml; charset=utf-8")
xh.setRequestHeader ("SOAPAction", "http://pitagor-bg.com/CheckLicenseLimit")

xh.send(lcEnv)
iF (xh.readyState <> 4) OR NOT ok
    WAIT "Error at comunication with WS!" WINDOW AT 0,0 TIMEOUT 3
    this.lcResponse=.f.
    xh=.null.
    RELEASE xh
    RETURN ""
ELSE
    lcXML = xh.responseXML.xml
    xh=.null.
    RELEASE xh
    RETURN lcXML
ENDIF 

****************
*!*	FUNCTION tstLcnz
*!*	ok=.t.
*!*	loHTTP = CREATEOBJECT('WinHttp.WinHttpRequest.5.1')
*!*	lcID = MLINE(FILETOSTR("c:\PTG\id.ptg"), 1)
*!*	lcURL = "http://pitagor-bg.com/Service.aspx?Method=GetLicense&PCId="+lcID
*!*	loHTTP.Open("POST", lcURL, .f.)
*!*	loHttp.setrequestheader("Content-Type", "text/html")
*!*	lcTmpFile =  ADDBS(SYS(2023))+"tmp.txt"
*!*	lcErrMsg = ""

*!*	If ok   
*!*	   loHTTP.Send( " ???" )
*!*	   lcTxt = loHTTP.ResponseBody
*!*	   IF TYPE("lcTxt") # "C"
*!*	      RETURN .f.
*!*	   ENDIF 
*!*	ELSE
*!*	   RETURN .f.
*!*	EndIf


*!*	STRTOFILE( lcTxt, lcTmpFile)
*!*	SELECT 0
*!*	Create Cursor tst_ (user c(10),UsrName c(40),pcN n(5), item c(10),;
*!*	      dt1 d, opt c(20),xxx c(1), dtRls d,dtc c(10),opt2 c(10), csum c(16))
*!*	APPEND FROM (lcTmpFile ) DELIMITED WITH _ WITH CHARACTER "|"
*!*	SELECT * from tst_ WHERE item="IMPORT" INTO CURSOR tst_
*!*	IF dtRls > DATE()
*!*	   RETURN .t.
*!*	ENDIF 

*!*	MESSAGEBOX("Некоректен лиценз!", 48, "Импорт на данни:")

*!*	RETURN dtRls > DATE() - 50


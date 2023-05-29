CLEAR
*= soap("Set", "PTG", "125038839", DATE(2022,1,12), 77766)
??

?? soap("Check", "PTG", "125038839", DATE(), 654)


*************
FUNCTION soap
PARAMETERS cSetChk, cUSER, cEIK, Dt1, nRecords
*     http://pitagor-bg.com/CheckLicenseLimit
* или http://pitagor-bg.com/SetLicenseLimit

PRIVATE xh, lcDt
clear

lcDt=DTOS(dt1)
lcDt=LEFT(lcDt,4)+"-"+SUBSTR(lcDt,5,2)+'-'+RIGHT(lcDt,2)

* POST - data 
SET MEMOWIDTH TO 160
SET TEXTMERGE on
TEXT TO lcEnv NOSHOW 
<?xml version="1.0" encoding="UTF-8" standalone="no"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
            xmlns:xsd="http://www.w3.org/2001/XMLSchema_Gluposti" 
           xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
<soap:Body>
<@@@!!!LicenseLimit xmlns="http://pitagor-bg.com/">
<aUserId><<cUSER>></aUserId>
<aPartnerEIK><<cEIK>></aPartnerEIK>
<aDate><<lcDt>></aDate>
<aRecords><<TRANSFORM(nRecords)>></aRecords>
</@@@!!!LicenseLimit>
</soap:Body>
</soap:Envelope>
ENDTEXT

lcEnv = STRTRAN(lcEnv, "@@@!!!", cSetChk)
STRTOFILE(lcEnv, "soap.xml")
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
xh.setRequestHeader ("SOAPAction", "http://pitagor-bg.com/"+cSetChk+"LicenseLimit")
*xh.setRequestHeader ("SOAPAction", "http://pitagor-bg.com/CheckLicenseLimit")

xh.send(lcEnv)
iF (xh.readyState <> 4) OR NOT ok
    WAIT "Error at comunication with WS!" WINDOW AT 0,0 TIMEOUT 3
    this.lcResponse=.f.
    xh=.null.
    RELEASE xh
ELSE
    lcXML = xh.responseXML.xml
    STRTOFILE(lcXML, "xml.xml")
    STRTOFILE(xh.responsebody, "d:\temp\responcebody.xml")
    xh=.null.
    RELEASE xh
    RETURN lcXML
ENDIF 
retu



lcUsr = MLINE(FILETOSTR("c:\ptg\ptg-xp.lcnz"), 1)

lcXml=chkLLimit("FDML", "203231690", DATE(), 4321 ) 

STRTOFILE(lcXml, "xml.xml")
loXML = CREATEOBJECT('MSXML2.DomDocument')  
loXML.ASYNC = .f.
loXML.LOADXML(lcXml)
loRoot = loXML.documentElement

lnCount = val( loRoot.getElementsByTagName("LimitRecords").item[0].nodeTypedValue )
_dt = loRoot.getElementsByTagName("LimitDate").item[0].nodeTypedValue
_dt = TTOD( CTOT(_dt))
??_dt



******************
FUNCTION chkLLimit
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
FUNCTION tstLcnz
ok=.t.
loHTTP = CREATEOBJECT('WinHttp.WinHttpRequest.5.1')
lcID = MLINE(FILETOSTR("c:\PTG\id.ptg"), 1)
lcURL = "http://pitagor-bg.com/Service.aspx?Method=GetLicense&PCId="+lcID
loHTTP.Open("POST", lcURL, .f.)
loHttp.setrequestheader("Content-Type", "text/html")
lcTmpFile =  ADDBS(SYS(2023))+"tmp.txt"
lcErrMsg = ""

If ok   
   loHTTP.Send( " ???" )
   lcTxt = loHTTP.ResponseBody
   IF TYPE("lcTxt") # "C"
      RETURN .f.
   ENDIF 
ELSE
   RETURN .f.
EndIf


STRTOFILE( lcTxt, lcTmpFile)
SELECT 0
Create Cursor tst_ (user c(10),UsrName c(40),pcN n(5), item c(10),;
      dt1 d, opt c(20),xxx c(1), dtRls d,dtc c(10),opt2 c(10), csum c(16))
APPEND FROM (lcTmpFile ) DELIMITED WITH _ WITH CHARACTER "|"
SELECT * from tst_ WHERE item="IMPORT" INTO CURSOR tst_
IF dtRls > DATE()
   RETURN .t.
ENDIF 

MESSAGEBOX("Некоректен лиценз!", 48, "Импорт на данни:")

RETURN dtRls > DATE() - 50



RETURN


ok=.t.
loHTTP = CREATEOBJECT('WinHttp.WinHttpRequest.5.1')
lcID = MLINE(FILETOSTR("c:\PTG\id.ptg"), 1)
lcURL = "http://pitagor-bg.com/Service.aspx?Method=GetLicense&PCId="+lcID
loHTTP.Open("POST", lcURL, .f.)
loHttp.setrequestheader("Content-Type", "text/html")
lcTmpFile =  ADDBS(SYS(2023))+"tmp.txt"
lcErrMsg = ""

If ok   
   loHTTP.Send( " ???" )
   lcTxt = loHTTP.ResponseBody
   IF TYPE("lcTxt") # "C"
      RETURN .f.
   ENDIF 
ELSE
   RETURN .f.
EndIf


STRTOFILE( lcTxt, lcTmpFile)

retu


lcFN = GETFILE("txt", "Файл:","Избери",0, "Документи във формат INV.BG:")
CREATE CURSOR s_ (f1 c(1), Dt d, fan n(10), f2 c(1),st2 n(15,2), f9 c(1),;
    KName c(50), f3 c(1),f4 c(1), f5 c(1),;
    id_dds c(15), efn c(13), f6 c(1), f7 c(1), f8 c(1),;
    nVAT n(15,2))

*APPEND FROM (lcFN) DELIMITED WITH CHARACTER  "|"
APPEND FROM (lcFN) DELIMITED WITH _ with CHARACTER  "|"

SELECT " "as fas,fan,dt, efn, SPACE(10)as ktrg, KName, ID_DDS, ;
    "411" as skaw,  st2 as ttl, st2 - nVAT as StFar, St2 FROM s_ ;
  into CURSOR fa_ readw
BROWSE

SELECT SPACE(10)as ktrg,KName as name, efn, id_dds  FROM s_  ;
  WHERE NOT EMPTY(efn) ;
  GROUP BY efn ;
  INTO CURSOR ktrg_ readw
BROWSE

retu


*!*	lcFN = GETFILE("txt", "Файл:","Избери",0, "Документи във формат INV.BG:")
*!*	CREATE CURSOR s_ (f1 c(1), Dt d, fan n(10), f2 c(1),st2 n(15,2), f9 c(1),;
*!*	    KName c(50), f3 c(1),f4 c(1), f5 c(1),;
*!*	    id_dds c(15), efn c(13), f6 c(1), f7 c(1), f8 c(1),;
*!*	    nVAT n(15,2))
*!*	*STRTOFILE(STRTRAN(FILETOSTR(lcFN),["], [ ]), "_txt.txt")
*!*	*APPEND FROM (lcFN) DELIMITED WITH CHARACTER  "|"
*!*	APPEND FROM (lcFN) DELIMITED WITH _ with CHARACTER  "|"
*!*	BROWSE FOR fan>4220
*!*	RETURN

CLOSE DATABASES all
SET DEFAULT TO d:\temp\acserv-atlia\

lcFN=GETFILE("xls","Фактури")
IMPORT FROM (lcFN) xl5
lcAlias=ALIAS()

SELECT c as KName, d as efn, e as id_dds, ;
  ROUND(VAL(n),0)as fan,CTOD(o) as dt, p as art_, q as aname, u as q  ;
  FROM (lcAlias) INTO CURSOR fa_
* BROWSE
  
* _path="d:\temp\acserv-atlia\"
*!*	SELECT NVL(artkl, SPACE(10))as artkl, ;
*!*	       NVL(name, SPACE(40))as name, ;
*!*	    fa_.art_,aname from fa_  ;
*!*	  left JOIN d:\temp\acserv-atlia\artkl on VAL(fa_.art_)=VAL(artkl)  ;
*!*	  INTO CURSOR art_
  
	SELECT fa_.art_ as art, AName as name, SPACE(5)as mka, SPACE(10)as artkl ;
	  FROM fa_  GROUP BY art   ;
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
      
BROWSE
  
DO FORM d:\vfp\s2ptg\art_


SET DEFAULT TO d:\vfp\s2ptg

retu

lcFN=GETFILE("xls", "Данни за плащанията")
IMPORT FROM (lcFN) xl5
lcAlias=ALIAS()
SELECT c as fan, i as stst, s as cPmnt, t as VidDkm FROM (lcAlias) INTO CURSOR p_
SELECT * from p_ where LOWER(cPmnt)="в брой" AND LOWER(vidDkm)="фактура" INTO CURSOR p_

BROWSE
retu

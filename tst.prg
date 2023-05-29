CREATE CURSOR fa_ (cFAN c(10), Dt d, cTxt c(30), ;
   cTtl c(12), cVAT c(12),;
   efn c(14), id_dds c(15), KName c(40))

lcS3=GETFILE("xml", "*.xml","Избери",0, "Документи за продажби:")
lcXml=FILETOSTR(lcS3)

loXML = CREATEOBJECT('MSXML2.DomDocument')  
loXML.ASYNC = .f.
loXML.LOADXML(lcXml)
CLEAR
loXML.save("tst1251.xml")
SET DATE YMD 
loRoot = loXML.documentElement
?
FOR EACH loNode IN loRoot.childNodes
  APPEND BLANK IN fa_
*  REPLACE Dt WITH loNode.getAttribute("Date")
  REPLACE cTxt WITH STRCONV( loNode.getAttribute("Term"), 11)  
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

brow
RETURN




*.item[0].nodeTypedValue'
IF TYPE(lcNODE)="C"

ENDIF       













CREATE CURSOR s3_ (Operation m, Document m, Company m )

*!*	CREATE CURSOR s3_ (Date c(15), Number n(10), InvoiceDate c(15), ;
*!*	   Name c(40), Bulstat c(15) )

WAIT "Конвертиране..." WINDOW AT 0,0 nowa
=XMLTOCURSOR(lcS3, "s3_",512+8192)
WAIT clea

brow
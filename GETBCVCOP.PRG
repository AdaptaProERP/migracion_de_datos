// Programa   : GETBCVCOP
// Fecha/Hora : 25/08/2023 21:18:45
// Propósito  : Obtener el valor del Peso Colombiano Según el BCV
// https://www.bcv.org.ve/sites/default/files/EstadisticasGeneral/2_1_2b23_smc.xls el trimestre anterior, el siguiente trimestre podrá ser d
// Creado Por : Juan Navas
// Llamado por:
// Aplicación :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN(lSave,cMoneda,dFecha)
  LOCAL cUrl,cTrim:="a"
  LOCAL nMes:=MONTH(oDp:dFecha)
  LOCAL cFile:=oDp:cBin+"DBCDIVISA\DIVISAS_"+DTOS(oDp:dFecha)+".XLS"
  LOCAL oExcel,nLine:=11+2,cValue:=SPACE(10),nValorC:=0,nValorV:=0,lFind:=.F.,cCodMon:="",cSql

  lMkDir("DBCDIVISA")

  DEFAULT lSave  :=.T.,;
          cMoneda:="COP",;
          dFecha :=oDp:dFecha

  oDp:cMsgErr:=""

  // Ya fué descargado
  IF !FILE(cFile)
  
    cTrim:=IF(nMes>3 .AND. nMes<8 ,"b",cTrim)
    cTrim:=IF(nMes>7 .AND. nMes<10,"c",cTrim)
    cTrim:=IF(nMes>9              ,"d",cTrim)
 
    cTrim:=cTrim+RIGHT(STRZERO(YEAR(oDp:dFecha),4),2)
    cUrl :="https://www.bcv.org.ve/sites/default/files/EstadisticasGeneral/2_1_2"+cTrim+"_smc.xls"

    URLDownLoad(cUrl,cFile)

    IF !FILE(cFile)
      oDp:cMsgErr:="Archivo "+cFile+" no fué descargado"
      RETURN 0
      // EJECUTAR("TESTXLSTOARRAY",cFile,11,7,2) 
   ENDIF

  ENDIF

  lFind:=COUNT("DPHISMON","HMN_CODIGO"+GetWhere("=",cMoneda)+" AND HMN_FECHA"+GetWhere("=",dFecha))>	0 

// ? lFind,CLPCOPY(oDp:cSql),cCodMon,dFecha

  IF !lFind

    oExcel := TExcelScript():New()
    oExcel:Open( cFile )

    WHILE .T.

      cCodMon:=oExcel:Get( nLine , 2 ,@cValue ) 
 
      IF cCodMon="COP" .OR. Empty(cCodMon)
         EXIT
      ENDIF

      nLine++

    ENDDO

    IF !Empty(cCodMon)

      cValue :=0.00000

      SET DECI TO 8

      //  nValorC:=oExcel:oSheet:Cells( nLine, 7 ):Value
      nValorC:=oExcel:Get( nLine , 6 ,@cValue ) 
      nValorV:=oExcel:Get( nLine , 7 ,@cValue ) 

    ENDIF

  ENDIF

  //? cCodMon,nLine,nValorC,nValorV

  oExcel:End(.F.)

 

  IF !lFind

    EJECUTAR("CREATERECORD","DPHISMON",{"HMN_CODIGO","HMN_FECHA","HMN_VALOR","HMN_HORA"},; 
                                       {cCodMon     ,dFecha     , nValorC   ,"00:00:00"},;
                                        NIL,.T.,"HMN_CODIGO"+GetWhere("=",cCodMon)+" AND HMN_FECHA"+GetWhere("=",dFecha))

    cSql:=[UPDATE DPHISMON SET HMN_VALOR]+GetWhere("=",nValorC)+[ WHERE HMN_CODIGO]+GetWhere("=",cCodMon)+" AND HMN_FECHA"+GetWhere("=",dFecha)
    OpenOdbc(oDp:cDsnData):EXECUTE(cSql)

  ENDIF

 SET DECI TO 2

RETURN nValorC
// EOF

// Programa   : IMPORTCXP
// Fecha/Hora : 06/10/2003 16:18:52
// Propósito  : Importar Datos de Proveedores DP20
// Creado Por : Juan Navas
// Llamado por: IMPORTDP20
// Aplicación : Todas
// Tabla      : Todas

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cDir,oMeterT,oMeterR,oSayT,oSayR,nTables,lIniciar,lCxP)
   LOCAL cSql:=" SET FOREIGN_KEY_CHECKS = 0"
   LOCAL oDb :=OpenOdbc(oDp:cDsnData)
   LOCAL oTableAct:=NIL
   LOCAL oTablePro:=NIL
   LOCAL oTableCla:=NIL
   LOCAL oTablePer:=NIL
   LOCAL oTableCla:=NIL
   LOCAL oDoc     :=NIL
   LOCAL oTableCar:=NIL
   LOCAL oTableCta:=NIL
   LOCAL oMemo    :=NIL

   LOCAL lMeter:=(ValType(oMeterR)="O")

   DEFAULT cDir     :="C:\CRUZROJA\EMP713\",;
           lIniciar :=.T.

// oDp:lDpXbase:=.T.

   DEFAULT cDir:="C:\DP20\EMP100\",;
           nTables  :=0    ,;
           lIniciar :=.T.  ,;
           lCxP     :=.T.  ,;
           oDp:cMemo:=""


   // lCxP (Solo CxP)

   // oDp:lTracer:=.F.


   oDb:EXECUTE(cSql)

   IF lIniciar
      EJECUTAR("DELETEPRO",.T.)
   ENDIF

   CLOSE ALL

//   oDp:lTracer:=.T.
   oDp:lAuditar:=.F.
   oDp:lMYSQLCHKCONN:=.F.

   oTableAct:=OpenTable("SELECT * FROM DPACTIVIDAD_E",.F.)
   oTableAct:lAuditar:=.F.

   oTablePro:=OpenTable("SELECT * FROM DPPROVEEDOR"  ,.F.)
   oTablePro:lAuditar:=.F.

   oTablePer:=OpenTable("DPPROVEEDORPER" , .F. )
   oTablePer:lAuditar:=.F.
 
   oTableCla:=OpenTable("SELECT * FROM DPPROCLA",.F.)
   oTableCla:lAuditar:=.F.

   oDoc:=OpenTable("SELECT * FROM DPDOCPRO",.F.)
   oDoc:lAuditar:=.F.

   oMemo:=OpenTable("DPMEMO",.F.)
   oMemo:lAuditar:=.F.

   oTableCta:=OpenTable("SELECT * FROM DPCTA",.F.)
   oTableCta:lAuditar:=.F.

   IF(lMeter , oMeterT:Set(nTables++) , NIL)
   DPPROCLA()

//   IF lCxP 
//      IMPORTDPCOM()
//      RETURN 
//   ENDIF
 


   IF(lMeter , oMeterT:Set(nTables++) , NIL)
   DPPROVEEDOR(NIL,lCxP)

   CLOSE ALL

   cSql:=" SET FOREIGN_KEY_CHECKS = 1"
   oDb:EXECUTE(cSql)

   oTableAct:End()
   oTableCla:End()
   oTableCar:End()

   oTablePro:End()
   oTablePer:End()
   oDoc:End()

   oDp:lAuditar:=.T.


// ? "CONCLUIDO"

RETURN nTables

PROCE DPPROVEEDOR(cCodpro,lCxP)
  LOCAL cFile:=cDir+"DPPRO.DBF",nContar:=0,oTable,cPrecio:="",cCodCla:="",cCodAct:=""
  LOCAL cRepres:="Representante Legal",nSysR:=0,nRecno:=0,I,uValue,nPos,nLen
  LOCAL cCodCta,cWhere,cPROCODIGO:="",cRif:="",cFileDoc:=cDir+"DPCOM.DBF"

  DEFAULT cCodpro:=""

  BUILDCARGO(cRepres)

  // oTable :=oTablePro // OpenTable("SELECT * FROM DPPROVEEDOR ", .F. )
  nLen   :=LEN(oTablePro:PRO_CODIGO)
  // oTable:End()

  // 08/01/2024 oTableP:=OpenTable("DPPROVEEDORPER" , .F. )

  CLOSE ALL

  IF lCxP
    SELE F
    USE (cFileDoc) VIA "DBFCDX" SHARED ALIAS "FACTURA"
  ENDIF

  SELE A
  USE (cFile) VIA "DBFCDX" SHARED NEW ALIAS "PROVEEDOR"

// ? FILE(cFile),cFile
// browse()

  SET FILTER TO !DELETED()
  SET ORDE TO 0

  IF lMeter
    oMeterR:SetTotal(RecCount())
    oSayT:SetText("Actividad Económica")
  ELSE
    oDp:oFrameDp:SetText("Actividad Económica")
  ENDIF

  // REVISAR 

  SELE A
  GO TOP

  WHILE !A->(EOF())

     nContar++

     SysRefresh(.T.)

     IF lMeter

      oMeterR:Set(nContar)
      oSayR:SetText(LSTR(nContar)+"/"+LSTR(RECCOUNT()))

    ELSE

      oDp:oFrameDp:SetText(LSTR(nContar)+"/"+LSTR(RECCOUNT())+" "+A->CLI_CODIGO)

    ENDIF

    // Crear Actividad Económica
    BUILDACTIVI(A->CLI_ACTIVI)

    A->(DBSKIP())

  ENDDO

  GO TOP

  IF lMeter
    oMeterR:SetTotal(RecCount())
    oSayT:SetText("Proveedor")
  ELSE
    oDp:oFrameDp:SetText("PROVEEDOR ")
  ENDIF

  nContar:=0

  SELE A
  GO TOP

  WHILE !A->(EOF())

     nContar++

     IF lMeter

        oMeterR:Set(nContar)
        oSayR:SetText(LSTR(nContar)+"/"+LSTR(RECCOUNT())+" "+A->CLI_CODIGO)

     ELSE

       oDp:oFrameDp:SetText("PROVEEDOR "+LSTR(nContar)+"/"+LSTR(RECCOUNT())+" "+A->CLI_CODIGO)

       IF nSysR++>10
         nSysR:=0
         SysRefresh(.T.)
       ENDIF

     ENDIF

     cCodPro:=ALLTRIM(A->CLI_CODIGO) 

     // 26/12/2022 Solo proveedores Activos con CxP.
     IF lCxP .AND. .F.

        SELECT F
        // SET FILTER TO FAC_CODCLI=cCodPro .AND. FAC_ESTADO="P"
        SET FILTER TO FAC_CODCLI=cCodPro .AND. FAC_ESTADO="P" 
//.AND. (FAC_TIPO="FC" .OR. FAC_TIPO="AN")

        GO TOP

        IF EMPTY(F->FAC_CODCLI)
           SELECT A
           A->(DBSKIP())
           LOOP
        ENDIF

        SELECT A

     ENDIF

     cRif   :=STRTRAN(A->CLI_RIF,"-","")

/*
     IF !Empty(A->CLI_RIF) .AND. ISSQLGET("DPPROVEEDOR","PRO_RIF",cRif)
       cWhere :="PRO_RIF"   +GetWhere("=",cRif)
       cCodPro:=SQLGET("DPPROVEEDOR","CLI_CODIGO","PRO_RIF"+GetWhere("=",cRif))
     ELSE
       cWhere :="PRO_CODIGO"+GetWhere("=",cCodPro   )
     ENDIF
*/

     cWhere :="PRO_CODIGO"+GetWhere("=",cCodPro   )
/*
     oTable :=OpenTable("SELECT * FROM DPPROVEEDOR WHERE "+cWhere, .T. )

     IF oTable:RecCount()=0
       cWhere:=""
       oTable:Append()
     ENDIF
*/   

     oTable:=oTablePro // 08/01/2024

     oTablePro:Append()
     AEVAL(DBSTRUCT(),{|a,n,uValue,cField|uValue:=FieldGet(n),;
                                          cField:="PRO"+SUBS(FieldName(n),4,10),;
                                          uValue:=IIF(ValType(uValue)="C",OEMTOANSI(uValue),uValue),;
                                          IIF(oTablePro:FieldPos(cField)>0,;
                                          oTablePro:Replace(cField,uValue),NIL)})

     cCodpro:=ALLTRIM(A->CLI_CODIGO) 

     IF ALLDIGIT(cCodpro)
        cCodPro:=STRZERO(VAL(cCodPro) ,nLen )
     ENDIF

/*
     IF ISSQLGET("DPPROVEEDOR","CLI_CODIGO",cCodpro)

        oDp:cMemo:=oDp:cMemo+;
                   IIF( Empty(oDp:cMemo) , "" , CRLF )+;
                   "Código de Proveedor "+A->CLI_CODIGO+" ya Existe"

        A->(DBSKIP())
        LOOP

     ENDIF
*/

     cCodCla:=BUILDCLAPRO(A->CLI_PROCLA)
     cCodAct:=BUILDACTIVI(A->CLI_ACTIVI)
     cCodCta:=BUILDCODCTA(A->CLI_CUENTA)

     oTablePro:Replace("PRO_CODIGO" ,cCodPro)
     oTablePro:Replace("PRO_RIF"    ,cRif   )
     oTablePro:Replace("PRO_CODCLA" ,cCodCla)
     oTablePro:Replace("PRO_CUENTA" ,cCodCta)
     oTablePro:Replace("PRO_ACTIVI" ,cCodAct)
     oTablePro:Replace("PRO_PAIS"   ,oDp:cPais )
     oTablePro:Replace("PRO_ESTADO" ,oDp:cEstado)
     oTablePro:Replace("PRO_MUNICI" ,oDp:cMunicipio)
     oTablePro:Replace("PRO_PARROQ" ,oDp:cParroquia)

     oTablePro:Commit("") // cWhere)
     // oTable:End()

      IF !Empty(CLI_REPRES) 
        oTablePer:Append()
        oTablePer:Replace("PDP_CODIGO",cCodpro)
        oTablePer:Replace("PDP_CARGO" ,cRepres)
        oTablePer:Replace("PDP_PERSON",OEMTOANSI(CLI_REPRES))
        oTablePer:Commit("")
     ENDIF

     IF lCxP
        IMPORTDPCOM(cCodPro)
     ENDIF

     // Ahora crearemos el Cargo Representante

     A->(DBSKIP())

  ENDDO

  oTable:End()
//  oTablePer:End()

  IIF( lMeter , oMeterR:Set(RecCount()) , NIL )

  IF Empty(cCodPro)
    CLOSE ALL
  ENDIF
    
RETURN NIL


/*
// Obtiene el Grupo
*/
FUNCTION BUILDCLAPRO(cCodCla)
  LOCAL oTable

  IF Empty(cCodCla)
    cCodCla:=SQLINCREMENTAL("DPPROCLA","CLP_CODIGO")
  ENDIF

  cCodCla:=ALLTRIM(cCodCla)

  IF ISSQLGET("DPPROCLA","CLP_CODIGO",cCodCla)
     RETURN cCodCla
  ENDIF
    
  ///oTable:=OpenTable("DPPROCLA",.F.)
  oTableCla:AppendBlank()
  oTableCla:Replace("CLP_CODIGO",cCodCla)
  oTableCla:Replace("CLP_DESCRI",cCodCla)
  oTableCla:Commit()

  //oTable:End()

RETURN cCodCla


/*
// Genera Cargos
*/
FUNCTION BUILDCARGO(cNombre)
  LOCAL oTable

  IF COUNT("DPCARGOS","CAR_CODIGO"+GETWHERE("=",cNombre))>0
      RETURN cNombre
  ENDIF

  // oTable:=OpenTable("SELECT * FROM DPCARGOS",.F.)
  oTableCar:AppendBlank()
  oTableCar:Replace("CAR_CODIGO",cNombre)
  oTableCar:Commit()

RETURN cNombre

/*
// Genera Actividad Econ¢mica
*/
FUNCTION BUILDACTIVI(cNombre)
  LOCAL cCodigo,oTable

  IF !Empty(cNombre)

     cCodigo:=SQLGET("DPACTIVIDAD_E","ACT_CODIGO","ACT_DESCRI"+GETWHERE("=",cNombre))

     IF !EMPTY(cCodigo)
       RETURN cCodigo
     ENDIF

  ENDIF

  cCodigo:=SQLINCREMENTAL("DPACTIVIDAD_E","ACT_CODIGO")

//  oTable:=OpenTable("SELECT * FROM DPACTIVIDAD_E",.F.)
  oTableAct:AppendBlank()
  oTableAct:Replace("ACT_CODIGO",cCodigo)
  oTableAct:Replace("ACT_DESCRI",cNombre)
  oTableAct:Replace("ACT_MEMO"  ,"Desde DP20")
  oTableAct:Commit()
//  oTable:End()

RETURN cCodigo

//
// DPPROCLA
//
FUNCTION DPPROCLA()
   LOCAL cFile:=cDir+"DPPROCLA.DBF",cCodCla

   LOCAL oTable

   CLOSE ALL

   oTable:=OpenTable("DPPROCLA",.F.)
   oTable:lAuditar:=.F.

   SELE A
   USE (ALLTRIM(cFile)) VIA "DBFCDX" SHARED NEW 

   DBGOTOP()

   WHILE !A->(EOF())

     cCodCla:=A->PRO_CLACOD

     IF !ISSQLGET("DPPROCLA","CLP_CODIGO",cCodCla)
        oTable:AppendBlank()
        oTable:Replace("CLP_CODIGO",cCodCla)
        oTable:Replace("CLP_DESCRI",ANSITOOEM(DPPROCLA->PRO_CLANOM))
        oTable:Commit()
     ENDIF

     A->(DBSKIP())

   ENDDO

   USE

   oTable:End()

RETURN .T.

FUNCTION BUILDMEMO(nNumMem,cDescri)
  LOCAL cAlias:=ALIAS(),cFile,cIndex // ,oMemo

  DEFAULT cDescri:=""

  IF nNumMem=0
     RETURN 0
  ENDIF

  IF !DPSELECT("DPMEMO")
     cFile:=ALLTRIM(cDir)+"DPMEMO"
     cIndex:=cFile+".CDX"
     SELE G
     USE (cFile) SHARED VIA "DBFCDX" NEW
     SET INDEX TO (cIndex)
  ENDIF

  SET ORDE TO 1
  IF !DBSEEK(nNumMem)
     GO TOP
     LOCATE FOR MEM_NUMERO=nNumMem
  ENDIF

  nNumMem:=SQLINCREMENTAL("DPMEMO","MEM_NUMERO")
  // oMemo:=OpenTable("DPMEMO",.F.)
  oMemo:REPLACE("MEM_NUMERO",nNumMem)
  oMemo:REPLACE("MEM_MEMO"  ,OEMTOANSI(MEM_MEMO ))
  oMemo:REPLACE("MEM_DESCRI",IIF(EMPTY(MEM_DESCRI),OEMTOANSI(cDescri),MEM_DESCRI))
  oMemo:Commit()
  // oMemo:End()

  DPSELECT(cAlias)

RETURN nNumMem

FUNCTION PROREPITE()
  LOCAL nRecno:=RECNO(),cCodpro
  LOCAL aData:={}

  AEVAL( Dbstruct(), {|a,n| AADD(aData,FieldGet(n)) } )

  DBGOBOTTOM()
  cCodpro:=STRZERO(VAL(CLI_CODIGO)+1,8)
  APPEND BLANK
  BLOC()
  AEVAL( aData ,{|a,n| FieldPut(n,a) } )

  REPLACE CLI_CODIGO WITH cCodpro
 
  DBGOTO(nRecno)

RETURN cCodpro


/*
// Genera Cuenta Contable
*/
FUNCTION BUILDCODCTA(cCodCta)
  LOCAL oTable

  IF Empty(cCodCta)
     cCodCta:=oDp:cCtaIndef
  ENDIF

  // ya Existe
  IF !EMPTY(SQLGET("DPCTA","CTA_CODIGO","CTA_CODIGO"+GetWhere("=",cCodCta)))
      RETURN cCodCta
  ENDIF

  // oTable:=OpenTable("SELECT * FROM DPCTA",.F.)
  oTableCta:AppendBlank()
  oTableCta:Replace("CTA_CODIGO",cCodCta)
  oTableCta:Replace("CTA_DESCRI","Importado desde DP20 Proveedores")
  oTableCta:Replace("CTA_CODIGO",cCodCta)
  oTableCta:Replace("CTA_CODMOD",oDp:cCtaMod)
  oTableCta:Commit()
  // oTable:End()

RETURN cCodCta

FUNCION IMPORTDPCOM(cCodPro)
   LOCAL cTipDoc:="",cWhere:="",nCxP  

   DEFAULT lCxP:=.T.

   SELECT F
   SET FILTER TO FAC_CODCLI=cCodPro .AND. FAC_ESTADO="P" 

//.AND. (FAC_TIPO="FC" .OR. FAC_TIPO="AN")

   GO TOP
//   BROWSE()

   oDoc:=OpenTable("SELECT * FROM DPDOCPRO",.F.)

   WHILE !F->(EOF())

      cTipDoc:=""
      cTipDoc:=IF(FAC_TIPO="FC","FAC",cTipDoc)
      cTipDoc:=IF(FAC_TIPO="AN","ANT",cTipDoc)
      cTipDoc:=IF(FAC_TIPO="CR","CRE",cTipDoc)
      cTipDoc:=IF(FAC_TIPO="DB","DEB",cTipDoc)
      cTipDoc:=IF(FAC_TIPO="GR","GIR",cTipDoc)

      IF Empty(cTipDoc)
         MensajeErr(F->FAC_TIPO,"Requiere Equivalente para DOC_TIPDOC")
         F->(DBSKIP())
         LOOP
      ENDIF

      nCxP   :=SQLGET("DPTIPDOCPRO","TDC_CXP","TDC_TIPO"+GetWhere("=",cTipDoc))
      nCxP   :=IIF(nCxP="D", 1,nCxP)
      nCxP   :=IIF(nCxP="C",-1,nCxP)
      nCxP   :=IIF(nCxP="N",0 ,nCxP)
     
/*
      cWhere :="DOC_CODSUC"+GetWhere("=",oDp:cSucursal)+" AND "+;
               "DOC_TIPDOC"+GetWhere("=",cTipDoc      )+" AND "+;
               "DOC_NUMERO"+GetWhere("=",F->FAC_NUMERO)+" AND "+;
               "DOC_TIPTRA"+GetWhere("=","D")

      oDoc:=OpenTable("SELECT * FROM DPDOCPRO WHERE "+cWhere,.T.)
*/
      IF oDoc:RecCount()=0 .OR. .T.
         oDoc:AppendBlank()
         cWhere:=""
      ENDIF

      oDoc:Replace("DOC_TIPDOC",cTipDoc)
      oDoc:Replace("DOC_CODSUC",oDp:cSucursal     )
      oDoc:Replace("DOC_NUMERO",F->FAC_NUMERO     )
      oDoc:Replace("DOC_TIPTRA","D"               )
      oDoc:Replace("DOC_CODIGO",cCodPro)
//      oDoc:Replace("DOC_CONDIC",F->FAC_CONDIC     )
      oDoc:Replace("DOC_FECHA" ,F->FAC_FECHA      )
      oDoc:Replace("DOC_FCHVEN",F->FAC_FECHAV     )
      oDoc:Replace("DOC_NETO"  ,ABS(F->FAC_NETO)-ABS(F->FAC_PAGADO))
      oDoc:Replace("DOC_ESTADO","AC"              )
      oDoc:Replace("DOC_DOCORG","C"               )
      oDoc:Replace("DOC_ORIGEN","I"               )  // Importación
      oDoc:Replace("DOC_HORA"  , TIME()           )
      oDoc:Replace("DOC_CXP"   , nCxP             )
      oDoc:Replace("DOC_USUARI", oDp:cUsuario     )
      oDoc:Replace("DOC_CENCOS", oDp:cCenCos      )
      oDoc:Replace("DOC_CODMON", oDp:cMoneda      )
      oDoc:Replace("DOC_FACAFE", ""               ) // oPlaGraOrC:oFrmMain:cNumero)
      oDoc:Replace("DOC_ACT"   , 1                )

      oDoc:Commit(cWhere)


      F->(DBSKIP())

   ENDDO

   SELECT A

RETURN 

// EOF


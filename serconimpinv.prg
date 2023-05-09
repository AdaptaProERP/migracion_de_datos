// Programa   : SERCOMIMPINV
// Fecha/Hora : 09/05/2023 09:04:58
// Propósito  : Importar Archivo PRODUCTOS.CSV Obtenido desde PRODUCTOS.XLSX
// Creado Por : Juan Navas
// Llamado por: Desde Programación
// Aplicación : Programación
// Tabla      : DPINV,DPGRU,DPPRECIOS
// Descargar  : https://mega.nz/file/5UV3gLDI#1Iu0hHTneHS6ke5OBQbbmJDzr2vJeiIYtmGhwunlrl8

#INCLUDE "DPXBASE.CH"

PROCE MAIN()
  LOCAL cFile:="PRODUCTOS.CSV"
  LOCAL cMemo:=MemoRead(cFile),aData,I,aLine,cWhere,cGrupo

  IF !FILE(cFile)
     MsgMemo("Archivo "+cFile+" no Encontrado")
     RETURN .F.
  ENDIF

  cMemo:=STRTRAN(cMemo,CRLF,CHR(10)) // Reemplazar CHR(13)+CHR(10)
  aData:=_VECTOR(cMemo,CHR(10))      // Crear Arreglo desde Memo separados por CHR(10)

  AEVAL(aData,{|a,n| aData[n]:=_VECTOR(a,";")}) // Convierte cada Linea en Multilinea separados por ;

  aData:=ADEPURA(aData,{|a,n| LEN(a)<>11})            // Remueve las lineas con Columnas diferentes a 11
  aData:=ADEPURA(aData,{|a,n| Empty(a[1]) })          // Remueve lineas Vacias
  aData:=ADEPURA(aData,{|a,n| "GReport 1.1.3"$a[1] }) // Remueve textos innecesarios
  aData:=ADEPURA(aData,{|a,n| "CODIGO"$a[1] })        // Remueve sin Encabezados
  aData:=ADEPURA(aData,{|a,n| "NOTA:"$a[1] })         // Remueve Texto innecesario

   FOR I=1 TO LEN(aData)
      aLine:=aData[I]

      IF Empty(aLine[3])
         cGrupo:=aLine[1]
         cGrupo:=BUILDGRUPO(cGrupo)
      ENDIF

      aLine[9]:=STRTRAN(aLine[9],"MESES","") // Remueve Texto Innecesario
      aLine[9]:=STRTRAN(aLine[9],"N/A"  ,"") // Sin Garantias
      aLine[9]:=STRTRAN(aLine[9],"MES"  ,"") // Remueve Texto Innecesario


      EJECUTAR("CREATERECORD","DPINV",{"INV_CODIGO","INV_DESCRI","INV_ESTADO","INV_GRUPO","INV_IVA","INV_OBS1","INV_MESGAR"  },; 
                                      {aLine[1]    ,aLine[3]    ,"A"         ,cGrupo      ,"GN"    ,aLine[3]  ,aLine[9]       },;
                                      NIL,.T.,"INV_CODIGO"+GetWhere("=",aLine[1]))
     
      aLine[7]:=STRTRAN(aLine[7],",",".")

      EJECUTAR("DPINVCREAUND",aLine[1],"UND") 
      EJECUTAR("DPPRECIOSCREAR",aLine[1],"A","UND","DBC",VAL(aLine[7]))

   NEXT I


RETURN

/*
// Obtiene el Grupo
*/
FUNCTION BUILDGRUPO(cGrupo)
  LOCAL oTable,cCodigo

  IF Empty(cGrupo)
     cGrupo:=STRZERO(0,6)
  ENDIF

  cCodigo:=SQLGET("DPGRU","GRU_CODIGO","GRU_DESCRI"+GetWhere("=",cGrupo))
  
  IF !Empty(cCodigo)
      RETURN cCodigo
  ENDIF

  cCodigo:=SQLINCREMENTAL("DPGRU","GRU_CODIGO")
  
  oTable:=OpenTable("SELECT * FROM DPGRU",.F.)
  oTable:Append()
  oTable:Replace("GRU_CODIGO",cCodigo)
  oTable:Replace("GRU_DESCRI",cGrupo )
  oTable:Replace("GRU_ACTIVO",.T.    )
  oTable:Commit(NIL,.F.)
  oTable:End()

RETURN cCodigo
// EOF



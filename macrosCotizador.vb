 Sub CopiarRango()
 
ImportarEstimates:
    ' Pregunta si desea insertar Estimates
    AskEst = "Deseas Importar Estimates?"
    ImportEstimate = MsgBox(AskEst, vbQuestion + vbYesNo)
    If ImportEstimate = vbNo Then
        'Codigo si la respuesta es NO
        GoTo ImportarInfra
    Else
        'Codigo si la respuesta es SI
        ImpEstimates
    End If
  
ImportarInfra:
    ' Pregunta si desea insertar Infra
    AskInf = "Deseas Importar Archivo de INFRA?"
    ImportInfra = MsgBox(AskInf, vbQuestion + vbYesNo)
    If ImportInfra = vbNo Then
        'Codigo si la respuesta es NO
        GoTo ImportarTerceros
    Else
        'Codigo si la respuesta es SI
        ImpInfra
    End If
 
 
ImportarTerceros:
    ' Pregunta si desea insertar Terceros
    Askterc = "Deseas Importar Archivo de Terceros (Partners de Ecosistema)?"
    Importterc = MsgBox(Askterc, vbQuestion + vbYesNo)
    If Importterc = vbNo Then
        'Codigo si la respuesta es NO
        GoTo ImportarCSP
    Else
        'Codigo si la respuesta es SI
        ImpTerceros
    End If
 
ImportarCSP:
    ' Pregunta si desea insertar CSP
    AskCSP = "Deseas Importar Archivo CSP?"
    ImportCSP = MsgBox(AskCSP, vbQuestion + vbYesNo)
    If ImportCSP = vbNo Then
        'Codigo si la respuesta es NO
        GoTo ContinuarCot
    Else
        'Codigo si la respuesta es SI
        CSP
    End If
 
ContinuarCot:
 
    ConsolidaTemp
 
    LLENAQUOTE
    
'InsertSubtot:
    ' Pregunta si desea insertar Subtotales
    'AskST = "Deseas Insertar SUBTOTALES ?"
    'InsST = MsgBox(AskST, vbQuestion + vbYesNo)
    'If InsST = vbNo Then
        'Codigo si la respuesta es NO
    '    GoTo ContinuarCot2
    'Else
        'Codigo si la respuesta es SI
        InsertaSubtotales
    'End If
 
ContinuarCot2:
    Range("A20").Activate
    
End Sub

Sub ImpInfra()

  
   ' Obtener elNombre del Archivo Actual
   CotFile = ActiveWorkbook.Name
   
   'Crea Hoja Temporal para INFRA
   Worksheets.Add(After:=Worksheets("P&L")).Name = "INFRATEMP"
   Worksheets("INFRATEMP").Activate
   Range("A1").Value = "Partida"
   Range("B1").Value = "No. Parte"
   Range("C1").Value = "Marca"
   Range("D1").Value = "Descripcion"
   Range("E1").Value = "Cantidad"
   Range("F1").Value = "Unidad"
   Range("G1").Value = "Proveedor"
   Range("H1").Value = "Precio Lista"
   
IniciaINFRA:
    ' Indicar el libro origen desde donde copiar
    InfraCotFile = Application.InputBox("Introduzca El Nombre del Archivo Origen de INFRA", "INFRA COT FILE Name")

    ' Indicar el libro origen desde donde copiar
    Workbooks(InfraCotFile & ".xlsx").Activate
    On Error GoTo NoINFRAFile

    'Dim wksht As Worksheet
    'For Each wksht In Workbooks(InfraCotFile & ".xlsx").Worksheets
    
    Dim Contador As Integer
    Numhojas = Workbooks(InfraCotFile & ".xlsx").Worksheets.Count
    Contador = 0
    
    For Contador = 1 To Numhojas
    Workbooks(InfraCotFile & ".xlsx").Activate
    Sheets(Contador).Activate
    
    
    If ActiveSheet.Name <> "Resumen" Then
    
        ' Indicar la hoja de origen desde donde copiar
        ' Worksheets(wkshtname).Activate

        ' Seleccionar el rango fila/Columna con datos. Desde celda A15
        'lastCol = ActiveSheet.Range("A15").End(xlToRight).Column
        lastRow = ActiveSheet.Cells(65536, "G").End(xlUp).Row
        'ActiveSheet.Range("A14:" & ActiveSheet.Cells(lastRow, lastCol).Address).Select
        ActiveSheet.Range("A14:H" & lastRow).Select


        ' Copiamos el rango
        Selection.Copy

        ' Indicar el libro destino donde pegar
        Workbooks(CotFile).Activate

        ' Indicar la hoja de destino donde pegar
        Worksheets("INFRATEMP").Activate

        ' Indicar la celda donde pegar los datos
        lastCol = ActiveSheet.Range("A1").End(xlToRight).Column
        lastRow = ActiveSheet.Cells(65536, lastCol).End(xlUp).Row
        Range("A" & lastRow + 1).Select

        ' Para pegar solo como valor
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        ' Para pegar el formato (si se desea)
        ' Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        ' Limpia Portapapeles
        Application.CutCopyMode = False

        ' Al finalizar el pegado se coloca en celda A2 por si hay muchas filas para que no quede abajo
        Range("A2").Select
    
    End If
    Workbooks(InfraCotFile & ".xlsx").Activate
    Next Contador

    ' SI queremos que nos confirme la importaciãn
    'MsgBox "Los datos han sido copiados correctamente."
    
    Workbooks(CotFile).Activate
    Worksheets("INFRATEMP").Activate
        Range("M1").Value = "Partida"
        Range("N1").Value = "No. Parte"
        Range("P1").Value = "Marca"
        Range("O1").Value = "Descripcion"
        Range("R1").Value = "Cantidad"
        Range("Q1").Value = "Unidad de Medida"
        Range("S1").Value = "Proveedor"
        Range("T1").Value = "Precio Lista"
        Range("U1").Value = "Duracion del Servicio"
        Range("V1").Value = "Offer Type"

        
        If ExisteHojaEst Then
            'MsgBox "Ya exist la hoja."
            UltimaPartida = Worksheets("ESTTEMP").Cells(65536, "P").End(xlUp).Row
            Range("K1").FormulaLocal = Int(Worksheets("ESTTEMP").Range("P" & UltimaPartida))
        Else
            Range("K1") = 0
        End If
        
    With ActiveSheet.Range("A2").CurrentRegion
        UltimaFilaI = .Rows(.Rows.Count).Row
        ActiveSheet.Range("J2:J" & UltimaFilaI).FormulaLocal = "=SI(ESNUMERO(A2),TEXTO(A2/100,"".#0""),1)"
        ActiveSheet.Range("K2:K" & UltimaFilaI).FormulaLocal = "=SI(J2=1,K1+1,K1)"
        ActiveSheet.Range("L2:L" & UltimaFilaI).FormulaLocal = "=SI(J2=1,"".00"",SI(J1=1,"".01"",TEXTO(L1+0.01,"".#0"")))"
        ActiveSheet.Range("M2:M" & UltimaFilaI).FormulaLocal = "=CONCATENAR(K2,L2)"
        ActiveSheet.Range("N2:N" & UltimaFilaI).FormulaLocal = "=SI(ESNUMERO(A2),B2,"""")"
        ActiveSheet.Range("O2:O" & UltimaFilaI).FormulaLocal = "=SI(ESNUMERO(A2),D2,A2)"
        ActiveSheet.Range("P2:P" & UltimaFilaI).FormulaLocal = "=SI(ESNUMERO(A2),MAYUSC(C2),"""")"
        ActiveSheet.Range("Q2:Q" & UltimaFilaI).FormulaLocal = "=SI(ESNUMERO(A2),MAYUSC(F2),"""")"
        ActiveSheet.Range("R2:R" & UltimaFilaI).FormulaLocal = "=SI(ESNUMERO(A2),E2,1)"
        ActiveSheet.Range("S2:S" & UltimaFilaI).FormulaLocal = "=SI(ESNUMERO(A2),MAYUSC(G2),"""")"
        ActiveSheet.Range("T2:T" & UltimaFilaI).FormulaLocal = "=SI(ESNUMERO(A2),H2,0)"
        ActiveSheet.Range("U2:U" & UltimaFilaI).Value = "N/A"
        ActiveSheet.Range("V2:V" & UltimaFilaI).Value = "OS"
    
    End With
    GoTo FINALINFRA
    
NoINFRAFile:
    MsgBox ("Archivo de INFRA No Encontrado")
    GoTo IniciaINFRA
    
FINALINFRA:

End Sub
Sub ImpTerceros()

  
   ' Obtener elNombre del Archivo Actual
   CotFile = ActiveWorkbook.Name
   
   'Crea Hoja Temporal para TERCEROS
   Worksheets.Add(After:=Worksheets("P&L")).Name = "TERCEROS"
   Worksheets("TERCEROS").Activate
   Range("A1").Value = "Partida"
   Range("B1").Value = "No. Parte"
   Range("C1").Value = "Marca"
   Range("D1").Value = "Descripcion"
   Range("E1").Value = "Cantidad"
   Range("F1").Value = "Unidad"
   Range("G1").Value = "Proveedor"
   Range("H1").Value = "Precio Lista"
   
IniciaTERCEROS:
    ' Indicar el libro origen desde donde copiar
    tercerosCotFile = Application.InputBox("Introduzca El Nombre del Archivo de TERCEROS (Ecosystem Partners)", "TERCEROS COT FILE Name")

    ' Indicar el libro origen desde donde copiar
    Workbooks(tercerosCotFile & ".xlsx").Activate
    On Error GoTo NoTERCEROSFile

    'Dim wksht As Worksheet
    'For Each wksht In Workbooks(InfraCotFile & ".xlsx").Worksheets
    
    Dim Contador As Integer
    Numhojas = Workbooks(tercerosCotFile & ".xlsx").Worksheets.Count
    Contador = 0
    
    For Contador = 1 To Numhojas
    Workbooks(tercerosCotFile & ".xlsx").Activate
    Sheets(Contador).Activate
    
    
    If ActiveSheet.Name <> "Resumen" Then
    
        ' Indicar la hoja de origen desde donde copiar
        ' Worksheets(wkshtname).Activate

        ' Seleccionar el rango fila/Columna con datos. Desde celda A15
        'lastCol = ActiveSheet.Range("A15").End(xlToRight).Column
        lastRow = ActiveSheet.Cells(65536, "G").End(xlUp).Row
        'ActiveSheet.Range("A14:" & ActiveSheet.Cells(lastRow, lastCol).Address).Select
        ActiveSheet.Range("A14:H" & lastRow).Select


        ' Copiamos el rango
        Selection.Copy

        ' Indicar el libro destino donde pegar
        Workbooks(CotFile).Activate

        ' Indicar la hoja de destino donde pegar
        Worksheets("TERCEROS").Activate

        ' Indicar la celda donde pegar los datos
        lastCol = ActiveSheet.Range("A1").End(xlToRight).Column
        lastRow = ActiveSheet.Cells(65536, lastCol).End(xlUp).Row
        Range("A" & lastRow + 1).Select

        ' Para pegar solo como valor
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        ' Para pegar el formato (si se desea)
        ' Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        ' Limpia Portapapeles
        Application.CutCopyMode = False

        ' Al finalizar el pegado se coloca en celda A2 por si hay muchas filas para que no quede abajo
        Range("A2").Select
    
    End If
    Workbooks(tercerosCotFile & ".xlsx").Activate
    Next Contador

    ' SI queremos que nos confirme la importaciãn
    'MsgBox "Los datos han sido copiados correctamente."
    
    Workbooks(CotFile).Activate
    Worksheets("TERCEROS").Activate
        Range("M1").Value = "Partida"
        Range("N1").Value = "No. Parte"
        Range("P1").Value = "Marca"
        Range("O1").Value = "Descripcion"
        Range("R1").Value = "Cantidad"
        Range("Q1").Value = "Unidad de Medida"
        Range("S1").Value = "Proveedor"
        Range("T1").Value = "Precio Lista"
        Range("U1").Value = "Duracion del Servicio"
        Range("V1").Value = "Offer Type"

        
        If ExisteHojaInf Then
            'MsgBox "Ya existe la hoja."
            UltimaPartida = Worksheets("INFRATEMP").Cells(65536, "P").End(xlUp).Row
            Range("K1").FormulaLocal = Int(Worksheets("INFRATEMP").Range("M" & UltimaPartida))
        Else
            If ExisteHojaEst Then
                'MsgBox "Ya existe la hoja."
                UltimaPartida = Worksheets("ESTTEMP").Cells(65536, "P").End(xlUp).Row
                Range("K1").FormulaLocal = Int(Worksheets("ESTTEMP").Range("P" & UltimaPartida))
            Else
                Range("K1") = 0
            End If
        End If
  'OJO
  'OJO
  'OJO
  'OJO
  'OJO
  'OJO
  'OJO
  'OJO
  'OJO
  
    With ActiveSheet.Range("A2").CurrentRegion
        UltimaFilaI = .Rows(.Rows.Count).Row
        ActiveSheet.Range("J2:J" & UltimaFilaI).FormulaLocal = "=SI(ESNUMERO(A2),TEXTO(A2/100,"".#0""),1)"
        ActiveSheet.Range("K2:K" & UltimaFilaI).FormulaLocal = "=SI(J2=1,K1+1,K1)"
        ActiveSheet.Range("L2:L" & UltimaFilaI).FormulaLocal = "=SI(J2=1,"".00"",SI(J1=1,"".01"",TEXTO(L1+0.01,"".#0"")))"
        ActiveSheet.Range("M2:M" & UltimaFilaI).FormulaLocal = "=CONCATENAR(K2,L2)"
        ActiveSheet.Range("N2:N" & UltimaFilaI).FormulaLocal = "=SI(ESNUMERO(A2),B2,"""")"
        ActiveSheet.Range("O2:O" & UltimaFilaI).FormulaLocal = "=SI(ESNUMERO(A2),D2,A2)"
        ActiveSheet.Range("P2:P" & UltimaFilaI).FormulaLocal = "=SI(ESNUMERO(A2),MAYUSC(C2),"""")"
        ActiveSheet.Range("Q2:Q" & UltimaFilaI).FormulaLocal = "=SI(ESNUMERO(A2),MAYUSC(F2),"""")"
        ActiveSheet.Range("R2:R" & UltimaFilaI).FormulaLocal = "=SI(ESNUMERO(A2),E2,1)"
        ActiveSheet.Range("S2:S" & UltimaFilaI).FormulaLocal = "=SI(ESNUMERO(A2),MAYUSC(G2),"""")"
        ActiveSheet.Range("T2:T" & UltimaFilaI).FormulaLocal = "=SI(ESNUMERO(A2),H2,0)"
        ActiveSheet.Range("U2:U" & UltimaFilaI).Value = "N/A"
        ActiveSheet.Range("V2:V" & UltimaFilaI).Value = "OS"
    
    End With
    GoTo FINALTERCEROS
    
NoTERCEROSFile:
    MsgBox ("Archivo de TERCEROS No Encontrado")
    GoTo IniciaTERCEROS
    
FINALTERCEROS:

End Sub



Sub ImpEstimates()
    ' Obtener elNombre del Archivo Actual
    CotFile = ActiveWorkbook.Name
    
       'Crea Hoja Temporal para ESTIMATES
   Worksheets.Add(After:=Worksheets("P&L")).Name = "ESTTEMP"
   Worksheets("ESTTEMP").Activate
   Range("A1").Value = "Partida"
   Range("B1").Value = "No. Parte"
   Range("C1").Value = "Descripcion"
   Range("D1").Value = "P.L."
   Range("E1").Value = "Cantidad"
   Range("F1").Value = "P.L. Extendido"
   Range("G1").Value = "Descuento"
   Range("H1").Value = "Precio Venta"
   Range("I1").Value = "Duracion del Servicio"
   Range("J1").Value = "Categoria"
   
IniciaEstimates:
    ' Indicar el libro origen desde donde copiar
    EstimateName = Application.InputBox("Introduzca El Nombre del Estimate Origen", "Estimate Name")

    ' Indicar el libro origen desde donde copiar
    Workbooks(EstimateName & ".xls").Activate
    On Error GoTo NoEstFile

    ' Indicar la hoja de origen desde donde copiar
    Sheets(1).Activate

    ' Seleccionar el rango fila/Columna con datos. Desde celda B26
    lastCol = ActiveSheet.Range("B26").End(xlToRight).Column
    lastRow = ActiveSheet.Cells(65536, lastCol).End(xlUp).Row
    ActiveSheet.Range("B26:" & ActiveSheet.Cells(lastRow, lastCol).Address).Select

    ' Copiamos el rango
    Selection.Copy

    ' Indicar el libro destino donde pegar
    Workbooks(CotFile).Activate

    ' Indicar la hoja de destino donde pegar
    Sheets("ESTTEMP").Select

    ' Indicar la celda donde pegar los datos
    lastCol = ActiveSheet.Range("A1").End(xlToRight).Column
    lastRow = ActiveSheet.Cells(65536, lastCol).End(xlUp).Row
    Range("A" & lastRow + 1).Select

    ' Para pegar solo como valor
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    ' Limpia Portapapeles
    Application.CutCopyMode = False

    ' Al finalizar el pegado se coloca en celda A2 por si hay muchas filas para que no quede abajo
    Range("A2").Select

    ' Pregunta si desea insertar otro Estimate
    OtroEstimate = "Deseas Importar Otro Estimate?"

    'Mostrar mensaje
    RespOtroEst = MsgBox(OtroEstimate, vbQuestion + vbYesNo)

    If RespOtroEst = vbNo Then
        'Codigo si la respuesta es NO
        ' MsgBox "Selecciono NO!", vbInformation
        GoTo FinImportEst
    Else
        'Codigo si la respuesta es SI
        ' MsgBox "Selecciono SI!", vbInformation
        GoTo IniciaEstimates
    End If

FinImportEst:

    ' SI queremos que nos confirme la importaciãn
    'MsgBox "Los datos han sido copiados correctamente."
    
    Workbooks(CotFile).Activate
    Worksheets("ESTTEMP").Activate
        Range("R1").Value = "Partida"
        Range("S1").Value = "No. Parte"
        Range("T1").Value = "Descripcion"
        Range("U1").Value = "Marca"
        Range("V1").Value = "Unidad de Medida"
        Range("W1").Value = "Cantidad"
        Range("X1").Value = "Proveedor"
        Range("Y1").Value = "Precio Lista"
        Range("Z1").Value = "Duracion del Servicio"
        Range("AA1").Value = "Offer Type"

'Insertar Formulas de Numeracion de Partidas
With ActiveSheet.Range("A2").CurrentRegion
    UltimaFilaE = .Rows(.Rows.Count).Row
    Range("N2:N" & UltimaFilaE).FormulaLocal = "=@EXTRAE(A2,1,@ENCONTRAR(""."",A2)-1)"
    Range("O2:O" & UltimaFilaE).FormulaLocal = "=EXTRAE(A2,ENCONTRAR(""."",A2,1),LARGO(A2)-LARGO(N2))"
    Range("P2").Formula = 1
    Range("P3:P" & UltimaFilaE).FormulaLocal = "=SI(N3=N2,P2,P2+1)"
    Range("P2").Formula = 1
    Range("R2:R" & UltimaFilaE).FormulaLocal = "=CONCATENAR(P2,O2)"
    Range("S2:S" & UltimaFilaE).FormulaLocal = "=B2"
    Range("T2:T" & UltimaFilaE).FormulaLocal = "=C2"
    Range("U2:U" & UltimaFilaE).Value = "CISCO"
    Range("V2:V" & UltimaFilaE).Value = "PIEZA"
    Range("W2:W" & UltimaFilaE).FormulaLocal = "=E2"
    Range("X2:X" & UltimaFilaE).Value = "GRUPO DICE"
    Range("Y2:Y" & UltimaFilaE).FormulaLocal = "=D2"
    Range("Z2:Z" & UltimaFilaE).FormulaLocal = "=I2"
    Range("AA2:AA" & UltimaFilaE).FormulaLocal = "=SI(O(J2=""SUBSCRIPTION"",J2=""SERVICE""),""RO"",SI(J2=""PRODUCT"",""OS"",""""))"
    
End With
    
    ' Al finalizar la insercion de Formulas se coloca en A2 para contar filas
    Range("A2").Select
    'RangeCount = Range(Selection, Selection.End(xlDown)).Rows.Count
    ' MsgBox (RangeCount)
    GoTo FinEstimates
    
NoEstFile:
    MsgBox ("Archivo Estimate No Encontrado")
    GoTo IniciaEstimates

FinEstimates:

End Sub

Function ExisteHojaEst() As Boolean
For h = 1 To Sheets.Count
If Sheets(h).Name = "ESTTEMP" Then
ExisteHojaEst = True
Exit Function
Else
ExisteHojaEst = False
End If
Next h
End Function

Function ExisteHojaInf() As Boolean
For h = 1 To Sheets.Count
If Sheets(h).Name = "INFRATEMP" Then
ExisteHojaInf = True
Exit Function
Else
ExisteHojaInf = False
End If
Next h
End Function
Function ExisteHojaCon() As Boolean
For h = 1 To Sheets.Count
If Sheets(h).Name = "CONSOLIDA" Then
ExisteHojaCon = True
Exit Function
Else
ExisteHojaCon = False
End If
Next h
End Function
Function ExisteHoja3eros() As Boolean
For h = 1 To Sheets.Count
If Sheets(h).Name = "TERCEROS" Then
ExisteHoja3eros = True
Exit Function
Else
ExisteHoja3eros = False
End If
Next h
End Function

Sub CSP()
   ' Obtener elNombre del Archivo Actual
   CotFile = ActiveWorkbook.Name
   
ImportCSP:
    ' Indicar el libro origen desde donde copiar la CSP
    CSPName = Application.InputBox("Introduzca El Nombre del CSP Origen", "CSP Name")

        
    ' Indicar la hoja de destino donde pegar
    Worksheets("INST&SOP").Activate
    Worksheets("INST&SOP").Unprotect "Unified2020!!" 'Desprotege la hoja

' Indicar el libro origen desde donde copiar
    Workbooks(CSPName & ".xlsx").Activate
    On Error GoTo NoCSPFile
    
     ' Indicar la hoja de origen del CSP desde donde copiar
    Worksheets("INST&SOP").Activate

    ' Seleccionar el rango fila/Columna con datos. Rango B6:E15
    ActiveSheet.Range("B6:E15").Select

    ' Copiamos el rango
    Selection.Copy

    ' Indicar el libro destino donde pegar
    Workbooks(CotFile).Activate

    ' Indicar la hoja de destino donde pegar
    Worksheets("INST&SOP").Activate
    Range("B6:E15").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False  'Limpia Portapapeles
    Worksheets("INST&SOP").Protect "Unified2020!!"    'Protege hoja
    
    Workbooks(CSPName & ".xlsx").Activate
    Worksheets("INST&SOP").Activate
    ActiveSheet.Range("C2").Select
    Selection.Copy
    Workbooks(CotFile).Activate
    Worksheets("INST&SOP").Activate
    ActiveSheet.Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False  'Limpia Portapapeles
    
    GoTo FinalizaCSP
    
NoCSPFile:
    MsgBox ("Archivo CSP No Encontrado")
    'GoTo ImportCSP

FinalizaCSP:

End Sub

Sub ConsolidaTemp()

    ' Obtener elNombre del Archivo Actual
    CotFile = ActiveWorkbook.Name
   
    'Crea Hoja Temporal CONSOLIDADA
    Worksheets.Add(After:=Worksheets("P&L")).Name = "CONSOLIDA"
    Worksheets("CONSOLIDA").Activate
    Range("A1").Value = "Partida"
    Range("B1").Value = "No. Parte"
    Range("C1").Value = "Descripcion"
    Range("D1").Value = "Marca"
    Range("E1").Value = "Unidad de Medida"
    Range("F1").Value = "Cantidad"
    Range("G1").Value = "Proveedor"
    Range("H1").Value = "Precio Lista"
    Range("I1").Value = "Duracion del Servicio"
    Range("J1").Value = "Offer Type"

    If ExisteHojaEst Then
        ' Indicar la hoja de origen desde donde copiar
        Worksheets("ESTTEMP").Activate

        ' Seleccionar el rango fila/Columna con datos. Desde celda R2
        Range("R2").Select
        lastRowC = Range(Selection, Selection.End(xlDown)).Rows.Count
        ActiveSheet.Range("R2:AA" & lastRowC + 1).Select

        ' Copiamos el rango
        Selection.Copy
        
        'Selecciona donde pegar y pega valores
        Worksheets("CONSOLIDA").Activate
        Range("A2").Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False  'Limpia Portapapeles
        
        
    Else
        'MsgBox ("NO EXISTE HOJA ESTTEMP")
    End If

    If ExisteHojaInf Then
        ' Indicar la hoja de origen desde donde copiar
        Worksheets("INFRATEMP").Activate

        ' Seleccionar el rango fila/Columna con datos. Desde celda M2
        Range("M2").Select
        lastRowI = Range(Selection, Selection.End(xlDown)).Rows.Count
        ActiveSheet.Range("M2:V" & lastRowI + 1).Select

        ' Copiamos el rango
        Selection.Copy
        
        'Selecciona donde pegar y pega valores
        Worksheets("CONSOLIDA").Activate
        Range("A1").Select
        If ExisteHojaEst Then
            UltimaPartida = Worksheets("ESTTEMP").Cells(65536, "P").End(xlUp).Row
            Worksheets("INFRATEMP").Range("K1").FormulaLocal = Worksheets("ESTTEMP").Range("P" & UltimaPartida)
            lastRowII = Range(Selection, Selection.End(xlDown)).Rows.Count
        Else
            lastRowII = 1
        End If
        Range("A" & lastRowII + 1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False  'Limpia Portapapeles
        
        
    Else
        'MsgBox ("NO EXISTE HOJA INFRATEMP")
    End If
    
    
       If ExisteHoja3eros Then
        ' Indicar la hoja de origen desde donde copiar
        Worksheets("TERCEROS").Activate

        ' Seleccionar el rango fila/Columna con datos. Desde celda M2
        Range("M2").Select
        lastRowI = Range(Selection, Selection.End(xlDown)).Rows.Count
        ActiveSheet.Range("M2:V" & lastRowI + 1).Select

        ' Copiamos el rango
        Selection.Copy
        
        'Selecciona donde pegar y pega valores
        Worksheets("CONSOLIDA").Activate
        Range("A1").Select
        If ExisteHojaInf Then
            UltimaPartida2 = Worksheets("INFRATEMP").Cells(65536, "P").End(xlUp).Row
            Worksheets("TERCEROS").Range("K1").FormulaLocal = Worksheets("INFRATEMP").Range("P" & UltimaPartida2)
            lastRowII = Range(Selection, Selection.End(xlDown)).Rows.Count
        Else
                If ExisteHojaEst Then
                    UltimaPartida2 = Worksheets("ESTTEMP").Cells(65536, "P").End(xlUp).Row
                    Worksheets("TERCEROS").Range("K1").FormulaLocal = Worksheets("ESTTEMP").Range("P" & UltimaPartida2)
                    lastRowII = Range(Selection, Selection.End(xlDown)).Rows.Count
                Else
                    lastRowII = 1
                End If
        End If
        Range("A" & lastRowII + 1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False  'Limpia Portapapeles
        
        
    Else
        'MsgBox ("NO EXISTE HOJA INFRATEMP")
    End If
    
    
    'Borra Hojas Temporales
    Application.DisplayAlerts = False 'switching off the alert button
    If ExisteHojaInf Then
        Worksheets("INFRATEMP").Delete
    Else
    End If
    If ExisteHojaEst Then
        Worksheets("ESTTEMP").Delete
    Else
    End If
    If ExisteHoja3eros Then
    Worksheets("TERCEROS").Delete
    Else
    End If
    Application.DisplayAlerts = True 'switching on the alert button
    
    TotInst = Application.WorksheetFunction.Sum(Worksheets("INST&SOP").Range("B6:B15"))
    TotSop = Application.WorksheetFunction.Sum(Worksheets("INST&SOP").Range("C6:C15"))
    
    If TotInst <> 0 And TotSop <> 0 Then
        Worksheets("CONSOLIDA").Activate
        Range("A1").Select
        lastRowS = Range(Selection, Selection.End(xlDown)).Rows.Count
        If lastRowS > 1000000 Then
            Range("L1").Value = "'1.0"
            Range("M1").FormulaLocal = "=Encontrar(""."",L1)"
            Range("N1").FormulaLocal = "=izquierda(L1,M1-1)"
            Range("O1").Value = Range("N1")
            Range("P1").Value = "'.00"
            lastRowS = 1
        Else
            Range("L1").Value = Range("A" & lastRowS)
            Range("M1").FormulaLocal = "=Encontrar(""."",L1)"
            Range("N1").FormulaLocal = "=izquierda(L1,M1-1)"
            Range("O1").Value = Range("N1") + 1
            Range("P1").Value = "'.00"
        End If
        Range("A" & lastRowS + 1).Value = "=CONCAT(O1:P1)"
        Range("B" & lastRowS + 1).Value = "SRV-INST-UN-CSCO"
        Range("C" & lastRowS + 1).Value = "Servicios de Instalacion, Configuracion y Puesta a Punto"
        Range("D" & lastRowS + 1).Value = "UNIFIED"
        Range("E" & lastRowS + 1).Value = "SERV"
        Range("F" & lastRowS + 1).Value = 1
        Range("G" & lastRowS + 1).Value = "UNIFIED"
        Range("H" & lastRowS + 1).Formula = "='P&L'!B9+'P&L'!B18+'P&L'!B32+'P&L'!B47+'P&L'!B51+'P&L'!B67+'P&L'!B78+'P&L'!B84"
        Range("I" & lastRowS + 1).Value = "N/A"
        Range("J" & lastRowS + 1).Value = "OS"
        Range("O2").Value = Range("O1") + 1
        Range("P2").Value = "'.00"
        Range("Q2").Value = " Meses"
        Range("A" & lastRowS + 2).Value = "=CONCAT(O2:P2)"
        Range("B" & lastRowS + 2).Value = "SRV-SOP-UN-CSCO"
        Range("C" & lastRowS + 2).Value = "Servicio Mensual de Soporte de Ingenieiaa (UN)"
        Range("D" & lastRowS + 2).Value = "UNIFIED"
        Range("E" & lastRowS + 2).Value = "POLIZA"
        Range("F" & lastRowS + 2).Value = "='INST&SOP'!C2"
        Range("G" & lastRowS + 2).Value = "UNIFIED"
        Range("H" & lastRowS + 2).Formula = "=('P&L'!C11+'P&L'!C20+'P&L'!C34+'P&L'!C48+'P&L'!C77+'P&L'!C86)/'INST&SOP'!C2"
        Range("I" & lastRowS + 2).Value = "=CONCAT('INST&SOP'!C2,Q2)"
        Range("J" & lastRowS + 2).Value = "RO"
    ElseIf TotInst <> 0 And TotSop = 0 Then
        Worksheets("CONSOLIDA").Activate
        Range("A1").Select
        lastRowS = Range(Selection, Selection.End(xlDown)).Rows.Count
        If lastRowS > 1000000 Then
            Range("L1").Value = "'1.0"
            Range("M1").FormulaLocal = "=Encontrar(""."",L1)"
            Range("N1").FormulaLocal = "=izquierda(L1,M1-1)"
            Range("O1").Value = Range("N1")
            Range("P1").Value = "'.00"
            lastRowS = 1
        Else
            Range("L1").Value = Range("A" & lastRowS)
            Range("M1").FormulaLocal = "=Encontrar(""."",L1)"
            Range("N1").FormulaLocal = "=izquierda(L1,M1-1)"
            Range("O1").Value = Range("N1") + 1
            Range("P1").Value = "'.00"
        End If
        Range("A" & lastRowS + 1).Value = "=CONCAT(O1:P1)"
        Range("B" & lastRowS + 1).Value = "SRV-INST-UN-CSCO"
        Range("C" & lastRowS + 1).Value = "Servicios de Instalacion, Configuracion y Puesta a Punto"
        Range("D" & lastRowS + 1).Value = "UNIFIED"
        Range("E" & lastRowS + 1).Value = "SERV"
        Range("F" & lastRowS + 1).Value = 1
        Range("G" & lastRowS + 1).Value = "UNIFIED"
        Range("H" & lastRowS + 1).Formula = "='P&L'!B9+'P&L'!B18+'P&L'!B32+'P&L'!B47+'P&L'!B51+'P&L'!B67+'P&L'!B78+'P&L'!B84"
        Range("I" & lastRowS + 1).Value = "N/A"
        Range("J" & lastRowS + 1).Value = "OS"
    ElseIf TotInst = 0 And TotSop <> 0 Then
        Worksheets("CONSOLIDA").Activate
        Range("A1").Select
        lastRowS = Range(Selection, Selection.End(xlDown)).Rows.Count
        If lastRowS > 1000000 Then
            Range("L1").Value = "'1.0"
            Range("M1").FormulaLocal = "=Encontrar(""."",L1)"
            Range("N1").FormulaLocal = "=izquierda(L1,M1-1)"
            Range("O1").Value = Range("N1")
            Range("P1").Value = "'.00"
            lastRowS = 1
        Else
            Range("L1").Value = Range("A" & lastRowS)
            Range("M1").FormulaLocal = "=Encontrar(""."",L1)"
            Range("N1").FormulaLocal = "=izquierda(L1,M1-1)"
            Range("O1").Value = Range("N1") + 1
            Range("P1").Value = "'.00"
        End If
        Range("Q2").Value = " Meses"
        Range("A" & lastRowS + 1).Value = "=CONCAT(O1:P1)"
        Range("B" & lastRowS + 1).Value = "SRV-SOP-UN-CSCO"
        Range("C" & lastRowS + 1).Value = "Servicio Mensual de Soporte de Ingenieiaa (UN)"
        Range("D" & lastRowS + 1).Value = "UNIFIED"
        Range("E" & lastRowS + 1).Value = "POLIZA"
        Range("F" & lastRowS + 1).Value = "='INST&SOP'!C2"
        Range("G" & lastRowS + 1).Value = "UNIFIED"
        Range("H" & lastRowS + 1).Formula = "=('P&L'!C11+'P&L'!C20+'P&L'!C34+'P&L'!C48+'P&L'!C77+'P&L'!C86)/'INST&SOP'!C2"
        Range("I" & lastRowS + 1).Value = "=CONCAT('INST&SOP'!C2,Q2)"
        Range("J" & lastRowS + 1).Value = "RO"
    End If
    
    Range("A" & lastRowS + 1 & ":J" & lastRowS + 2).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Range("L1:P2").Clear
    
End Sub

Sub LLENAQUOTE()

   ' Obtener elNombre del Archivo Actual
   CotFile = ActiveWorkbook.Name
   filainicial = 20
   
    Worksheets("Cotizacion").Activate
    Hoja1.Unprotect Password:="Unified2020!!"
    Range("A" & filainicial).FormulaLocal = "=CONSOLIDA!A2"
    Range("A" & filainicial).Locked = True
    Range("B" & filainicial).FormulaLocal = "=CONSOLIDA!B2"
    Range("B" & filainicial).Locked = True
    Range("C" & filainicial).FormulaLocal = "=CONSOLIDA!C2"
    Range("C" & filainicial).Locked = True
    Range("D" & filainicial).FormulaLocal = "=CONSOLIDA!D2"
    Range("D" & filainicial).Locked = True
    Range("E" & filainicial).FormulaLocal = "=CONSOLIDA!E2"
    Range("E" & filainicial).Locked = True
    Range("F" & filainicial).FormulaR1C1 = "=RC[12]*R2C12"
    Range("F" & filainicial).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Range("F" & filainicial).Locked = True
    Range("G" & filainicial).FormulaLocal = "=CONSOLIDA!F2"
    Range("G" & filainicial).Locked = True
    Range("H" & filainicial).FormulaR1C1 = "=RC[-2]*RC[-1]"
    Range("H" & filainicial).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Range("H" & filainicial).Locked = True
    Range("I" & filainicial).FormulaLocal = "=SI(O(CONSOLIDA!I2=""N/A"",CONSOLIDA!I2=""""),"""",CONCATENAR(IZQUIERDA(CONSOLIDA!I2,2),"" Meses""))"
    Range("I" & filainicial).Locked = True
    Range("J" & filainicial).FormulaLocal = "=CONSOLIDA!H2"
    Range("J" & filainicial).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Range("J" & filainicial).Locked = True
    Range("K" & filainicial).FormulaR1C1 = "=RC[-4]"
    Range("K" & filainicial).Locked = True
    Range("L" & filainicial).FormulaR1C1 = "=RC[-2]*RC[-1]"
    Range("L" & filainicial).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Range("L" & filainicial).Locked = True
    Range("N" & filainicial).FormulaR1C1Local = "=SI(FC[-1]="""",0,BUSCARH(FC[-1],DESCUENTOS,5,0))"
    Range("N" & filainicial).NumberFormat = "0.00%"
    Range("N" & filainicial).Locked = True
    Range("O" & filainicial).FormulaR1C1 = "=RC[-5]*(1-RC[-1])"
    Range("O" & filainicial).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Range("O" & filainicial).Locked = True
    Range("P" & filainicial).FormulaR1C1 = "=RC[-1]*RC[-9]"
    Range("P" & filainicial).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Range("P" & filainicial).Locked = True
    Range("Q" & filainicial).FormulaR1C1Local = "=SI(FC[-4]="""",0,BUSCARH(FC[-4],DESCUENTOS,10,0))"
    Range("Q" & filainicial).NumberFormat = "0.00%"
    Range("Q" & filainicial).Locked = True
    'Range("R" & filainicial).FormulaR1C1Local = "=REDONDEAR(RC[-3]/(1-RC[-1]),2)"
    Range("R" & filainicial).FormulaR1C1 = "=RC[-3]/(1-RC[-1])"
    Range("R" & filainicial).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Range("R" & filainicial).Locked = True
    Range("S" & filainicial).FormulaR1C1 = "=RC[-1]*RC[-12]"
    Range("S" & filainicial).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Range("S" & filainicial).Locked = True
    Range("T" & filainicial).FormulaR1C1Local = "=SI(FC[-10]=0,0,1-(FC[-2]/FC[-10]))"
    Range("T" & filainicial).NumberFormat = "0.00%"
    Range("T" & filainicial).Locked = True
    With Range("M" & filainicial).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=$K$8:$AB$8"
    End With
    Range("M" & filainicial).Locked = False
    With Range("U" & filainicial).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=ARQUITECTURAS"
    End With
    Range("U" & filainicial).Value = "DC_CLD"
    With Range("V" & filainicial).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=INDIRECT(U20)"
    End With
    Range("U" & filainicial).Locked = False
    Range("V" & filainicial).Locked = False
    Range("U" & filainicial).Value = ""
    Range("W" & filainicial).FormulaLocal = "=CONSOLIDA!J2"
    Range("W" & filainicial).Locked = True
    Range("X" & filainicial).FormulaLocal = "=CONSOLIDA!G2"
    Range("X" & filainicial).Locked = True
    Range("Z" & filainicial).FormulaR1C1Local = "=SI(FC[-13]="""",0,(1-BUSCARH(FC[-13],DESCUENTOS,2,0))*FC[-16]*FC[-15])"
    Range("Z" & filainicial).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Range("Z" & filainicial).Locked = True
    Range("AA" & filainicial).FormulaR1C1Local = "=SI(FC[-14]="""",0,(BUSCARH(FC[-14],DESCUENTOS,3,0))*FC[-17]*FC[-16])"
    Range("AA" & filainicial).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Range("AA" & filainicial).Locked = True
    Range("AB" & filainicial).FormulaR1C1 = "=RC[-9]-RC[-12]"
    Range("AB" & filainicial).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Range("AB" & filainicial).Locked = True
    Range("AC" & filainicial).FormulaR1C1Local = "=SI(FC[-16]="""",0,(BUSCARH(FC[-16],DESCUENTOS,4,0))*FC[-19]*FC[-18])"
    Range("AC" & filainicial).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Range("AC" & filainicial).Locked = True
    
    Range("A" & filainicial & ":I" & filainicial).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).Weight = xlMedium
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeLeft).Weight = xlMedium

    
    Worksheets("CONSOLIDA").Activate
    Range("A2").Select
    RangeCount = Range(Selection, Selection.End(xlDown)).Rows.Count
    
    If RangeCount > 1000000 Then
    GoTo UnaSolaLinea
    Else
        Worksheets("Cotizacion").Activate
        Range("A" & filainicial & ":AC" & filainicial).Copy Destination:=Range("A" & filainicial + 1 & ":AC" & RangeCount + filainicial - 1)
        ' ######################################
        Range("A" & filainicial & ":E" & RangeCount + filainicial - 1).Copy
        Range("A" & filainicial & ":E" & RangeCount + filainicial - 1).PasteSpecial Paste:=xlPasteValues
        Range("G" & filainicial & ":G" & RangeCount + filainicial - 1).Copy
        Range("G" & filainicial & ":G" & RangeCount + filainicial - 1).PasteSpecial Paste:=xlPasteValues
        Range("I" & filainicial & ":J" & RangeCount + filainicial - 1).Copy
        Range("I" & filainicial & ":J" & RangeCount + filainicial - 1).PasteSpecial Paste:=xlPasteValues
        Range("W" & filainicial & ":X" & RangeCount + filainicial - 1).Copy
        Range("W" & filainicial & ":X" & RangeCount + filainicial - 1).PasteSpecial Paste:=xlPasteValues
        ' ######################################
        Hoja1.Protect Password:="Unified2020!!"
        Application.CutCopyMode = False  'Limpia Portapapeles
        GoTo FinLlenado
    End If
    
UnaSolaLinea:
        ' ######################################
        Range("A" & filainicial & ":E" & filainicial).Copy
        Range("A" & filainicial & ":E" & filainicial).PasteSpecial Paste:=xlPasteValues
        Range("G" & filainicial & ":G" & filainicial).Copy
        Range("G" & filainicial & ":G" & filainicial).PasteSpecial Paste:=xlPasteValues
        Range("I" & filainicial & ":J" & filainicial).Copy
        Range("I" & filainicial & ":J" & filainicial).PasteSpecial Paste:=xlPasteValues
        Range("W" & filainicial & ":X" & filainicial).Copy
        Range("W" & filainicial & ":X" & filainicial).PasteSpecial Paste:=xlPasteValues
        ' ######################################
        Hoja1.Protect Password:="Unified2020!!"
        Application.CutCopyMode = False  'Limpia Portapapeles
        
FinLlenado:

End Sub

Sub LimpiarArchivo()
    filainit = 20
    If ExisteHojaCon Then
        Worksheets("CONSOLIDA").Delete
    Else
    End If
    Worksheets("Cotizacion").Activate
    Hoja1.Unprotect Password:="Unified2020!!"
    Range("A" & filainit & ":AC2000").Select
        Selection.ClearContents
        Selection.Borders(xlInsideVertical).LineStyle = Excel.XlLineStyle.xlLineStyleNone
        Selection.Clear
    Range("M4").ClearContents
    Range("A" & filainit).Select
    Hoja1.Protect Password:="Unified2020!!"
    Worksheets("INST&SOP").Activate
    Range("B6:B15").ClearContents
    Range("C6:C9").ClearContents
    Range("C12").ClearContents
    Range("C15").ClearContents
    Range("D6:D15").ClearContents
    Range("E6:E9").ClearContents
    Range("E12").ClearContents
    Range("E15").ClearContents
    Range("C2").ClearContents
    Worksheets("Cotizacion").Activate
    Range("A20").Select
    Worksheets("P&L").Unprotect Password:="Unified2020!!"
    Worksheets("P&L").Range("N97").ClearContents
    Worksheets("P&L").Range("I3").ClearContents
    Worksheets("P&L").Range("I13").ClearContents
    Worksheets("P&L").Range("I19").ClearContents
    Worksheets("P&L").Range("I31").ClearContents
    Worksheets("P&L").Range("I44").ClearContents
    Worksheets("P&L").Range("I49").ClearContents
    Worksheets("P&L").Range("I56").ClearContents
    Worksheets("P&L").Range("I63").ClearContents
    Worksheets("P&L").Range("I74").ClearContents
    Worksheets("P&L").Range("I80").ClearContents
    Worksheets("P&L").Range("I83").ClearContents
    Worksheets("P&L").Protect Password:="Unified2020!!"
End Sub

Sub InsertaSubtotales()
   ' Obtener elNombre del Archivo Actual
   CotFile = ActiveWorkbook.Name
 
    Worksheets("Cotizacion").Activate
    Hoja1.Unprotect Password:="Unified2020!!"
    'Range("A19").Activate
    'ActiveCell.EntireRow.Insert
    'ActiveCell.EntireRow.Clear
    'Range(ActiveCell, ActiveCell.Offset(0, 8)).Locked = False
    
    Worksheets("Cotizacion").Activate
    'Hoja1.Protect Password:="Unified2020!!"
    
    
    Range("A21").Select
    
    Dim NumFil As Integer
    Dim TotFilas As Integer
    Dim nnn As String
    
    'TotFilas = Range(Selection, Selection.End(xlDown)).Rows.Count + 20
    NumFil = 21
    Do While Range("A" & NumFil) <> ""
        nnn = Left(Range("A" & NumFil), InStr(Range("A" & NumFil), ".") - 1)
        Prevnnn = Left(Range("A" & NumFil - 1), InStr(Range("A" & NumFil - 1), ".") - 1)
        If nnn <> Prevnnn Then
            Range("A" & NumFil & ":A" & NumFil + 3).Select
            Selection.EntireRow.Insert
            Selection.EntireRow.Clear
            Range("A" & NumFil & ":I" & NumFil).Borders(xlEdgeTop).LineStyle = xlContinuous
            Range("A" & NumFil & ":I" & NumFil).Borders(xlEdgeTop).Weight = xlMedium
            Range("A" & NumFil + 3 & ":I" & NumFil + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
            Range("A" & NumFil + 3 & ":I" & NumFil + 3).Borders(xlEdgeBottom).Weight = xlMedium
            Range("G" & NumFil).Value = "Subtotal (USD)"
            Range("A19:AC19").Copy Range("A" & NumFil + 3 & ":AC" & NumFil + 3)
            Sumare = NumFil - 1
            Do While Range("H" & Sumare) <> ""
                    Sumare = Sumare - 1
            Loop
            Worksheets("CONSOLIDA").Range("L1").Value = "'=SUM(H"
            Worksheets("CONSOLIDA").Range("M1").Value = Sumare + 2
            Worksheets("CONSOLIDA").Range("N1").Value = ":H"
            Worksheets("CONSOLIDA").Range("O1").Value = NumFil - 1
            Worksheets("CONSOLIDA").Range("P1").Value = ")"
            Worksheets("CONSOLIDA").Range("Q1").Value = "=CONCAT(L1:P1)"
            Worksheets("Cotizacion").Range("H" & NumFil).Value = Worksheets("CONSOLIDA").Range("Q1").Value
            Range("G" & NumFil & ":H" & NumFil).Font.Size = 12
            Range("G" & NumFil & ":H" & NumFil).Font.FontStyle = "Bold"
            Range("G" & NumFil & ":H" & NumFil).Font.Color = RGB(0, 51, 153)
            Range(ActiveCell.Offset(2, 0), ActiveCell.Offset(2, 8)).Locked = False
            NumFil = NumFil + 5
        Else
            NumFil = NumFil + 1
        End If
    Loop
    Range("G" & NumFil).Value = "Subtotal (USD)"
    Sumare = NumFil - 1
    Do While Range("H" & Sumare) <> ""
        Sumare = Sumare - 1
    Loop
    Worksheets("CONSOLIDA").Range("L1").Value = "'=SUM(H"
    Worksheets("CONSOLIDA").Range("M1").Value = Sumare + 2
    Worksheets("CONSOLIDA").Range("N1").Value = ":H"
    Worksheets("CONSOLIDA").Range("O1").Value = NumFil - 1
    Worksheets("CONSOLIDA").Range("P1").Value = ")"
    Worksheets("CONSOLIDA").Range("Q1").Value = "=CONCAT(L1:P1)"
    Worksheets("Cotizacion").Range("H" & NumFil).Value = Worksheets("CONSOLIDA").Range("Q1").Value
    'Worksheets("CONSOLIDA").Range("L1").Value = "'=SUM(H"
    'Worksheets("CONSOLIDA").Range("M1").Value = NumFil - 1
    'Worksheets("CONSOLIDA").Range("N1").Value = ":H"
    'Worksheets("CONSOLIDA").Range("O1").Value = NumFil - 1
    'Worksheets("CONSOLIDA").Range("P1").Value = ")"
    'Worksheets("CONSOLIDA").Range("Q1").Value = "=CONCAT(L1:M1,P1)"
    'Worksheets("Cotizacion").Range("H" & NumFil).Value = Worksheets("CONSOLIDA").Range("Q1").Value
    Range("A" & NumFil & ":I" & NumFil).Borders(xlEdgeTop).LineStyle = xlContinuous
    Range("A" & NumFil & ":I" & NumFil).Borders(xlEdgeTop).Weight = xlMedium
    Range("G" & NumFil & ":H" & NumFil).Font.Size = 12
    Range("G" & NumFil & ":H" & NumFil).Font.FontStyle = "Bold"
    Range("G" & NumFil & ":H" & NumFil).Font.Color = RGB(0, 51, 153)
    Worksheets("CONSOLIDA").Range("L1:Q1").ClearContents
    
    Range("G" & NumFil + 2).Value = "TOTAL (USD):"
    Range("G" & NumFil + 2).Font.Size = 14
    Range("G" & NumFil + 2).Font.FontStyle = "Bold"
    Range("G" & NumFil + 2).Font.Color = RGB(0, 51, 153)
    Range("G" & NumFil + 2).Interior.Color = RGB(192, 192, 192)
    
    Dim CeldaBUS As Range
    Dim CeldaTTL As Range
    Range("G" & NumFil).Activate
    Set CeldaBUS = ActiveCell
    Range("H" & NumFil).Activate
    Set CeldaTTL = ActiveCell
    Worksheets("CONSOLIDA").Range("L1").Value = "'=SUMIF(G18:H"
    Worksheets("CONSOLIDA").Range("M1").Value = NumFil
    Worksheets("CONSOLIDA").Range("N1").Value = ",G"
    Worksheets("CONSOLIDA").Range("O1").Value = ",H18:H"
    Worksheets("CONSOLIDA").Range("P1").Value = ")"
    Worksheets("CONSOLIDA").Range("Q1").Value = "=CONCAT(L1,M1,N1,M1,O1,M1,P1)"
    Worksheets("Cotizacion").Range("H" & NumFil + 2).Value = Worksheets("CONSOLIDA").Range("Q1").Value
    
    ' Aqui
    Worksheets("Cotizacion").Range("M4").Formula = Worksheets("Cotizacion").Range("H" & NumFil + 2).Formula
    'Worksheets("Cotizacion").Range("M4").Formula = RC[-5]
    Range("H" & NumFil + 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Range("H" & NumFil + 2).Font.Size = 14
    Range("H" & NumFil + 2).Font.FontStyle = "Bold"
    Range("H" & NumFil + 2).Font.Color = RGB(0, 51, 153)
    Range("H" & NumFil + 2).Interior.Color = RGB(192, 192, 192)
    Worksheets("CONSOLIDA").Range("L1:Q1").ClearContents
    Sheets("CONSOLIDA").Visible = False

    
Range("A" & NumFil + 5).Value = "Los Terminos y Condiciones Comerciales Generales de Venta que rigen los bienes, mercancias, productos, servicios, soportes, licencias y suscripciones comercializados por Unified Networks SA de CV., constituyen el acuerdo y "
    Range("A" & NumFil + 6).Value = "  aceptacion del Cliente   y/o Comprador para las transacciones comerciales encomendadas."
    Range("A" & NumFil + 7).Value = "Definiciones:"
    Range("A" & NumFil + 7).Font.Size = 12
    Range("A" & NumFil + 7).Font.FontStyle = "Bold"
    Range("A" & NumFil + 8).Value = "Por Terminos y Condiciones Generales de Venta se entiende el presente conjunto de reglas, que establecidas de acuerdo definiran la relacion comercial entre las partes Cliente/Comprador y Vendedor."
    Range("A" & NumFil + 9).Value = "Por Vendedor, se entendera, para todos los efectos contractuales a Unified Networks, S.A. de C.V."
    Range("A" & NumFil + 10).Value = "Por Cliente/Comprador se entendera, para todos los efectos contractuales a la persona (fisica/moral) que solicita y acuerda con el Vendedor, la compra de algÏn bien, mercancia, producto, servicios y licencias (En adelante la Orden) que"
    Range("A" & NumFil + 11).Value = "  suscribe al final de la presente Propuesta (En adelante la Cotizacion) en manifestacion de su aceptacion."
    Range("A" & NumFil + 12).Value = ""
    Range("A" & NumFil + 13).Value = "CONDICIONES COMERCIALES GENERALES:"
    Range("A" & NumFil + 13).Font.Size = 12
    Range("A" & NumFil + 13).Font.FontStyle = "Bold"
    Range("B" & NumFil + 14).Value = "I.      A menos que se especifique lo contrario, los precios ofertados son antes de IVA. "
    Range("B" & NumFil + 15).Value = "II.     Los precios cotizados en la presente propuesta son en dolares de los Estados Unidos de Norteamerica."
    Range("B" & NumFil + 16).Value = "III.    Las obligaciones adquiridas en moneda extranjera dentro de la Republica Mexicana seran pagaderas al tipo de cambio publicado en el Diario Oficial de la federacion al dia de pago. "
    Range("B" & NumFil + 17).Value = "IV.    La presente propuesta tendra el caracter obligatoria e irrevocable y surtira los efectos mas amplios que en derecho corresponda."
    Range("B" & NumFil + 18).Value = "V.     El precio total y/o por cada partida considera la contratacion y adquisicion de todos los bienes, servicios, soporte y licencias especificadas en esta propuesta."
    Range("B" & NumFil + 19).Value = "VI.    Las licencias de software adquiridas no son ni seran cancelables. "
    Range("B" & NumFil + 20).Value = "VII.   Las licencias de software que se adquieren en modalidad de suscripcion podran ser interrumpidas temporalmente sin ninguna responsabilidad cuando exista por parte del Cliente un retraso en el pago de la factura"
    Range("B" & NumFil + 21).Value = "          de treinta dias continuos, los cuales se restableceran al pago corriente, la suspension temporal no cancela o exime de la obligacion de pago total al Cliente."
    Range("A" & NumFil + 22).Value = "  "
    Range("A" & NumFil + 23).Value = "CONDICIONES DE PAGO:"
    Range("A" & NumFil + 23).Font.Size = 12
    Range("A" & NumFil + 23).Font.FontStyle = "Bold"
    Range("B" & NumFil + 24).Value = "I.      Pago por Anticipo del 50% del precio total a los 10 dias habiles a la aceptacion y firma de la presente propuesta. "
    Range("B" & NumFil + 25).Value = "II.     Pago por *Entregables del 30% del precio total a la entrega de los bienes (hardware/software/licencias/suscripciones) de esta propuesta."
    Range("B" & NumFil + 26).Value = "III.    Pago por *Finiquito del 20% restante al acta de cierre y entrega. "
    Range("B" & NumFil + 27).Value = "        (*) El pago debera aplicarse dentro de los treinta dias naturales posteriores a la emision y entrega de la factura. "
    Range("B" & NumFil + 28).Value = "IV.    La presente propuesta es valida durante el periodo de 30 treinta dias naturales a partir de la fecha de la misma. "
    Range("B" & NumFil + 29).Value = "V.     A la aceptacion y firma de la presente propuesta al Cliente debera entregar una orden de compra formal con las especificaciones de la propuesta, la omision o no recepcion de la orden de compra, no libera o excluye"
    Range("B" & NumFil + 30).Value = "          al cliente de la obligacion de pago. "
    Range("B" & NumFil + 31).Value = "VI.    En caso de que el cliente cancele, respecto de licencias y suscripciones de software, estas generaran una penalizacion del 100% del pago, pues no se prevee cancelacion anticipada sobre estos."
    Range("B" & NumFil + 32).Value = "VII.   En caso de que el cliente cancele, respecto a bienes y productos (Hardware y Servicios), estos generaran una penalizacion del 30% sobre los bienes y productos cancelados."
    Range("B" & NumFil + 33).Value = "VIII.  El vendedor y los fabricantes se reservan el derecho de realizar cambios en sus precios."
    Range("B" & NumFil + 34).Value = "IX.    En caso de que el cliente, solicite cambios, respecto productos, componentes y servicios, estos generaran una penalizacion del 30% sobre los mismos."
    Range("B" & NumFil + 35).Value = ""
    Range("A" & NumFil + 36).Value = "CONDICIONES PARA LA OPCION DE FINANCIAMIENTO POR ARRENDAMIENTO :"
    Range("A" & NumFil + 36).Font.Size = 12
    Range("A" & NumFil + 36).Font.FontStyle = "Bold"
    Range("B" & NumFil + 37).Value = ""
    Range("B" & NumFil + 38).Value = "I.      Carta de Aceptacion de Condiciones Comerciales firmada por el representante legal del cliente, en donde se especifique el nombre de la Arrendadora/Financiera que estara otorgando la Linea de Credito y a quien"
    Range("B" & NumFil + 39).Value = "          estara facturando Unified Networks;"
    Range("B" & NumFil + 40).Value = "II.     Confirmacion por parte del cliente de tener una Linea de Credito aprobada por parte de una entidad financiera, pudiendo ser dicha entidad, alguna con las que Unified Networks tiene alianza o una tercera a eleccion"
    Range("B" & NumFil + 41).Value = "          del cliente y/o;"
    Range("B" & NumFil + 42).Value = "III.    Carta de autorizacion de procesamiento del pedido de parte de la Arrendadora/Financiera (se tendra que evaluar si la linea de credito aprobada, cubre el monto total de la operacion buscada por el cliente) y/o;"
    Range("B" & NumFil + 43).Value = "IV.    De ser el caso de aun no contar con Linea de Credito aprobada, la confirmacion de que se encuentra en proceso de apertura de Linea de Credito."
    Range("B" & NumFil + 44).Value = ""
    Range("A" & NumFil + 45).Value = "ALCANCES TECNICOS:"
    Range("A" & NumFil + 45).Font.Size = 12
    Range("A" & NumFil + 45).Font.FontStyle = "Bold"
    Range("B" & NumFil + 46).Value = "1.    Los tiempos de entrega establecidos iniciaràn a partir de la fecha de recepcion de la orden de compra y/o firma del presente acuerdo, mismos que estaran sujetos a la disponibilidad del fabricante y/o mayorista"
    Range("B" & NumFil + 47).Value = "         sin responsabilidad alguna para Unified Networks. Por tal motivo el cliente debera considerar la disponibilidad para evitar posibles retrasos en el proyecto."
    Range("B" & NumFil + 48).Value = ""
    Range("A" & NumFil + 49).Value = "ENTERADO DE SU CONTENIDO Y AL MANIFESTAR SU INTENCI”N DE COTIZAR, Y SU DECISION DE COMPRA,"
    Range("A" & NumFil + 49).Font.Size = 12
    Range("A" & NumFil + 49).Font.FontStyle = "Bold"
    Range("B" & NumFil + 50).Value = ""
    Range("A" & NumFil + 51).Value = "LA FIRMA DE LA PRESENTE PROPUESTA COMERCIAL ES VINCULANTE Y OBLIGA A LAS PARTES A LA ACEPTACI”N DE LOS TERMINOS Y CONDICIONES GENERALES DE "
    Range("A" & NumFil + 51).Font.Size = 12
    Range("A" & NumFil + 51).Font.FontStyle = "Bold"
    Range("A" & NumFil + 52).Value = "  VENTA, PAGO Y ENTREGABLES."
    Range("A" & NumFil + 52).Font.Size = 12
    Range("A" & NumFil + 52).Font.FontStyle = "Bold"
 
    
    Range("B" & NumFil + 53).Value = ""
    Range("B" & NumFil + 54).Value = "ACEPTACION"
    Range("B" & NumFil + 54).Font.Size = 12
    Range("B" & NumFil + 54).Font.FontStyle = "Bold"
    Range("B" & NumFil + 55).Value = "CLIENTE: "
    Range("B" & NumFil + 55).Font.Size = 12
    Range("B" & NumFil + 55).Font.FontStyle = "Bold"
    Range("B" & NumFil + 56).Value = "FIRMA:__________________________________"
    Range("B" & NumFil + 56).Font.Size = 12
    Range("B" & NumFil + 56).Font.FontStyle = "Bold"
    Range("B" & NumFil + 57).Value = "Nombre del solicitante y autorizante por el CLIENTE :"
    Range("B" & NumFil + 57).Font.Size = 12
    Range("B" & NumFil + 57).Font.FontStyle = "Bold"
    Range("B" & NumFil + 58).Value = "PUESTO:"
    Range("B" & NumFil + 58).Font.Size = 12
    Range("B" & NumFil + 58).Font.FontStyle = "Bold"
    Range("B" & NumFil + 59).Value = "FECHA:"
    Range("B" & NumFil + 59).Font.Size = 12
    Range("B" & NumFil + 59).Font.FontStyle = "Bold"
    
    
    Worksheets("Cotizacion").Activate
   
   textoBUS = "SRV-INST-UN"
      Set celda = Range("B:B").Find(What:=textoBUS, LookIn:=xlValues)
   If Not celda Is Nothing Then
      Range("M" & celda.Row).Locked = True
      With Range("M" & celda.Row).Validation
        .Delete
      End With
      Range("U" & celda.Row).Locked = True
      With Range("U" & celda.Row).Validation
        .Delete
      End With
      Range("V" & celda.Row).Locked = True
      With Range("V" & celda.Row).Validation
        .Delete
      End With
   Else
      'MsgBox "No se ha encontrado el texto a buscar"
   End If
   
   textoBUS2 = "SRV-SOP-UN"
      Set celda2 = Range("B:B").Find(What:=textoBUS2, LookIn:=xlValues)
   If Not celda2 Is Nothing Then
      Range("M" & celda2.Row).Locked = True
      With Range("M" & celda2.Row).Validation
        .Delete
      End With
      Range("U" & celda2.Row).Locked = True
      With Range("U" & celda2.Row).Validation
        .Delete
      End With
      Range("V" & celda2.Row).Locked = True
      With Range("V" & celda2.Row).Validation
        .Delete
      End With
   Else
      'MsgBox "No se ha encontrado el texto a buscar"
   End If
   
   Hoja1.Protect Password:="Unified2020!!"

    'If Worksheets("INST&SOP").Range("C2") <> "" Then
    'Worksheets("P&L").Unprotect Password:="Unified2020!!"
        
    TotInst2 = Application.WorksheetFunction.Sum(Worksheets("INST&SOP").Range("B6:B15"))
    TotSop2 = Application.WorksheetFunction.Sum(Worksheets("INST&SOP").Range("C6:C15"))
    Worksheets("P&L").Unprotect Password:="Unified2020!!"
    
    If TotInst2 = 0 And TotSop2 = 0 Then
        Worksheets("P&L").Range("N97").Value = 0
    ElseIf TotInst2 <> 0 And TotSop2 = 0 Then
        Worksheets("P&L").Range("N97").Value = Worksheets("Cotizacion").Range("H" & celda.Row).Value
    ElseIf TotInst2 = 0 And TotSop2 <> 0 Then
        Worksheets("P&L").Range("N97").Value = Worksheets("Cotizacion").Range("H" & celda2.Row).Value
    ElseIf TotInst2 <> 0 And TotSop2 <> 0 Then
        Worksheets("P&L").Range("N97").Value = Worksheets("Cotizacion").Range("H" & celda.Row).Value + Worksheets("Cotizacion").Range("H" & celda2.Row).Value
    End If
    Worksheets("P&L").Protect Password:="Unified2020!!"
    'Else
    'End If
End Sub

Sub TESTSUMA()
    Worksheets("CONSOLIDA").Range("L1").Value = "'=SUM(H"
    Worksheets("CONSOLIDA").Range("M1").Value = Application.InputBox("Celda Inicial", "Inicio")
    Worksheets("CONSOLIDA").Range("N1").Value = ":H"
    Worksheets("CONSOLIDA").Range("O1").Value = Application.InputBox("Celda Final", "Final")
    Worksheets("CONSOLIDA").Range("P1").Value = ")"
    Worksheets("CONSOLIDA").Range("Q1").Value = "=CONCAT(L1:P1)"
    Worksheets("CONSOLIDA").Range("L2").Value = Worksheets("CONSOLIDA").Range("Q1").Value

End Sub

Sub BusTexCel()
   Hoja1.Unprotect Password:="Unified2020!!"
   textoBUS = "SRV-INST-UN"
      Set celda = Range("B:B").Find(What:=textoBUS, LookIn:=xlValues)
   If Not celda Is Nothing Then
      Range("M" & celda.Row).Locked = True
      With Range("M" & celda.Row).Validation
        .Delete
      End With
      Range("U" & celda.Row).Locked = True
      With Range("U" & celda.Row).Validation
        .Delete
      End With
      Range("V" & celda.Row).Locked = True
      With Range("V" & celda.Row).Validation
        .Delete
      End With
   Else
      'MsgBox "No se ha encontrado el texto a buscar"
   End If
   
   textoBUS2 = "SRV-SOP-UN"
      Set celda = Range("B:B").Find(What:=textoBUS2, LookIn:=xlValues)
   If Not celda Is Nothing Then
      Range("M" & celda.Row).Locked = True
      With Range("M" & celda.Row).Validation
        .Delete
      End With
      Range("U" & celda.Row).Locked = True
      With Range("U" & celda.Row).Validation
        .Delete
      End With
      Range("V" & celda.Row).Locked = True
      With Range("V" & celda.Row).Validation
        .Delete
      End With
   Else
      'MsgBox "No se ha encontrado el texto a buscar"
   End If
   
   Hoja1.Protect Password:="Unified2020!!"
End Sub


Sub IntegraCotizaciones()

    ' Obtener elNombre del Archivo Actual
    ConsolidatedFile = ActiveWorkbook.Name
    

    ' Indicar el libro origen desde donde copiar
    CotOrigen = Application.InputBox("Introduzca El Nombre de la Cotizacion Origen que deseas importar", "Cotizacion a consolidar")

    ' Indicar el libro origen desde donde copiar
    Workbooks(CotOrigen & ".xlsm").Activate
    On Error GoTo NoQuoteFile
    
    Worksheets ("Cotizacion"), Range("K8").Select
    lastMargCol = Range(Selection, Selection.End(xlToRight)).Column
    ActiveSheet.Range("K8:" & lastMargCol & "11").Select
    



NoQuoteFile:


End Sub

Sub ProtectCot()
    Hoja1.Protect Password:="Unified2020!!"
End Sub

Sub UnProtectCot()
    Hoja1.Unprotect Password:="Unified2020!!"
End Sub

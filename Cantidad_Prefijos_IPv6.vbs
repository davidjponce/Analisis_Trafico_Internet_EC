Attribute VB_Name = "Module1"
'----------------------------------------------------
' Búsqueda de cantidad de prefijos IPv6
' Elaborado por: David Ponce Ortz
'----------------PREFIJOS IPv6-----------------------
Public Sub PrefijosIPv6()
    Set AppExcel = CreateObject("Excel.Application")
    '1) Abrir de archivo TXT
    Dim fd As Office.FileDialog
    Dim archivoTXT As String
InicioPrograma:
    'Reseteo variables
    i = 0
    ij = 0
ValorHito = "Abro archivo TXT"
    ChDir "D:\Tesis"
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.Clear
        .Filters.Add "Text document (*.txt*)", "*.txt*", 1
        .Title = "Escoge el archivo TXT resumen"
        .AllowMultiSelect = False
        .InitialFileName = "*resumen*.*"    '"D:\Tesis"
        result = .Show
        
        If (result <> 0) Then
            archivoTXT = Trim(.SelectedItems.Item(1))
        Else
            GoTo FinPrograma
        End If
    End With
    
    '2)Obtengo la columna de ASN y la columna de Prefijos IPv4 y almaceno en ARRAYS
    'Declaro Arrays y Dictionario
    ReDim ASN(20)
    ReDim CantidadIPv4(20)
    ReDim dataTXT(6)
    Set dictASNs = CreateObject("Scripting.Dictionary")
    dictASNs.RemoveAll
    numeroColumna = 0
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(archivoTXT)
    
ValorHito = "Almacenamiento ASN e IPv6"
    'Almaceno la columna de ASN y Cantidad en el array
    Do Until objFile.AtEndOfStream
        dataArchivo = objFile.ReadLine
        If dataArchivo = "# RESUMEN:" Then GoTo SaltoHeaders
        If InStr(dataArchivo, "EMPRESA") <> 0 Then GoTo SaltoHeaders
        
        If Len(dataArchivo) > 0 Then
            'Obtengo el último registro
            prueba = Split(dataArchivo, vbTab)
            'Almaceno solo los valores en un Array de 4
            ij = 0
            For iteradorPrueba = 0 To 12
                On Error Resume Next
                If prueba(iteradorPrueba) <> "" Then
                    dataTXT(ij) = prueba(iteradorPrueba)
                    ij = ij + 1
                End If
            Next
            ASN(i) = dataTXT(1) 'ASN
            CantidadIPv4(i) = dataTXT(4) 'IPv6
            dictASNs.Add CStr(ASN(i)), CInt(CantidadIPv4(i)) 'Agrego diccionario
            ij = 0 'Reseteo variable
            'Limpio 2 arrays
            Erase dataTXT()
            ReDim dataTXT(6)
            i = i + 1 'Aumento el iterador
        End If
SaltoHeaders:
    Loop
    objFile.Close
    'dictASNs.Items() (IPv4)
    'dictASNs.Keys() (ASN)
    'Limpio objetos
    Set objFSO = Nothing
    Set objFile = Nothing
    i = 0
  ValorHito = "Asignacion ASN"
AsignacionASN:
    '3) Hago Match con la data de la hoja Excel
    '-- Me ubico en la última columna
    Set Thiswb = ThisWorkbook
    Set DataWorkSheet = Thiswb.Sheets("Prefijos_IPv6") 'Seteo la hoja
    fila = 5 'Iterador de programas
    DataWorkSheet.Activate
    Range("C" & fila).Select 'Me ubico en el primer ASN
    'Ubico la última columna
    iteradorMeses = Sheets("Prefijos_IPv6").Cells(5, Columns.Count).End(xlToLeft).Column 'Iterador de columnas (itera los Meses)
    celdaIPv4 = Cells(fila, iteradorMeses).Value 'marco la celda del valor IPv4
    nombreASN = UCase(CStr(DataWorkSheet.Range("C" & fila)))
        'Recorro todo el diccionario y escribo el que hace match
        'key --> ASN
        'dict(key) --> Cantidad IPv4
        Dim key As Variant
        For Each key In dictASNs.Keys
MatchASN:
                    ValorHito = "Match ASN"
            
                        MatchASN = AppExcel.WorksheetFunction.Match(UCase(key), DataWorkSheet.Range("C5:C1000"), 0)
                        celdaASN = MatchASN + 1
                        'Escribo el valor en el excel
                        Cells(celdaASN + 3, iteradorMeses + 1).Value = dictASNs(key)
                        GoTo SiguienteASN
                        
NuevoASN:
             ValorHito = "Crear nuevo ASN"
                       'Ingreso el nuevo ASN
                       ultimoRegistro = Sheets("Prefijos_IPv6").Cells(Rows.Count, 3).End(xlUp).Row
                       ultimoRegistro = ultimoRegistro + 1
                       DataWorkSheet.Range("C" & ultimoRegistro).Value = UCase(key) 'Ingreso el valor que no dio Match
                        GoTo MatchASN 'verifico nuevamente la información
            
SiguienteASN:
            
             ValorHito = "Siguiente ASN"
                    On Error GoTo -1 'Desactivo Errores con Resume Next
        Next key

        'Marco con 0 las columnas que no salieron
        ultimoRegistro = Sheets("Prefijos_IPv6").Cells(Rows.Count, 3).End(xlUp).Row
        For it = 5 To ultimoRegistro
            If Cells(it, iteradorMeses + 1).Value = Empty Then
                Cells(it, iteradorMeses + 1).Value = 0
            End If
        Next
        GoTo InicioPrograma

FinPrograma:
End Sub







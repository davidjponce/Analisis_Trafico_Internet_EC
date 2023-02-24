Attribute VB_Name = "Module3"
'---------------------------------------------------------
' Búsqueda de capacidades Router UIO/GYE
' Elaborado por: David Ponce Ortz
'---------------Capacidades ROUTER UIO/GYE--------------------

Public Sub Capacidades_UIO()
    Set AppExcel = CreateObject("Excel.Application")
    '1) Abrir de archivo TXT
    Dim fd As Office.FileDialog
    Dim archivoEXCEL As String
InicioPrograma:
    'Reseteo variables
    i = 0
    k = 0
    ij = 0
ValorHito = "Abro archivo Excel"
    ChDir "D:\Tesis"
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx", 1
        .Title = "Escoge el archivo Excel resumen"
        .AllowMultiSelect = False
        .InitialFileName = "*ixp*.*"    '"D:\Tesis"
        result = .Show

        If (result <> 0) Then
            archivoEXCEL = Trim(.SelectedItems.Item(1))
        Else
            GoTo FinPrograma
        End If
    End With

    '2)Obtengo la columna de ASN y la columna de Prefijos IPv4 y almaceno en ARRAYS
    'Declaro Arrays y Dictionario
    ReDim ASN(20)
    ReDim CantidadIPv4(20)
    ReDim dataTXT(6)
    Set dictENTRANTE_UIO = CreateObject("Scripting.Dictionary")
    Set dictSALIENTE_UIO = CreateObject("Scripting.Dictionary")
    Set dictENTRANTE_GYE = CreateObject("Scripting.Dictionary")
    Set dictSALIENTE_GYE = CreateObject("Scripting.Dictionary")
    dictENTRANTE_UIO.RemoveAll
    dictSALIENTE_UIO.RemoveAll
    dictENTRANTE_GYE.RemoveAll
    dictSALIENTE_GYE.RemoveAll
    numeroColumna = 0
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(archivoEXCEL)

ValorHito = "Almacenamiento ASN y Capacidades"
    'Declaro variables del archivo a abrir
    Dim DataWB, Thiswb As Workbook
    Dim DataSHEET As Object
    Dim sheetName As String
    
    'Abro libro de la DATA a importar
    Set DataWB = Workbooks.Open(archivoEXCEL)
    Set DataSHEET = DataWB.Sheets("Total")
    'Me ubico en la hoja Total
    DataSHEET.Activate 'Me ubico en la hoja Total
    'Pongo ceros en los valores
    Range("C5:F33").Select   '(medida desde ENERO 2021 hasta )
    'Range("C5:F29").Select '(medida  HASTA FEBRERO 2020)
    'Range("C8:F29").Select '(medida desde MARZO 2020 hasta DICIEMBRE 2020)
    
    Dim cell As Range
    For Each cell In Selection
        If cell.Value = "" Then
            cell.Value = 0
        End If
    Next cell
    
    'Declaro variables del archivo MAIN
    Set Thiswb = Workbooks("ProcesamientoData.xlsm")
    Set DataCapacidadUIO = Thiswb.Sheets("Capacidad_UIO") 'Seteo la hoja
    Set DataCapacidadGYE = Thiswb.Sheets("Capacidad_GYE") 'Seteo la hoja
    
    'Declaro Arrays y Dictionario
    ReDim pruebaASN(30), pruebaEntradaUIO(30), pruebaSalidaUIO(30), pruebaEntradaGYE(30), pruebaSalidaGYE(30)
    ReDim ASN(30)
    ReDim entradaUIO(30)
    ReDim salidaUIO(30)
    ReDim entradaGYE(30)
    ReDim salidaGYE(30)
    
    'Recupero los nombres de ASNs y datos de capacidades
    DataSHEET.Activate 'Me ubico en la hoja DATA
    filaASN = 5
    filaEntradaUIO = 5
    filaSalidaUIO = 5
    filaEntradaGYE = 5
    filaSalidaGYE = 5
    
    'Almaceno todos los valores
    m = 0
    For ij = 5 To 33 '(medida desde ENERO 2021 hasta NOVIEMBRE 2022)
    'For ij = 8 To 29 '(medida desde MARZO 2020 hasta DICIEMBRE 2020)
    'For ij = 5 To 27  (medida hasta FEBRERO 2020)
        pruebaASN(m) = DataSHEET.Range("B" & ij).Value 'ASN
        pruebaEntradaUIO(m) = DataSHEET.Range("C" & ij).Value 'Tráfico de Entrada UIO
        pruebaSalidaUIO(m) = DataSHEET.Range("D" & ij).Value 'Tráfico de Salida UIO
        pruebaEntradaGYE(m) = DataSHEET.Range("E" & ij).Value 'Tráfico de Entrada GYE
        pruebaSalidaGYE(m) = DataSHEET.Range("F" & ij).Value 'Tráfico de Salida GYE
        m = m + 1
    Next
    
    'Procesamiento Data
    ij = 0 'Seteo variable
    For iteradorPrueba = LBound(pruebaASN) To UBound(pruebaASN)
        'Si el valor no es nulo, almaceno en el nuevo Array de Data
        If pruebaASN(iteradorPrueba) <> "" Or pruebaASN(iteradorPrueba) <> 0 Then
            ASN(ij) = pruebaASN(iteradorPrueba) 'ASN
            entradaUIO(ij) = pruebaEntradaUIO(iteradorPrueba) 'Tráfico de Entrada UIO
            salidaUIO(ij) = pruebaSalidaUIO(iteradorPrueba) 'Tráfico de Salida UIO
            entradaGYE(ij) = pruebaEntradaGYE(iteradorPrueba) 'Tráfico de Entrada UIO
            salidaGYE(ij) = pruebaSalidaGYE(iteradorPrueba) 'Tráfico de Salida UIO
            ij = ij + 1
        End If
    Next
    
    'Reseteo variable
    ijk = 0
    'Almaceno en los diccionarios
    While ASN(ijk) <> "" 'Agrego en los diccionarios los valores NO nulos
        
        If ASN(ijk) = 0 Then GoTo SiguienteRegistro
        'For iteradorDiccionarios = LBound(pruebaASN) To UBound(pruebaASN)
        dictENTRANTE_UIO.Add UCase(CStr(ASN(ijk))), entradaUIO(ijk)
        dictSALIENTE_UIO.Add UCase(CStr(ASN(ijk))), salidaUIO(ijk)
        dictENTRANTE_GYE.Add UCase(CStr(ASN(ijk))), entradaGYE(ijk)
        dictSALIENTE_GYE.Add UCase(CStr(ASN(ijk))), salidaGYE(ijk)

'Si el ASN es 0 voy al siguiente registro
SiguienteRegistro:
        'Aumento iterador
        ijk = ijk + 1
        'Next
    Wend
    'Reseteo variable
    ijk = 0
    'Limpio Arrays
    Erase pruebaASN(), pruebaEntradaUIO(), pruebaSalidaUIO(), pruebaEntradaGYE(), pruebaSalidaGYE()
    Erase ASN(), entradaUIO(), salidaUIO(), entradaGYE(), salidaGYE()

    ReDim pruebaASN(30), pruebaEntradaUIO(30), pruebaSalidaUIO(30), pruebaEntradaGYE(30), pruebaSalidaGYE(30)
    ReDim ASN(30)
    ReDim entradaUIO(30)
    ReDim salidaUIO(30)
    ReDim entradaGYE(30)
    ReDim salidaGYE(30)

    

  ValorHito = "Asignacion ASN UIO"
AsignacionASN:
    '3) Hago Match con la data de la hoja Excel
    '-- Me ubico en la última columna
    'Set Thiswb = ThisWorkbook
    'Set DataWorkSheet = Thiswb.Sheets("Capacidad_UIO") 'Seteo la hoja
    fila = 6 'Iterador de programas
    'Primero marco los ASNs de QUITO
    DataCapacidadUIO.Activate
    Range("C" & fila).Select 'Me ubico en el primer ASN
    'Ubico la última columna
    iteradorMeses = Sheets("Capacidad_UIO").Cells(6, Columns.Count).End(xlToLeft).Column 'Iterador de columnas (itera los Meses)
    celdaCapacidadEntrante = Cells(fila, iteradorMeses).Value 'marco la celda del valor trafico entrante UIO
    'nombreASN = UCase(CStr(DataWorkSheet.Range("C" & fila)))
        'Recorro todo el diccionario y escribo el que hace match
        'key --> ASN
        'dict(key) --> Trafico Entrante UIO
        Dim keyUIO_Entrante, keyUIO_Saliente  As Variant
        For Each keyUIO_Entrante In dictENTRANTE_UIO.Keys
                keyUIO_Saliente = keyUIO_Entrante
MatchASN:
                    ValorHito = "Match ASN"
                                On Error GoTo NuevoASN
                                MatchASN = AppExcel.WorksheetFunction.Match(UCase(keyUIO_Entrante), DataCapacidadUIO.Range("C6:C1000"), 0)
                                celdaASN = MatchASN + 1
                                'Escribo el valor en el excel
                                Cells(celdaASN + 4, iteradorMeses + 1).Value = dictENTRANTE_UIO(keyUIO_Entrante) 'Entrante UIO
                                Cells(celdaASN + 4, iteradorMeses + 2).Value = dictSALIENTE_UIO(keyUIO_Saliente) 'Saliente UIO
                                GoTo SiguienteASN
        
NuevoASN:
                     ValorHito = "Crear nuevo ASN"
                               'Ingreso el nuevo ASN
                               ultimoRegistro = Sheets("Capacidad_UIO").Cells(Rows.Count, 3).End(xlUp).Row
                               ultimoRegistro = ultimoRegistro + 1
                               DataCapacidadUIO.Range("C" & ultimoRegistro).Value = UCase(keyUIO_Entrante) 'Ingreso el valor que no dio Match
                                GoTo MatchASN 'verifico nuevamente la información
        
SiguienteASN:
        
                     ValorHito = "Siguiente ASN"
                            On Error GoTo -1 'Desactivo Errores con Resume Next
                
                'Next keyUIO_Saliente
        Next keyUIO_Entrante

        'Marco con 0 las columnas que no salieron
        ultimoRegistro = Sheets("Capacidad_UIO").Cells(Rows.Count, 3).End(xlUp).Row
        For it = 6 To ultimoRegistro
            If Cells(it, iteradorMeses + 1).Value = Empty Or Cells(it, iteradorMeses + 2).Value = Empty Then
                Cells(it, iteradorMeses + 1).Value = 0
                Cells(it, iteradorMeses + 2).Value = 0
            End If
        Next
        
        
        
        '----------- ASIGNACION DATA GUAYAQUIL
         ValorHito = "Asignacion ASN GYE"
AsignacionASN_2:
    '3) Hago Match con la data de la hoja Excel
    '-- Me ubico en la última columna
    'Set Thiswb = ThisWorkbook
    'Set DataWorkSheet = Thiswb.Sheets("Capacidad_UIO") 'Seteo la hoja
    fila = 6 'Iterador de programas
    'Primero marco los ASNs de QUITO
    DataCapacidadGYE.Activate
    Range("C" & fila).Select 'Me ubico en el primer ASN
    'Ubico la última columna
    iteradorMeses = Sheets("Capacidad_GYE").Cells(6, Columns.Count).End(xlToLeft).Column 'Iterador de columnas (itera los Meses)
    celdaCapacidadEntrante = Cells(fila, iteradorMeses).Value 'marco la celda del valor trafico entrante UIO
    'nombreASN = UCase(CStr(DataWorkSheet.Range("C" & fila)))
        'Recorro todo el diccionario y escribo el que hace match
        'key --> ASN
        'dict(key) --> Trafico Entrante UIO
        Dim keyGYE_Entrante, keyGYE_Saliente  As Variant
        For Each keyGYE_Entrante In dictENTRANTE_GYE.Keys
                'For Each keyGYE_Saliente In dictSALIENTE_GYE.Keys
                keyGYE_Saliente = keyGYE_Entrante
MatchASN_2:
                    ValorHito = "Match ASN"
                                On Error GoTo NuevoASN_2
                                MatchASN = AppExcel.WorksheetFunction.Match(UCase(keyGYE_Entrante), DataCapacidadGYE.Range("C6:C1000"), 0)
                                celdaASN = MatchASN + 1
                                'Escribo el valor en el excel
                                Cells(celdaASN + 4, iteradorMeses + 1).Value = dictENTRANTE_GYE(keyGYE_Entrante) 'Entrante GYE
                                Cells(celdaASN + 4, iteradorMeses + 2).Value = dictSALIENTE_GYE(keyGYE_Saliente) 'Saliente GYE
                                GoTo SiguienteASN_2
        
NuevoASN_2:
                     ValorHito = "Crear nuevo ASN"
                               'Ingreso el nuevo ASN
                               ultimoRegistro = Sheets("Capacidad_GYE").Cells(Rows.Count, 3).End(xlUp).Row
                               ultimoRegistro = ultimoRegistro + 1
                               DataCapacidadGYE.Range("C" & ultimoRegistro).Value = UCase(keyGYE_Entrante) 'Ingreso el valor que no dio Match
                                GoTo MatchASN_2 'verifico nuevamente la información
        
SiguienteASN_2:
        
                     ValorHito = "Siguiente ASN"
                            On Error GoTo -1 'Desactivo Errores con Resume Next
                
                'Next keyGYE_Saliente
        Next keyGYE_Entrante



        'Marco con 0 las columnas que no salieron
        ultimoRegistro = Sheets("Capacidad_GYE").Cells(Rows.Count, 3).End(xlUp).Row
        For it = 6 To ultimoRegistro
            If Cells(it, iteradorMeses + 1).Value = Empty Or Cells(it, iteradorMeses + 2).Value = Empty Then
                Cells(it, iteradorMeses + 1).Value = 0
                Cells(it, iteradorMeses + 2).Value = 0
            End If
        Next
        
        'Me ubico hoja UIO
        DataCapacidadUIO.Activate
        
        'Cierro el archivo y NO guardo
        DataWB.Close SaveChanges:=False
        
        GoTo InicioPrograma

FinPrograma:
End Sub

Attribute VB_Name = "mdlProcesamientoTelcel"
Option Explicit

Private Const COL_NUMA As Long = 1        ' A - Numero que hizo la consulta (dueño de la linea investigada)
Private Const COL_TIPO As Long = 2        ' B - Tipo de evento tal como lo entrega Telcel (texto original)
Private Const COL_SUBTIPO As Long = 3     ' C - Clasificacion simplificada (LL-E, LL-S, M-E, M-S, Datos...)
Private Const COL_SENTIDO As Long = 4     ' D - "E" (entrante) o "S" (saliente)
Private Const COL_COLOR As Long = 5       ' E - Color sugerido para semaforizar en el analisis de vinculos
Private Const COL_NUMERO_A As Long = 6    ' F - Numero A (quien origina), sin prefijo 52
Private Const COL_NUMB As Long = 7        ' G - Numero B (contraparte), sin prefijo 52
Private Const COL_NUMBIFETEL As Long = 8  ' H - Primeros 6 digitos de NUMB (prefijo IFT de operador/region)
Private Const COL_MUNICIPIO As Long = 9   ' I - Municipio de la antena (dato original de Telcel)
Private Const COL_ESTADO As Long = 10     ' J - Estado de la antena (dato original de Telcel)
Private Const COL_RED As Long = 11        ' K - Tipo de red (2G/3G/4G, dato original)
Private Const COL_PROVEEDOR As Long = 12  ' L - Proveedor de servicios (dato original)
Private Const COL_FECHA As Long = 13      ' M - Fecha del evento
Private Const COL_DIA As Long = 14        ' N - Dia de la semana, calculado a partir de FECHA
Private Const COL_HORA As Long = 15       ' O - Hora tal como la entrega Telcel (texto variable: 12h/24h/AM-PM)
Private Const COL_HORA24 As Long = 16     ' P - Hora normalizada a formato 24h (0-23), calculada
Private Const COL_ETAPA As Long = 17      ' Q - Madrugada/Mañana/Tarde/Noche, calculada a partir de HORA24
Private Const COL_SEG As Long = 18        ' R - Duracion en segundos (dato original)
Private Const COL_MIN As Long = 19        ' S - Duracion en minutos, calculada (SEG / 60)
Private Const COL_IMEI As Long = 20       ' T - IMEI del dispositivo (dato original, se limpia el texto "null")
Private Const COL_LATITUD As Long = 21    ' U - Latitud de la antena en formato DMS (dato original, se corrige comillas)
Private Const COL_LONGITUD As Long = 22   ' V - Longitud de la antena en formato DMS (idem)
Private Const COL_AZIMUTH As Long = 23    ' W - Azimuth de la antena (dato original, se conserva tal cual)
Private Const COL_BTS As Long = 24        ' X - "Latitud, Longitud" concatenado (para identificar antena unica)
Private Const COL_BTS_AZIMUTH As Long = 25 ' Y - "BTS / Azimuth" concatenado (referencia dentro de "Procesado")
Private Const COL_NUMERO_A_COPIA As Long = 26 ' Z  - copia de NUMERO_A (insumo de la tabla dinamica de contactos)
Private Const COL_NUMB_COPIA As Long = 27     ' AA - copia de NUMB (idem)
Private Const ULTIMA_COLUMNA As Long = COL_NUMB_COPIA

Private Const MIN_COLUMNAS_ORIGINAL As Long = 11

Private Const SUB_DATOS As String = "Datos"
Private Const SUB_DATOS_WIFI As String = "Datos WiFi"
Private Const SUB_MENSAJE_ENTRANTE As String = "M - E"
Private Const SUB_MENSAJE_SALIENTE As String = "M - S"
Private Const SUB_LLAMADA_ENTRANTE As String = "LL - E"
Private Const SUB_LLAMADA_SALIENTE As String = "LL - S"
Private Const SUB_TRANSITO As String = "Transito"
Private Const SUB_TRANSFER As String = "Transfer"

Private Const COLOR_AMARILLO As String = "Amarillo"
Private Const COLOR_AZUL As String = "Azul"
Private Const COLOR_ROJO As String = "Rojo"

Private Const SENTIDO_ENTRANTE As String = "E"
Private Const SENTIDO_SALIENTE As String = "S"

Private Const NOMBRE_HOJA_PROCESADO As String = "Procesado"
Private Const NOMBRE_HOJA_FRECUENCIAS As String = "Frecuencias de comunicacion"

Private Const DIBUJAR_SECTORES As Boolean = True
Private Const RADIO_SECTOR_M As Double = 500        ' radio de la cuña en metros
Private Const ANCHO_SECTOR_GRADOS As Double = 120   ' apertura de la cuña
Private Const PI_VALOR As Double = 3.14159265358979

Public Sub ProcesarRegistrosTelcel()

    Dim wsOriginal As Worksheet
    Dim wsProcesado As Worksheet
    Dim headerRow As Long
    Dim lastRow As Long
    Dim mensajeError As String

    Set wsOriginal = ActiveSheet

    mensajeError = ValidacionPrevia(wsOriginal)
    If mensajeError <> "" Then
        MsgBox mensajeError, vbExclamation, "No se puede continuar"
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    ConfigurarEntornoRapido True

    Set wsProcesado = PrepararHojaProcesado(wsOriginal)
    headerRow = 1

    InsertarColumnasDeTrabajo wsProcesado, headerRow

    lastRow = wsProcesado.Cells(wsProcesado.Rows.Count, COL_NUMA).End(xlUp).Row
    If lastRow <= headerRow Then
        ConfigurarEntornoRapido False
        MsgBox "No quedaron registros despues de la limpieza inicial (¿el archivo estaba vacio o el " & _
               "formato de telefono no coincide con 10 digitos?).", vbExclamation
        Exit Sub
    End If

    EnriquecerDatos wsProcesado, headerRow, lastRow

    CrearHojaFrecuencias wsProcesado, headerRow, lastRow

    ConfigurarEntornoRapido False

    If MsgBox("¿Deseas generar el archivo KML con las antenas?", vbYesNo + vbQuestion, "KML") = vbYes Then
        GenerarKML wsProcesado.Parent.Worksheets(NOMBRE_HOJA_FRECUENCIAS), "P1"
    End If

    MsgBox "El procesamiento y la creacion de las tablas de resultados han finalizado exitosamente.", vbInformation
    Exit Sub

ErrorHandler:

    ConfigurarEntornoRapido False
    MsgBox "Ocurrio un error durante el procesamiento y la macro se detuvo de forma segura." & vbCrLf & vbCrLf & _
           "Descripcion: " & Err.Description & " (codigo " & Err.Number & ")" & vbCrLf & vbCrLf & _
           "La hoja '" & NOMBRE_HOJA_PROCESADO & "' puede haber quedado a medio transformar; " & _
           "se recomienda eliminarla manualmente antes de volver a correr la macro.", vbCritical
End Sub

Private Function ValidacionPrevia(ByVal wsOriginal As Worksheet) As String

    Dim hoja As Worksheet
    Dim headerRow As Long
    Dim lastColOriginal As Long

    If wsOriginal.Name = NOMBRE_HOJA_PROCESADO Or wsOriginal.Name = NOMBRE_HOJA_FRECUENCIAS Then
        ValidacionPrevia = "La hoja activa ('" & wsOriginal.Name & "') es una hoja de salida generada " & _
                            "por esta misma macro. Selecciona la hoja con los datos originales de Telcel " & _
                            "antes de ejecutar el proceso."
        Exit Function
    End If

    For Each hoja In wsOriginal.Parent.Worksheets
        If hoja.Name = NOMBRE_HOJA_PROCESADO Then
            If MsgBox("Ya existe una hoja '" & NOMBRE_HOJA_PROCESADO & "' (probablemente de una corrida " & _
                       "anterior). ¿Deseas eliminarla y continuar?", vbYesNo + vbQuestion, "Hoja existente") = vbYes Then
                Application.DisplayAlerts = False
                hoja.Delete
                Application.DisplayAlerts = True
            Else
                ValidacionPrevia = "Proceso cancelado por el usuario."
                Exit Function
            End If
        End If
    Next hoja

    headerRow = LocalizarFilaEncabezados(wsOriginal)
    If headerRow = 0 Then
        ValidacionPrevia = "No se encontro la fila de encabezados (se busco una fila donde la columna A " & _
                            "sea 'Telefono' y la columna B sea 'Tipo'). Verifica que sea un export de Telcel " & _
                            "sin modificar."
        Exit Function
    End If

    lastColOriginal = wsOriginal.Cells(headerRow, wsOriginal.Columns.Count).End(xlToLeft).Column
    If lastColOriginal < MIN_COLUMNAS_ORIGINAL Then
        ValidacionPrevia = "El archivo tiene solo " & lastColOriginal & " columnas en la fila de encabezados, " & _
                            "menos de las " & MIN_COLUMNAS_ORIGINAL & " esperadas para un export estandar de " & _
                            "Telcel. Es posible que el formato de exportacion haya cambiado; revisa el archivo " & _
                            "antes de continuar."
        Exit Function
    End If

    If wsOriginal.Cells(wsOriginal.Rows.Count, 1).End(xlUp).Row <= headerRow Then
        ValidacionPrevia = "No se detectaron filas de datos debajo del encabezado."
        Exit Function
    End If

    ValidacionPrevia = "" ' todo OK, se puede continuar
End Function

Private Function PrepararHojaProcesado(ByVal wsOriginal As Worksheet) As Worksheet

    Dim wsProcesado As Worksheet
    Dim headerRow As Long
    Dim i As Long
    Dim lastRow As Long
    Dim valorColA As String

    wsOriginal.Copy After:=wsOriginal
    Set wsProcesado = ActiveSheet
    wsProcesado.Name = NOMBRE_HOJA_PROCESADO   ' ya se valido antes que no hay colision de nombre

    headerRow = LocalizarFilaEncabezados(wsProcesado)   ' ya se valido antes que existe

    If headerRow > 1 Then
        wsProcesado.Rows("1:" & headerRow - 1).Delete
        headerRow = 1
    End If

    lastRow = wsProcesado.Cells(wsProcesado.Rows.Count, 1).End(xlUp).Row
    For i = lastRow To headerRow + 1 Step -1
        valorColA = Trim(CStr(wsProcesado.Cells(i, 1).Value))
        If Not (valorColA = "Telefono" Or (IsNumeric(valorColA) And Len(valorColA) = 10)) Then
            wsProcesado.Rows(i).Delete
        End If
    Next i

    Set PrepararHojaProcesado = wsProcesado
End Function

Private Function LocalizarFilaEncabezados(ByVal ws As Worksheet) As Long
    Dim i As Long
    For i = 1 To ws.UsedRange.Rows.Count
        If ws.Cells(i, 1).Value = "Telefono" And ws.Cells(i, 2).Value = "Tipo" Then
            LocalizarFilaEncabezados = i
            Exit Function
        End If
    Next i
    LocalizarFilaEncabezados = 0
End Function

Private Sub InsertarColumnasDeTrabajo(ByVal ws As Worksheet, ByVal headerRow As Long)

    ws.Columns("C:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove   ' SUBTIPO, SENTIDO, COLOR
    ws.Columns("H:L").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove   ' NUMBIFETEL..PROVEEDOR
    ws.Columns("N").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove     ' DIA
    ws.Columns("P:Q").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove   ' HORA 24, ETAPA
    ws.Columns("S").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove     ' MIN
    ws.Columns("X:Y").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove   ' BTS, BTS/AZIMUTH

    ws.Columns(COL_NUMERO_A).Copy Destination:=ws.Columns(COL_NUMERO_A_COPIA)
    ws.Columns(COL_NUMB).Copy Destination:=ws.Columns(COL_NUMB_COPIA)

    With ws
        .Cells(headerRow, COL_NUMA).Value = "NUMA"
        .Cells(headerRow, COL_TIPO).Value = "TIPO"
        .Cells(headerRow, COL_SUBTIPO).Value = "SUBTIPO"
        .Cells(headerRow, COL_SENTIDO).Value = "SENTIDO"
        .Cells(headerRow, COL_COLOR).Value = "COLOR"
        .Cells(headerRow, COL_NUMB).Value = "NUMB"
        .Cells(headerRow, COL_NUMBIFETEL).Value = "NUMBIFETEL"
        .Cells(headerRow, COL_MUNICIPIO).Value = "MUNICIPIO"
        .Cells(headerRow, COL_ESTADO).Value = "ESTADO"
        .Cells(headerRow, COL_RED).Value = "RED"
        .Cells(headerRow, COL_PROVEEDOR).Value = "PROVEEDOR DE SERVICIOS"
        .Cells(headerRow, COL_FECHA).Value = "FECHA"
        .Cells(headerRow, COL_DIA).Value = "DIA"
        .Cells(headerRow, COL_HORA).Value = "HORA"
        .Cells(headerRow, COL_HORA24).Value = "HORA 24"
        .Cells(headerRow, COL_ETAPA).Value = "ETAPA"
        .Cells(headerRow, COL_SEG).Value = "SEG"
        .Cells(headerRow, COL_MIN).Value = "MIN"
        .Cells(headerRow, COL_AZIMUTH).Value = "AZIMUTH"
        .Cells(headerRow, COL_BTS).Value = "BTS"
        .Cells(headerRow, COL_BTS_AZIMUTH).Value = "BTS / AZIMUTH"

        .Columns(COL_NUMERO_A).NumberFormat = "0"
        .Columns(COL_NUMB).NumberFormat = "0"
        .Columns(COL_HORA24).NumberFormat = "0"
        .Columns(COL_SEG).NumberFormat = "0.00"
        .Columns(COL_MIN).NumberFormat = "0.00"
        .Columns(COL_IMEI).NumberFormat = "0"
    End With
End Sub

Private Sub EnriquecerDatos(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long)

    Dim datos As Variant
    Dim i As Long
    Dim numF As String, numG As String
    Dim horaTexto As String
    Dim horaNumero As Variant
    Dim latitud As String, longitud As String

    datos = ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells(lastRow, ULTIMA_COLUMNA)).Value

    For i = 1 To UBound(datos, 1)

        ClasificarTipo CStr(datos(i, COL_TIPO)), datos(i, COL_SUBTIPO), datos(i, COL_SENTIDO), datos(i, COL_COLOR)

        numF = QuitarPrefijo52(Trim(CStr(datos(i, COL_NUMERO_A))))
        numG = QuitarPrefijo52(Trim(CStr(datos(i, COL_NUMB))))
        datos(i, COL_NUMERO_A) = numF

        If numG = Trim(CStr(datos(i, COL_NUMA))) Then
            numG = numF
        End If
        datos(i, COL_NUMB) = numG
        datos(i, COL_NUMBIFETEL) = Left(numG, 6)   ' primeros 6 digitos = prefijo IFT

        If IsDate(datos(i, COL_FECHA)) Then
            datos(i, COL_DIA) = Format(CDate(datos(i, COL_FECHA)), "dddd")
        Else
            datos(i, COL_DIA) = ""
        End If

        horaTexto = Trim(CStr(datos(i, COL_HORA)))
        horaNumero = ExtraerHora24(horaTexto)
        datos(i, COL_HORA24) = horaNumero
        datos(i, COL_ETAPA) = ObtenerEtapaDelDia(horaNumero)

        datos(i, COL_MIN) = Round(SafeToDouble(datos(i, COL_SEG)) / 60, 2)

        If CStr(datos(i, COL_IMEI)) = "null" Then
            datos(i, COL_IMEI) = ""
        End If

        latitud = Replace(CStr(datos(i, COL_LATITUD)), "N""", """N")
        longitud = Replace(CStr(datos(i, COL_LONGITUD)), "W""", """W")
        datos(i, COL_LATITUD) = latitud
        datos(i, COL_LONGITUD) = longitud

        If latitud <> "" And longitud <> "" Then
            datos(i, COL_BTS) = latitud & ", " & longitud
        Else
            datos(i, COL_BTS) = ""
        End If

        If datos(i, COL_BTS) <> "" And CStr(datos(i, COL_AZIMUTH)) <> "" Then
            datos(i, COL_BTS_AZIMUTH) = datos(i, COL_BTS) & " / " & datos(i, COL_AZIMUTH)
        Else
            datos(i, COL_BTS_AZIMUTH) = ""
        End If

    Next i

    ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells(lastRow, ULTIMA_COLUMNA)).Value = datos
End Sub

Private Sub ClasificarTipo(ByVal tipoValor As String, ByRef subtipo As Variant, ByRef sentido As Variant, ByRef color As Variant)
    Select Case tipoValor
        Case "DATOS"
            subtipo = SUB_DATOS: sentido = SENTIDO_SALIENTE: color = COLOR_AMARILLO
        Case "DATOS WiFi"
            subtipo = SUB_DATOS_WIFI: sentido = SENTIDO_SALIENTE: color = COLOR_AMARILLO
        Case "MENSAJE ENTRANTE"
            subtipo = SUB_MENSAJE_ENTRANTE: sentido = SENTIDO_ENTRANTE: color = COLOR_AZUL
        Case "MENSAJE SALIENTE"
            subtipo = SUB_MENSAJE_SALIENTE: sentido = SENTIDO_SALIENTE: color = COLOR_ROJO
        Case "VOZ ENTRANTE"
            subtipo = SUB_LLAMADA_ENTRANTE: sentido = SENTIDO_ENTRANTE: color = COLOR_AZUL
        Case "VOZ SALIENTE"
            subtipo = SUB_LLAMADA_SALIENTE: sentido = SENTIDO_SALIENTE: color = COLOR_ROJO
        Case "VOZ TRANSITO"
            subtipo = SUB_TRANSITO: sentido = "": color = ""
        Case "VOZ TRANSFER"
            subtipo = SUB_TRANSFER: sentido = "": color = ""
        Case Else
            subtipo = "": sentido = "": color = ""   ' tipo desconocido: se deja en blanco, no truena
    End Select
End Sub


Private Function QuitarPrefijo52(ByVal numero As String) As String
    If Left(numero, 2) = "52" And Len(numero) = 12 Then
        QuitarPrefijo52 = Mid(numero, 3)
    Else
        QuitarPrefijo52 = numero
    End If
End Function

Private Function SafeToDouble(ByVal valor As Variant) As Double
    On Error Resume Next
    SafeToDouble = Val(Trim(CStr(valor)))
    On Error GoTo 0
End Function

Private Function ExtraerHora24(ByVal horaTexto As String) As Variant
    Dim horaStr As String
    Dim partes() As String
    Dim horaNum As Integer

    If horaTexto = "" Then
        ExtraerHora24 = ""
        Exit Function
    End If

    If IsDate(horaTexto) Then
        ExtraerHora24 = Hour(CDate(horaTexto))
        Exit Function
    End If

    horaStr = LCase(horaTexto)
    horaStr = Replace(horaStr, "a. m.", "AM")
    horaStr = Replace(horaStr, "p. m.", "PM")
    horaStr = Replace(horaStr, "a.m.", "AM")
    horaStr = Replace(horaStr, "p.m.", "PM")
    horaStr = Replace(horaStr, "am", "AM")
    horaStr = Replace(horaStr, "pm", "PM")

    If IsDate(horaStr) Then
        ExtraerHora24 = Hour(CDate(horaStr))
        Exit Function
    End If

    partes = Split(horaStr, ":")
    If UBound(partes) >= 0 And IsNumeric(partes(0)) Then
        horaNum = CInt(Val(partes(0)))
        If InStr(horaStr, "PM") > 0 And horaNum < 12 Then
            horaNum = horaNum + 12
        ElseIf InStr(horaStr, "AM") > 0 And horaNum = 12 Then
            horaNum = 0
        End If
        If horaNum >= 0 And horaNum <= 23 Then
            ExtraerHora24 = horaNum
        Else
            ExtraerHora24 = ""
        End If
    Else
        ExtraerHora24 = ""
    End If
End Function


Private Function ObtenerEtapaDelDia(ByVal hora24 As Variant) As String
    If Not IsNumeric(hora24) Then
        ObtenerEtapaDelDia = ""
        Exit Function
    End If

    Select Case CInt(hora24)
        Case 0 To 5: ObtenerEtapaDelDia = "Madrugada"
        Case 6 To 11: ObtenerEtapaDelDia = "Mañana"
        Case 12 To 17: ObtenerEtapaDelDia = "Tarde"
        Case 18 To 23: ObtenerEtapaDelDia = "Noche"
        Case Else: ObtenerEtapaDelDia = ""
    End Select
End Function

Private Sub CrearHojaFrecuencias(ByVal wsProcesado As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long)

    Dim wsFrecuencias As Worksheet
    Dim lastCol As Long
    Dim dataRange As Range
    Dim errorComunicaciones As String
    Dim errorConexiones As String

    Application.DisplayAlerts = False
    On Error Resume Next
    Worksheets(NOMBRE_HOJA_FRECUENCIAS).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsFrecuencias = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    wsFrecuencias.Name = NOMBRE_HOJA_FRECUENCIAS

    lastCol = wsProcesado.Cells(headerRow, wsProcesado.Columns.Count).End(xlToLeft).Column
    Set dataRange = wsProcesado.Range(wsProcesado.Cells(headerRow, 1), wsProcesado.Cells(lastRow, lastCol))

    errorComunicaciones = IntentarEjecutarTarea("Comunicaciones", wsProcesado, headerRow, lastRow, wsFrecuencias, dataRange)
    errorConexiones = IntentarEjecutarTarea("Conexiones", wsProcesado, headerRow, lastRow, wsFrecuencias, dataRange)

    FormatearHojaFrecuencias wsFrecuencias

    If errorComunicaciones <> "" Or errorConexiones <> "" Then
        MsgBox "El procesamiento de datos termino bien, pero hubo problemas al construir " & _
               "alguna(s) tabla(s) de resultados:" & vbCrLf & vbCrLf & _
               IIf(errorComunicaciones <> "", "• Comunicaciones: " & errorComunicaciones & vbCrLf, "") & _
               IIf(errorConexiones <> "", "• Conexiones: " & errorConexiones & vbCrLf, ""), vbExclamation
    End If
End Sub

Private Function IntentarEjecutarTarea(ByVal nombre As String, ByVal wsProcesado As Worksheet, _
                                        ByVal headerRow As Long, ByVal lastRow As Long, _
                                        ByVal wsDestino As Worksheet, ByVal dataRange As Range) As String
    On Error GoTo Falla
    Select Case nombre
        Case "Comunicaciones"
            CrearPivotComunicaciones wsDestino, dataRange
        Case "Conexiones"
            GenerarTablaConexiones wsProcesado, headerRow, lastRow, wsDestino, "P1"
    End Select
    IntentarEjecutarTarea = ""
    Exit Function

Falla:
    IntentarEjecutarTarea = Err.Description
End Function

Private Sub CrearPivotComunicaciones(ByVal wsDestino As Worksheet, ByVal dataRange As Range)

    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim pi As PivotItem
    Dim subtiposVisibles As Variant

    subtiposVisibles = Array(SUB_LLAMADA_ENTRANTE, SUB_LLAMADA_SALIENTE, SUB_MENSAJE_ENTRANTE, SUB_MENSAJE_SALIENTE)

    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    Set pt = pc.CreatePivotTable(TableDestination:=wsDestino.Range("A1"), TableName:="TablaDinamicaComunicaciones")

    With pt
        .PivotFields("NUMB").Orientation = xlRowField
        .PivotFields("SUBTIPO").Orientation = xlColumnField
        With .PivotFields("TIPO")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "FRECUENCIA"
        End With
        
        .PivotFields("FECHA").Orientation = xlPageField
        .PivotFields("ETAPA").Orientation = xlPageField
        .PivotFields("DIA").Orientation = xlPageField
        .PivotFields("HORA").Orientation = xlPageField

        For Each pi In .PivotFields("SUBTIPO").PivotItems
            pi.Visible = EstaEnLista(pi.Name, subtiposVisibles)
        Next pi

        .RowAxisLayout xlTabularRow
        .PivotFields("NUMB").AutoSort xlDescending, "FRECUENCIA"
        .ColumnGrand = True
        .RowGrand = True
    End With
End Sub

Private Sub GenerarTablaConexiones(ByVal wsProcesado As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long, _
                                    ByVal wsDestino As Worksheet, ByVal celdaInicio As String)

    Dim subtipoArr As Variant, btsArr As Variant, azimuthArr As Variant
    Dim diccionario As Object   ' Scripting.Dictionary, enlace tardio (CreateObject):
                                ' asi no se necesita activar ninguna referencia en
                                ' Herramientas > Referencias para que la macro funcione
                                ' en otra PC.
    Dim subtiposValidos As Variant
    Dim i As Long
    Dim clave As String
    Dim coordenada As String, azimuth As String
    Dim claves() As String
    Dim frecuencias() As Long
    Dim totalUnicos As Long
    Dim idx As Long
    Dim k As Variant
    Dim partes() As String
    Dim celdaSalida As Range

    subtiposValidos = Array(SUB_LLAMADA_ENTRANTE, SUB_LLAMADA_SALIENTE, SUB_MENSAJE_ENTRANTE, _
                             SUB_MENSAJE_SALIENTE, SUB_DATOS, SUB_DATOS_WIFI)

    subtipoArr = wsProcesado.Range(wsProcesado.Cells(headerRow + 1, COL_SUBTIPO), wsProcesado.Cells(lastRow, COL_SUBTIPO)).Value
    btsArr = wsProcesado.Range(wsProcesado.Cells(headerRow + 1, COL_BTS), wsProcesado.Cells(lastRow, COL_BTS)).Value
    azimuthArr = wsProcesado.Range(wsProcesado.Cells(headerRow + 1, COL_AZIMUTH), wsProcesado.Cells(lastRow, COL_AZIMUTH)).Value

    Set diccionario = CreateObject("Scripting.Dictionary")

    For i = 1 To UBound(subtipoArr, 1)
        coordenada = Trim(CStr(btsArr(i, 1)))
        If coordenada <> "" And EstaEnLista(CStr(subtipoArr(i, 1)), subtiposValidos) Then
            azimuth = Trim(CStr(azimuthArr(i, 1)))
            clave = coordenada & "||" & azimuth
            If diccionario.Exists(clave) Then
                diccionario(clave) = diccionario(clave) + 1
            Else
                diccionario.Add clave, 1
            End If
        End If
    Next i

    totalUnicos = diccionario.Count
    If totalUnicos = 0 Then
        wsDestino.Range(celdaInicio).Value = "No se encontraron coordenadas validas para mostrar."
        Exit Sub
    End If

    ReDim claves(0 To totalUnicos - 1)
    ReDim frecuencias(0 To totalUnicos - 1)
    idx = 0
    For Each k In diccionario.Keys
        claves(idx) = CStr(k)
        frecuencias(idx) = diccionario(k)
        idx = idx + 1
    Next k

    OrdenarDescendentePorFrecuencia claves, frecuencias

    Set celdaSalida = wsDestino.Range(celdaInicio)
    celdaSalida.Offset(0, 0).Value = "COORDENADAS"
    celdaSalida.Offset(0, 1).Value = "AZIMUTH"
    celdaSalida.Offset(0, 2).Value = "FRECUENCIA"

    wsDestino.Range(celdaSalida.Offset(1, 0), celdaSalida.Offset(totalUnicos, 0)).NumberFormat = "@"
    wsDestino.Range(celdaSalida.Offset(1, 1), celdaSalida.Offset(totalUnicos, 1)).NumberFormat = "@"

    For idx = 0 To totalUnicos - 1
        partes = Split(claves(idx), "||")
        celdaSalida.Offset(idx + 1, 0).Value = partes(0)                                  ' Coordenadas
        celdaSalida.Offset(idx + 1, 1).Value = IIf(UBound(partes) >= 1, partes(1), "")     ' Azimuth
        celdaSalida.Offset(idx + 1, 2).Value = frecuencias(idx)                           ' Frecuencia
    Next idx
End Sub

Private Sub OrdenarDescendentePorFrecuencia(ByRef claves() As String, ByRef frecuencias() As Long)
    Dim i As Long, j As Long
    Dim claveTemp As String, frecTemp As Long

    For i = LBound(frecuencias) + 1 To UBound(frecuencias)
        claveTemp = claves(i)
        frecTemp = frecuencias(i)
        j = i - 1
        Do While j >= LBound(frecuencias)
            If frecuencias(j) >= frecTemp Then Exit Do   ' ya esta en su lugar
            frecuencias(j + 1) = frecuencias(j)
            claves(j + 1) = claves(j)
            j = j - 1
        Loop
        frecuencias(j + 1) = frecTemp
        claves(j + 1) = claveTemp
    Next i
End Sub

Private Function EstaEnLista(ByVal valor As String, ByVal lista As Variant) As Boolean
    Dim elemento As Variant
    For Each elemento In lista
        If valor = elemento Then
            EstaEnLista = True
            Exit Function
        End If
    Next elemento
    EstaEnLista = False
End Function

Private Sub ConfigurarEntornoRapido(ByVal activar As Boolean)
    If activar Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
        Application.DisplayStatusBar = False
    Else
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayStatusBar = True
    End If
End Sub

Private Sub GenerarKML(ByVal wsFrecuencias As Worksheet, ByVal celdaInicio As String)

    Dim ini As Range
    Dim lastRow As Long
    Dim datos As Variant
    Dim i As Long
    Dim partes() As String
    Dim lat As Double, lon As Double
    Dim okLat As Boolean, okLon As Boolean
    Dim freq As Long, maxFreq As Long
    Dim azText As String
    Dim kmlAntenas As String, kmlSectores As String
    Dim escala As Double
    Dim omitidas As Long, escritas As Long
    Dim rutaKml As Variant
    Dim kml As String

    On Error GoTo Falla

    Set ini = wsFrecuencias.Range(celdaInicio)
    lastRow = wsFrecuencias.Cells(wsFrecuencias.Rows.Count, ini.Column).End(xlUp).Row
    If lastRow <= ini.Row Or Trim(CStr(ini.Offset(0, 1).Value)) <> "AZIMUTH" Then
        MsgBox "No hay tabla de conexiones para convertir a KML.", vbExclamation
        Exit Sub
    End If

    datos = wsFrecuencias.Range(ini.Offset(1, 0), wsFrecuencias.Cells(lastRow, ini.Column + 2)).Value

    For i = 1 To UBound(datos, 1)
        If IsNumeric(datos(i, 3)) Then
            If CLng(datos(i, 3)) > maxFreq Then maxFreq = CLng(datos(i, 3))
        End If
    Next i
    If maxFreq = 0 Then maxFreq = 1

    For i = 1 To UBound(datos, 1)
        partes = Split(CStr(datos(i, 1)), ", ")   ' "Latitud, Longitud"
        If UBound(partes) <> 1 Then
            omitidas = omitidas + 1
        Else
            lat = DMSaDecimal(partes(0), False, okLat)
            lon = DMSaDecimal(partes(1), True, okLon)
            If Not (okLat And okLon) Then
                omitidas = omitidas + 1
            Else
                freq = CLng(Val(CStr(datos(i, 3))))
                azText = Trim(CStr(datos(i, 2)))
                escala = 0.7 + 1.3 * (freq / maxFreq)

                kmlAntenas = kmlAntenas & _
                    "<Placemark><name>" & EscaparXML("Antena " & (escritas + 1) & " (" & freq & ")") & "</name>" & _
                    "<description>" & EscaparXML("Frecuencia: " & freq & " | Azimuth: " & azText & _
                                                 " | Coordenadas: " & CStr(datos(i, 1))) & "</description>" & _
                    "<Style><IconStyle><scale>" & FormatoCoord(escala) & "</scale></IconStyle></Style>" & _
                    "<Point><coordinates>" & FormatoCoord(lon) & "," & FormatoCoord(lat) & ",0</coordinates></Point>" & _
                    "</Placemark>" & vbCrLf

                If DIBUJAR_SECTORES And IsNumeric(azText) And azText <> "" Then
                    kmlSectores = kmlSectores & ConstruirSectorKML(lat, lon, CDbl(Val(azText)), freq, escritas + 1)
                End If

                escritas = escritas + 1
            End If
        End If
    Next i

    If escritas = 0 Then
        MsgBox "No se pudo interpretar ninguna coordenada; no se genero el KML.", vbExclamation
        Exit Sub
    End If

    kml = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
          "<kml xmlns=""http://www.opengis.net/kml/2.2""><Document>" & vbCrLf & _
          "<name>Antenas CDR</name>" & vbCrLf & _
          "<Folder><name>Antenas</name>" & vbCrLf & kmlAntenas & "</Folder>" & vbCrLf
    If kmlSectores <> "" Then
        kml = kml & "<Folder><name>Sectores</name>" & vbCrLf & kmlSectores & "</Folder>" & vbCrLf
    End If
    kml = kml & "</Document></kml>"

    Dim carpetaBase As String
    carpetaBase = wsFrecuencias.Parent.Path
    If carpetaBase = "" Or Left(LCase(carpetaBase), 4) = "http" Then
        carpetaBase = Environ("USERPROFILE") & "\Documents"
    End If

    rutaKml = Application.GetSaveAsFilename( _
                InitialFileName:=carpetaBase & "\Antenas_CDR.kml", _
                FileFilter:="Archivo KML (*.kml), *.kml", Title:="Guardar KML")
    If VarType(rutaKml) = vbBoolean Then Exit Sub   ' el usuario cancelo

    If LCase(Right(CStr(rutaKml), 4)) <> ".kml" Then rutaKml = rutaKml & ".kml"

    If Dir(CStr(rutaKml)) <> "" Then
        If Not ArchivoDisponible(CStr(rutaKml)) Then
            MsgBox "El archivo esta abierto en otro programa (¿Google Earth?). Cierralo e intenta de nuevo.", vbExclamation
            Exit Sub
        End If
    End If

    EscribirArchivoUTF8 CStr(rutaKml), kml

    MsgBox "KML generado con " & escritas & " antenas." & _
           IIf(omitidas > 0, vbCrLf & omitidas & " fila(s) omitida(s) por coordenadas no interpretables.", "") & _
           IIf(escritas > 2000, vbCrLf & vbCrLf & "Nota: Google My Maps limita cada capa a 2,000 elementos.", ""), _
           vbInformation
    Exit Sub

Falla:
    MsgBox "No se pudo generar el KML: " & Err.Description & " (codigo " & Err.Number & ")" & _
           vbCrLf & vbCrLf & "Ruta: " & CStr(rutaKml), vbExclamation
End Sub

Private Function DMSaDecimal(ByVal texto As String, ByVal esLongitud As Boolean, ByRef ok As Boolean) As Double
    Dim i As Long, c As String, token As String, hemi As String
    Dim vals(0 To 2) As Double
    Dim n As Long, resultado As Double

    ok = False
    texto = Trim(texto)
    If texto = "" Then Exit Function

    For i = 1 To Len(texto) + 1
        If i <= Len(texto) Then c = Mid(texto, i, 1) Else c = " "
        If c Like "[0-9.]" Then
            token = token & c
        Else
            If token <> "" Then
                If n <= 2 Then vals(n) = Val(token)
                n = n + 1
                token = ""
            End If
            If UCase(c) Like "[NSEWO]" Then hemi = UCase(c)
        End If
    Next i

    If n = 0 Then Exit Function
    resultado = vals(0) + vals(1) / 60 + vals(2) / 3600

    If hemi = "S" Or hemi = "W" Or hemi = "O" Then
        resultado = -resultado
    ElseIf hemi = "" And esLongitud Then
        resultado = -resultado
    End If

    If (esLongitud And Abs(resultado) > 180) Or (Not esLongitud And Abs(resultado) > 90) Then Exit Function

    DMSaDecimal = resultado
    ok = True
End Function

Private Function ConstruirSectorKML(ByVal lat As Double, ByVal lon As Double, ByVal azimuth As Double, _
                                     ByVal freq As Long, ByVal numero As Long) As String
    Dim coords As String
    Dim ang As Double, desde As Double, hasta As Double
    Dim latP As Double, lonP As Double

    desde = azimuth - ANCHO_SECTOR_GRADOS / 2
    hasta = azimuth + ANCHO_SECTOR_GRADOS / 2

    coords = FormatoCoord(lon) & "," & FormatoCoord(lat) & ",0 "
    ang = desde
    Do While ang < hasta
        PuntoDestino lat, lon, ang, RADIO_SECTOR_M, latP, lonP
        coords = coords & FormatoCoord(lonP) & "," & FormatoCoord(latP) & ",0 "
        ang = ang + 10
    Loop
    PuntoDestino lat, lon, hasta, RADIO_SECTOR_M, latP, lonP
    coords = coords & FormatoCoord(lonP) & "," & FormatoCoord(latP) & ",0 "
    coords = coords & FormatoCoord(lon) & "," & FormatoCoord(lat) & ",0"

    ConstruirSectorKML = "<Placemark><name>" & EscaparXML("Sector " & numero & " (" & freq & ")") & "</name>" & _
        "<Style><LineStyle><color>ff00a5ff</color><width>1</width></LineStyle>" & _
        "<PolyStyle><color>5500a5ff</color></PolyStyle></Style>" & _
        "<Polygon><outerBoundaryIs><LinearRing><coordinates>" & coords & _
        "</coordinates></LinearRing></outerBoundaryIs></Polygon></Placemark>" & vbCrLf
End Function

Private Sub PuntoDestino(ByVal lat As Double, ByVal lon As Double, ByVal rumboGrados As Double, _
                         ByVal distM As Double, ByRef latOut As Double, ByRef lonOut As Double)
    Const R As Double = 6371000
    Dim lat1 As Double, lon1 As Double, brng As Double, d As Double
    Dim lat2 As Double, lon2 As Double

    lat1 = lat * PI_VALOR / 180
    lon1 = lon * PI_VALOR / 180
    brng = rumboGrados * PI_VALOR / 180
    d = distM / R

    lat2 = Asen(Sin(lat1) * Cos(d) + Cos(lat1) * Sin(d) * Cos(brng))
    lon2 = lon1 + Atn2(Sin(brng) * Sin(d) * Cos(lat1), Cos(d) - Sin(lat1) * Sin(lat2))

    latOut = lat2 * 180 / PI_VALOR
    lonOut = lon2 * 180 / PI_VALOR
End Sub

Private Function Asen(ByVal x As Double) As Double
    If x >= 1 Then
        Asen = PI_VALOR / 2
    ElseIf x <= -1 Then
        Asen = -PI_VALOR / 2
    Else
        Asen = Atn(x / Sqr(1 - x * x))
    End If
End Function

Private Function Atn2(ByVal y As Double, ByVal x As Double) As Double
    If x > 0 Then
        Atn2 = Atn(y / x)
    ElseIf x < 0 Then
        Atn2 = Atn(y / x) + IIf(y >= 0, PI_VALOR, -PI_VALOR)
    Else
        Atn2 = IIf(y > 0, PI_VALOR / 2, IIf(y < 0, -PI_VALOR / 2, 0))
    End If
End Function

Private Function FormatoCoord(ByVal valor As Double) As String
    FormatoCoord = Replace(Format(valor, "0.000000"), ",", ".")
End Function


Private Function EscaparXML(ByVal texto As String) As String
    texto = Replace(texto, "&", "&amp;")
    texto = Replace(texto, "<", "&lt;")
    texto = Replace(texto, ">", "&gt;")
    texto = Replace(texto, """", "&quot;")
    texto = Replace(texto, "'", "&apos;")
    EscaparXML = texto
End Function

Private Sub EscribirArchivoUTF8(ByVal ruta As String, ByVal contenido As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2            ' adTypeText
        .Charset = "utf-8"
        .Open
        .WriteText contenido
        .SaveToFile ruta, 2  ' adSaveCreateOverWrite
        .Close
    End With
End Sub

Private Function ArchivoDisponible(ByVal ruta As String) As Boolean
    Dim f As Integer
    On Error GoTo NoDisponible
    f = FreeFile
    Open ruta For Append Access Write Lock Read Write As #f
    Close #f
    ArchivoDisponible = True
    Exit Function
NoDisponible:
    ArchivoDisponible = False
End Function

Private Sub FormatearHojaFrecuencias(ByVal ws As Worksheet)

    Dim pt As PivotTable
    Dim hdr As Range, cuerpo As Range
    Dim ultimaFila As Long
    Dim colorEncabezado As Long

    colorEncabezado = RGB(31, 78, 121)   ' azul oscuro

    On Error Resume Next   ' formato opcional: no debe romper el proceso

    Set pt = ws.PivotTables("TablaDinamicaComunicaciones")
    If Not pt Is Nothing Then
        pt.TableStyle2 = "PivotStyleMedium2"
        pt.ShowTableStyleRowStripes = True
        pt.TableRange1.Borders.LineStyle = xlContinuous
        pt.TableRange1.Borders.Color = RGB(191, 191, 191)
        pt.DataBodyRange.HorizontalAlignment = xlCenter
    End If

    ultimaFila = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row
    If ultimaFila >= 1 Then
        Set hdr = ws.Range("P1:R1")
        With hdr
            .Interior.Color = colorEncabezado
            .Font.Color = vbWhite
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With

        If ultimaFila > 1 Then
            Set cuerpo = ws.Range("P2:R" & ultimaFila)
            With cuerpo
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(191, 191, 191)
            End With
            ws.Range("Q2:R" & ultimaFila).HorizontalAlignment = xlCenter
        End If
        hdr.Borders.LineStyle = xlContinuous
        hdr.Borders.Color = RGB(191, 191, 191)
    End If

    ws.Columns.AutoFit

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1:R35").Select
    ActiveWindow.Zoom = True                 ' ajusta el zoom a la seleccion
    If ActiveWindow.Zoom > 120 Then ActiveWindow.Zoom = 120
    If ActiveWindow.Zoom < 60 Then ActiveWindow.Zoom = 60
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ws.Range("A1").Select

    On Error GoTo 0
End Sub
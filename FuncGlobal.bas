Attribute VB_Name = "FuncGlobal"
Option Explicit

Public Enum gsDataTypes
    gsdtString
    gsdtNumber
    gsdtBoolean
    gsdtDateTime
End Enum

Public Enum gsDataTypesTransform
    gsdttNone
    gsdttZeroToNull
End Enum



Public Function DayOfWeek(Fecha As Variant) As String
    If IsDate(Fecha) Then
        Select Case Weekday(Fecha)
            Case vbSunday
                DayOfWeek = "Domingo"
            Case vbMonday
                DayOfWeek = "Lunes"
            Case vbTuesday
                DayOfWeek = "Martes"
            Case vbWednesday
                DayOfWeek = "Miércoles"
            Case vbThursday
                DayOfWeek = "Jueves"
            Case vbFriday
                DayOfWeek = "Viernes"
            Case vbSaturday
                DayOfWeek = "Sábado"
        End Select
    End If
End Function


Function ConvertToPeriodComma(Expression As Variant, Optional IncludeGroupSeparator As Boolean = False) As Variant
    Dim lngPeriodPosition As Long
    Dim lngCommaPosition As Long
    
    If Not IsNumeric(Expression) Then
        ConvertToPeriodComma = 0
        Exit Function
    End If
    
    lngPeriodPosition = InStr(1, Expression, ".")
    lngCommaPosition = InStr(1, Expression, ",")
    
    If lngPeriodPosition > 0 And lngCommaPosition > 0 Then
        'Encontré ambos símbolos
        If lngPeriodPosition > lngCommaPosition Then
            'Está con Coma como separador de Miles y Punto como separador decimal
            Do While lngCommaPosition > 0
                Expression = Left(Expression, lngCommaPosition - 1) + IIf(IncludeGroupSeparator, ".", "") + Right(Expression, Len(Expression) - lngCommaPosition)
                lngCommaPosition = InStr(1, Expression, ",")
            Loop
            lngPeriodPosition = InStr(1, Expression, ".")
            Expression = Left(Expression, lngPeriodPosition - 1) + "," + Right(Expression, Len(Expression) - lngPeriodPosition)
        Else
            'Está con Punto como separador de Miles y Coma como separador decimal,
            'por lo tanto, lo dejo como está.
            If Not IncludeGroupSeparator Then
                Do While lngPeriodPosition > 0
                    Expression = Left(Expression, lngPeriodPosition - 1) + Right(Expression, Len(Expression) - lngPeriodPosition)
                    lngPeriodPosition = InStr(1, Expression, ".")
                Loop
            End If
        End If
    Else
        If lngPeriodPosition > 0 Or lngCommaPosition > 0 Then
            If lngPeriodPosition = 0 Then
                'Tiene solamente coma(s)
                'Verifico si tiene otra comma para ver si no es un separador de miles
                If InStr(lngCommaPosition + 1, Expression, ",") > 0 Then
                    'Es un separador de miles
                    Do While lngCommaPosition > 0
                        Expression = Left(Expression, lngCommaPosition - 1) + IIf(IncludeGroupSeparator, ".", "") + Right(Expression, Len(Expression) - lngCommaPosition)
                        lngCommaPosition = InStr(1, Expression, ",")
                    Loop
                Else
                    'Verifico si tiene 3 caracteres después de la comma
                    If Len(Mid(Expression, lngCommaPosition + 1, 3)) = 3 Then
                        'Supongo que es un Separador de miles, aunque nada me asegura que sea así
                        Expression = Left(Expression, lngCommaPosition - 1) + IIf(IncludeGroupSeparator, ".", "") + Right(Expression, Len(Expression) - lngCommaPosition)
                    Else
                        'Supongo que es un separador decimal, por lo tanto, lo dejo como esta
                    End If
                End If
            Else
                If lngCommaPosition = 0 Then
                    'Tiene solamente punto(s)
                    'Verifico si tiene otro punto para ver si no es un separador de miles
                    If InStr(lngPeriodPosition + 1, Expression, ".") > 0 Then
                        'Es un separador de miles, por lo tanto lo dejo como está
                        If Not IncludeGroupSeparator Then
                            Do While lngPeriodPosition > 0
                                Expression = Expression = Left(Expression, lngPeriodPosition - 1) + Right(Expression, Len(Expression) - lngPeriodPosition)
                                lngPeriodPosition = InStr(1, Expression, ".")
                            Loop
                        End If
                    Else
                        'Verifico si tiene 3 caracteres después del punto
                        If Len(Mid(Expression, lngPeriodPosition + 1, Len(Expression))) = 3 Then
                            'Supongo que es un Separador de miles, aunque nada me asegura que sea así
                            'por lo tanto lo dejo como está
                            If Not IncludeGroupSeparator Then Expression = Left(Expression, lngPeriodPosition - 1) + Right(Expression, Len(Expression) - lngPeriodPosition)
                        Else
                            'Supongo que es un separador decimal
                            Expression = Left(Expression, lngPeriodPosition - 1) + "," + Right(Expression, Len(Expression) - lngPeriodPosition)
                        End If
                    End If
                End If
            End If
        End If
    End If
    ConvertToPeriodComma = Expression
End Function

Function ConvertToCommaPeriod(Expression As Variant, Optional IncludeGroupSeparator As Boolean = False) As Variant
    Dim lngPeriodPosition As Long
    Dim lngCommaPosition As Long
    
    If Not IsNumeric(Expression) Then
        ConvertToCommaPeriod = 0
        Exit Function
    End If
    
    lngCommaPosition = InStr(1, Expression, ",")
    lngPeriodPosition = InStr(1, Expression, ".")
    
    If lngCommaPosition > 0 And lngPeriodPosition > 0 Then
        'Encontré ambos símbolos
        If lngCommaPosition > lngPeriodPosition Then
            'Está con Punto como separador de Miles y Coma como separador decimal
            Do While lngPeriodPosition > 0
                Expression = Left(Expression, lngPeriodPosition - 1) + IIf(IncludeGroupSeparator, ",", "") + Right(Expression, Len(Expression) - lngPeriodPosition)
                lngPeriodPosition = InStr(1, Expression, ".")
            Loop
            lngCommaPosition = InStr(1, Expression, ",")
            Expression = Left(Expression, lngCommaPosition - 1) + "." + Right(Expression, Len(Expression) - lngCommaPosition)
        Else
            'Está con Coma como separador de Miles y Punto como separador decimal,
            'por lo tanto, lo dejo como está.
            If Not IncludeGroupSeparator Then
                Do While lngCommaPosition > 0
                    Expression = Left(Expression, lngCommaPosition - 1) + Right(Expression, Len(Expression) - lngCommaPosition)
                    lngCommaPosition = InStr(1, Expression, ",")
                Loop
            End If
        End If
    Else
        If lngCommaPosition > 0 Or lngPeriodPosition > 0 Then
            If lngCommaPosition = 0 Then
                'Tiene solamente punto(s)
                'Verifico si tiene otro punto para ver si no es un separador de miles
                If InStr(lngPeriodPosition + 1, Expression, ".") > 0 Then
                    'Es un separador de miles
                    Do While lngPeriodPosition > 0
                        Expression = Left(Expression, lngPeriodPosition - 1) + IIf(IncludeGroupSeparator, ",", "") + Right(Expression, Len(Expression) - lngPeriodPosition)
                        lngPeriodPosition = InStr(1, Expression, ".")
                    Loop
                Else
                    'Verifico si tiene 3 caracteres después del punto
                    If Len(Mid(Expression, lngPeriodPosition + 1, 3)) = 3 Then
                        'Supongo que es un Separador de miles, aunque nada me asegura que sea así
                        Expression = Left(Expression, lngPeriodPosition - 1) + IIf(IncludeGroupSeparator, ",", "") + Right(Expression, Len(Expression) - lngPeriodPosition)
                    Else
                        'Supongo que es un separador decimal, por lo tanto, lo dejo como esta
                    End If
                End If
            Else
                If lngPeriodPosition = 0 Then
                    'Tiene solamente coma(s)
                    'Verifico si tiene otra coma para ver si no es un separador de miles
                    If InStr(lngCommaPosition + 1, Expression, ",") > 0 Then
                        'Es un separador de miles, por lo tanto lo dejo como está
                        If Not IncludeGroupSeparator Then
                            Do While lngCommaPosition > 0
                                Expression = Expression = Left(Expression, lngCommaPosition - 1) + Right(Expression, Len(Expression) - lngCommaPosition)
                                lngCommaPosition = InStr(1, Expression, ",")
                            Loop
                        End If
                    Else
                        'Verifico si tiene 3 caracteres después de la coma
                        If Len(Mid(Expression, lngCommaPosition + 1, Len(Expression))) = 3 Then
                            'Supongo que es un Separador de miles, aunque nada me asegura que sea así
                            'por lo tanto lo dejo como está
                            If Not IncludeGroupSeparator Then Expression = Left(Expression, lngCommaPosition - 1) + Right(Expression, Len(Expression) - lngCommaPosition)
                        Else
                            'Supongo que es un separador decimal
                            Expression = Left(Expression, lngCommaPosition - 1) + "." + Right(Expression, Len(Expression) - lngCommaPosition)
                        End If
                    End If
                End If
            End If
        End If
    End If
    ConvertToCommaPeriod = Expression
End Function



Function MaxDayOfMonth(DateToCalculate As Variant) As Byte
    If Not IsDate(DateToCalculate) Then
        MaxDayOfMonth = 0
    Else
        Select Case Month(DateToCalculate)
            Case 1, 3, 5, 7, 8, 10, 12
                'Enero, Marzo, Mayo, Julio, Agosto, Octubre, Diciembre
                MaxDayOfMonth = 31
            Case 4, 6, 9, 11
                'Abril, Junio, Septiembre, Noviembre
                MaxDayOfMonth = 30
            Case 2
                'Febrero...calcular bisiestos
                If (Year(DateToCalculate) Mod 4) = 0 Then
                    'Es divisible por 4
                    If (Year(DateToCalculate) Mod 100) = 0 Then
                        'Es divisible por 100
                        If (Year(DateToCalculate) Mod 400) = 0 Then
                            'Es divisible por 400
                            MaxDayOfMonth = 29
                        Else
                            'No es divisible por 400
                            MaxDayOfMonth = 28
                        End If
                    Else
                        'No es divisible por 100
                        MaxDayOfMonth = 29
                    End If
                Else
                    'No es divisible por 4
                    MaxDayOfMonth = 28
                End If
        End Select
    End If
End Function


Public Function FillComboBoxHora(ComboBox As ComboBox, StartHour As Boolean, StartTime As String, EndTime As String, Interval As Byte)
    Dim intMinutos As Integer
    Dim intStart_Hour_Minutes As Integer
    Dim intEnd_Hour_Minutes As Integer
    
    intStart_Hour_Minutes = DateDiff("n", "00:00", StartTime)
    intEnd_Hour_Minutes = DateDiff("n", "00:00", EndTime)
    
    If StartHour Then
        'Es un combo de Horas de Inicio
        For intMinutos = intStart_Hour_Minutes To intEnd_Hour_Minutes Step Interval
            ComboBox.AddItem Format(DateAdd("n", intMinutos, "00:00"), "hh:nn")
        Next intMinutos
    Else
        'Es un combo de Horas de Fin
        For intMinutos = Interval - 1 To intEnd_Hour_Minutes Step Interval
            ComboBox.AddItem Format(DateAdd("n", intMinutos, "00:00"), "hh:nn")
        Next intMinutos
    End If
End Function



Public Function SetComboBoxListIndex(ComboBoxToSet As ComboBox, ItemDataToFind) As Boolean
    Dim ComboBoxListIndex As Long
    
    For ComboBoxListIndex = 0 To ComboBoxToSet.ListCount - 1
        If ComboBoxToSet.ItemData(ComboBoxListIndex) = ItemDataToFind Then
            ComboBoxToSet.ListIndex = ComboBoxListIndex
            SetComboBoxListIndex = True
            Exit Function
        End If
    Next ComboBoxListIndex
    SetComboBoxListIndex = False
End Function



Public Function GetErrorMessage() As String
    GetErrorMessage = "Error #" & Err.Number & ": " & Err.Description
End Function

Public Function GetADODBErrorMessage(ADODBErrors As ADODB.Errors) As String
    Dim ADODBError As ADODB.Error
    
    For Each ADODBError In ADODBErrors
        If ADODBError.NativeError <> 3621 Then
            GetADODBErrorMessage = GetADODBErrorMessage & "Error #" & ADODBError.NativeError & ": " & ADODBError.Description & vbCr
        End If
    Next ADODBError
    If GetADODBErrorMessage = "" Then
        GetADODBErrorMessage = "Este no es un error de ADO." & vbCr & GetErrorMessage
    Else
        If (InStr(1, UCase(GetADODBErrorMessage), "TIMEOUT") > 0) Or (InStr(1, UCase(GetADODBErrorMessage), "TIME OUT") > 0) Then
            GetADODBErrorMessage = "Se ha superado el tiempo de espera, por favor reintente nuevamente en unos segundos."
        Else
            GetADODBErrorMessage = Left(GetADODBErrorMessage, Len(GetADODBErrorMessage) - 1)
        End If
        
    End If
End Function

Public Sub ShowCustomADODBErrorMessage(ADODBErrors As ADODB.Errors, CustomMessage As String, Optional NativeErrorNumberToTrap As Long = 0, Optional MessageToTrappedError As String = "")
    Dim ADODBError As ADODB.Error
    Dim ErrorMessage As String
    
    For Each ADODBError In ADODBErrors
        If ADODBError.NativeError <> 3621 Then
            If ADODBError.NativeError = NativeErrorNumberToTrap And NativeErrorNumberToTrap <> 0 Then
                MsgBox MessageToTrappedError, vbInformation, App.Title
                Exit Sub
            Else
                ErrorMessage = ErrorMessage & "Error #" & ADODBError.NativeError & ": " & ADODBError.Description & vbCr
            End If
        End If
    Next ADODBError
    If ErrorMessage = "" Then
        ErrorMessage = "Este no es un error de ADO." & vbCr & GetErrorMessage
    Else
        ErrorMessage = Left(ErrorMessage, Len(ErrorMessage) - 1)
    End If
    If Trim(CustomMessage) <> "" Then ErrorMessage = CustomMessage & vbCr & vbCr & ErrorMessage
    MsgBox ErrorMessage, vbCritical, App.Title
End Sub

Public Function FormatToSQL(Value As Variant, DataType As gsDataTypes, Optional Transformation As gsDataTypesTransform = gsdttZeroToNull) As String
    Select Case DataType
        Case gsdtString
            If IsEmpty(Value) Or Trim(Value) = "" Then
                FormatToSQL = "NULL"
            Else
                FormatToSQL = "'" + ReplaceQuote(CStr(Value)) + "'"
            End If
        Case gsdtNumber
            If IsEmpty(Value) Or (Not IsNumeric(Value)) Or (Value = 0 And Transformation = gsdttZeroToNull) Then
                FormatToSQL = "NULL"
            Else
                FormatToSQL = ConvertToCommaPeriod(Value)
            End If
        Case gsdtDateTime
            If IsEmpty(Value) Or (Not IsDate(Value)) Then
                FormatToSQL = "NULL"
            Else
                FormatToSQL = "'" + Format(Value, "yyyy/mm/dd hh:nn") + "'"
            End If
        Case gsdtBoolean
            If Value Then
                FormatToSQL = "1"
            Else
                FormatToSQL = "0"
            End If
    End Select
End Function

Public Function ReplaceQuote(StringToReplace As String) As String
    Dim QuotePosition As Long

    ReplaceQuote = StringToReplace
    QuotePosition = InStr(1, StringToReplace, "'")
    Do While QuotePosition > 0
        ReplaceQuote = Left(ReplaceQuote, QuotePosition) & "'" & Right(ReplaceQuote, (Len(ReplaceQuote) - QuotePosition))
        QuotePosition = InStr(QuotePosition + 2, ReplaceQuote, "'")
    Loop
End Function


Public Sub SelAllText(Control As Control)
    On Error Resume Next
    
    Control.SelStart = 0
    Control.SelLength = Len(Control)
End Sub


Function CalcularEdad(FechaNacimiento As Date, FechaCalculo As Date, Optional MaxYearToCalculateMonth As Long = 5, Optional MaxYearToCalculateDay As Long = 5) As String
    Dim lngYearsElapsedTemp As Long
    Dim dtCompleteYear As Date
    Dim lngEdadAnios As Long
    Dim strYearsString As String
    
    Dim lngMonthsElapsedTemp As Long
    Dim dtCompleteMonth As Date
    Dim lngEdadMeses As Long
    Dim strMonthsString As String
    
    Dim lngEdadDias As Long
    Dim strDaysString As String
    
    '=========================== AÑOS ==================================================
    'Calculo la cantidad de años transcurridos
    lngYearsElapsedTemp = DateDiff("yyyy", FechaNacimiento, FechaCalculo)
    
    'A la fecha de hoy le resto la cantidad de años que calculé anteriormente,
    'esto lo hago para corregir el error del VB
    dtCompleteYear = DateAdd("yyyy", -lngYearsElapsedTemp, FechaCalculo)
    
    'Si me pasé del límite, le resto un año
    If dtCompleteYear < FechaNacimiento Then dtCompleteYear = DateAdd("yyyy", 1, dtCompleteYear)
            
    'Calculo los años reales
    lngEdadAnios = DateDiff("yyyy", dtCompleteYear, FechaCalculo)
    
    Select Case lngEdadAnios
        Case 0
        Case 1
            strYearsString = "1 año"
        Case Else
            strYearsString = Format(lngEdadAnios) + " años"
    End Select
    '===================================================================================
    
    '=========================== MESES =================================================
    'Calculo la cantidad de meses transcurridos desde el último Año Completo
    lngMonthsElapsedTemp = DateDiff("m", FechaNacimiento, dtCompleteYear)
    
    'A la fecha de hoy le resto la cantidad de años que calculé anteriormente,
    'esto lo hago para corregir el error del VB
    dtCompleteMonth = DateAdd("m", -lngMonthsElapsedTemp, dtCompleteYear)
    
    'Si me pasé del límite, le resto un año
    If dtCompleteMonth < FechaNacimiento Then dtCompleteMonth = DateAdd("m", 1, dtCompleteMonth)
            
    'Calculo los años reales
    lngEdadMeses = DateDiff("m", dtCompleteMonth, dtCompleteYear)
    
    Select Case lngEdadMeses
        Case 0
        Case 1
            strMonthsString = "1 mes"
        Case Else
            strMonthsString = Format(lngEdadMeses) + " meses"
    End Select
    '===================================================================================
    
    '=========================== DIAS ==================================================
    'Calculo los días restantes
    lngEdadDias = DateDiff("d", FechaNacimiento, dtCompleteMonth)
    
    Select Case lngEdadDias
        Case 0
        Case 1
            strDaysString = "1 día"
        Case Else
            strDaysString = Format(lngEdadDias) + " días"
    End Select
    '===================================================================================
    
    If lngEdadAnios > MaxYearToCalculateMonth And MaxYearToCalculateMonth > 0 Then strMonthsString = ""
    If lngEdadAnios > MaxYearToCalculateDay And MaxYearToCalculateDay > 0 Then strDaysString = ""
    
    'Armo el string final
    If strYearsString <> "" And strMonthsString <> "" And strDaysString <> "" Then
        CalcularEdad = strYearsString + ", " + strMonthsString + " y " + strDaysString
    Else
        If strYearsString <> "" And strMonthsString <> "" Then
            CalcularEdad = strYearsString + " y " + strMonthsString
        Else
            If strYearsString <> "" And strDaysString <> "" Then
                CalcularEdad = strYearsString + " y " + strDaysString
            Else
                If strMonthsString <> "" And strDaysString <> "" Then
                    CalcularEdad = strMonthsString + " y " + strDaysString
                Else
                    CalcularEdad = strYearsString + strMonthsString + strDaysString
                End If
            End If
        End If
    End If
End Function

Public Function ExistField(ByRef rec As ADODB.Recordset, strField As String) As Boolean
Dim CurrentField As ADODB.Field
    
    ExistField = False
    For Each CurrentField In rec.Fields
        If UCase(CurrentField.Name) = UCase(strField) Then
            ExistField = True
            Exit Function
        End If
    Next CurrentField
    
End Function

Function ElapsedTime(StartTime As Variant, Optional EndTime As Variant) As String
Dim intSegundos As Long
Dim intMinutos As Long
Dim intHoras As Long
Dim m_StartTime As Variant
Dim m_EndTime As Variant
    
    If IsMissing(EndTime) Then
        m_EndTime = Format(Now, "yyyy/mm/dd hh:nn:ss")
    Else
        m_EndTime = Format(EndTime, "yyyy/mm/dd hh:nn:ss")
    End If
    m_StartTime = Format(IIf(StartTime = "", Now, StartTime), "yyyy/mm/dd hh:nn:ss")
    intSegundos = DateDiff("s", m_StartTime, m_EndTime)
    intMinutos = (intSegundos - (intSegundos Mod 60)) / 60
    intHoras = (intMinutos - (intMinutos Mod 60)) / 60
    ElapsedTime = Format(intHoras, "00") & ":" & Format(intMinutos Mod 60, "00") & ":" & Format(intSegundos Mod 60, "00")

End Function

Public Function IsInCollection(CurrentCollection As Collection, Item As Variant) As Boolean
Dim CurrentItem As Variant
    IsInCollection = False
    For Each CurrentItem In CurrentCollection
        If Trim(CurrentItem) = Trim(Item) Then
            IsInCollection = True
            Exit Function
        End If
    Next CurrentItem
End Function


Public Function IsCompiled() As Boolean

  On Error GoTo NotCompiled

   Debug.Print 1 / 0
   IsCompiled = True

NotCompiled:

End Function

Public Function IsBoolean(Data As Variant) As Boolean
    On Error GoTo NotBoolean
    
    Debug.Print Data = True
    IsBoolean = True
    
NotBoolean:
    
End Function

Private Function ArrayBuscar(vntArray(), vntElementoBuscado As Variant, intColumna As Integer) As Long
Dim menor As Long
Dim mayor As Long
Dim medio As Long
Dim PosicionFinal As Long

    menor = LBound(vntArray)
    mayor = UBound(vntArray)
    PosicionFinal = -1
    Do
        medio = (menor + mayor) \ 2
        Select Case UCase(vntArray(medio, intColumna))
            Case Is = UCase(vntElementoBuscado)
                PosicionFinal = medio
            Case Is > UCase(vntElementoBuscado)
                mayor = medio - 1
            Case Is < UCase(vntElementoBuscado)
                menor = medio + 1
        End Select
    Loop Until (PosicionFinal <> -1) Or (mayor < menor)
    ArrayBuscar = PosicionFinal
End Function

Public Function TrimZeros(strString As String) As String
Dim strTemp As String
Dim i As Integer
Dim intFirstNotZero As Integer

    intFirstNotZero = 0
    For i = 1 To Len(strString)
        If Mid(strString, i, 1) <> "0" Then
            intFirstNotZero = i
            Exit For
        End If
    Next i
    If intFirstNotZero = 0 Then
        If i = Len(strString) + 1 Then
            TrimZeros = "0"
        Else
            TrimZeros = strString
        End If
    Else
        TrimZeros = Mid(strString, intFirstNotZero)
    End If
End Function

Public Function AddComma(strString As String) As String
    If Len(strString) > 2 Then
        AddComma = Left(strString, Len(strString) - 2) & "," & Right(strString, 2)
    Else
        AddComma = strString
    End If
End Function


Public Function GetIni(Grupo As String, Entrada As String) As Variant
    Dim Result As String
    Dim intResultLen As Integer
    
    Result = String$(255, 0)     'Se asigna espacio de memoria para la variable que contendra la respuesta
    intResultLen = GetPrivateProfileString(Grupo, Entrada, "", Result, Len(Result), App.Path + "\" + App.EXEName + ".ini")
    If intResultLen = 0 Then
        GetIni = Empty
    Else
        GetIni = Left(Result, intResultLen)
    End If
End Function

Public Sub WriteLog(sLogEntry As String, Optional FileName As String = "")
   Const ForReading = 1, ForWriting = 2, ForAppending = 8
   Dim sLogFile As String, sLogPath As String, iLogSize As Long
   Dim fso, f
   
On Error GoTo ErrHandler

   'Set the path and filename of the log
   If FileName = "" Then
    sLogPath = App.Path & "\" & App.EXEName
    sLogFile = sLogPath & ".log"
   Else
    sLogFile = App.Path & "\" & FileName
   End If
   
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.OpenTextFile(sLogFile, ForAppending, True)
   
    
   'Append the log-entry to the file together with time and date
   f.WriteLine Now() & vbTab & sLogEntry
   
ErrHandler:
    Exit Sub
End Sub

Public Function GetUserByCanal(strCanal As String) As String
Dim recTemp As New ADODB.Recordset
    recTemp.Open "SELECT * FROM users where name ='" & strCanal & "'", p_ADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not recTemp.EOF Then
        GetUserByCanal = recTemp("USERNAME").Value & ""
    Else
        GetUserByCanal = "CANTOLIN"
    End If
    Set recTemp = Nothing
End Function





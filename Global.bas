Attribute VB_Name = "Global"
Global Const ForReading = 1, ForWriting = 2, ForAppending = 3
Global Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
Public p_StrFieldSeperator As String
Public p_ADODBConnection As ADODB.Connection
Public p_ADODBConnectionOLDGM As ADODB.Connection

Public Function SplitStringInObject(strCurrent As String) As Contact
Dim CurrentContact As New Contact
Dim intPosInicial As Integer
Dim intPosFinal As Integer
Dim strTemp As String

    'Busco el CUIT
    intPosInicial = 1
    intPosFinal = InStr(intPosInicial, strCurrent, p_StrFieldSeperator)
    CurrentContact.CUIT = Mid(strCurrent, intPosInicial, intPosFinal - intPosInicial - 1)
    
    'Razon Social
    intPosInicial = intPosFinal
    intPosFinal = InStr(intPosInicial + 1, strCurrent, p_StrFieldSeperator)
    CurrentContact.RazonSocial = Mid(strCurrent, intPosInicial + 1, intPosFinal - intPosInicial - 1)
    
    'El Domicilio
    intPosInicial = intPosFinal
    intPosFinal = InStr(intPosInicial + 1, strCurrent, p_StrFieldSeperator)
    CurrentContact.Domicilio = Mid(strCurrent, intPosInicial + 1, intPosFinal - intPosInicial - 1)
    
    'Actividad
    intPosInicial = intPosFinal
    intPosFinal = InStr(intPosInicial + 1, strCurrent, p_StrFieldSeperator)
    CurrentContact.CodigoActividad = Mid(strCurrent, intPosInicial + 1, intPosFinal - intPosInicial - 1)
    
    'Periodo
    intPosInicial = intPosFinal
    intPosFinal = InStr(intPosInicial + 1, strCurrent, p_StrFieldSeperator)
    CurrentContact.Periodo = Mid(strCurrent, intPosInicial + 1, intPosFinal - intPosInicial - 1)
    
    'Empleados
    intPosInicial = intPosFinal
    intPosFinal = InStr(intPosInicial + 1, strCurrent, p_StrFieldSeperator)
    CurrentContact.Empleados = Mid(strCurrent, intPosInicial + 1, intPosFinal - intPosInicial - 1)
    
    'Masa Salarial
    intPosInicial = intPosFinal
    intPosFinal = InStr(intPosInicial + 1, strCurrent, p_StrFieldSeperator)
    CurrentContact.MasaSalarial = Mid(strCurrent, intPosInicial + 1, intPosFinal - intPosInicial - 1)
    
    'Fecha de presentación (Pago afip)
    intPosInicial = intPosFinal
    intPosFinal = InStr(intPosInicial + 1, strCurrent, p_StrFieldSeperator)
    CurrentContact.Fechapresentacion = Mid(strCurrent, intPosInicial + 1, intPosFinal - intPosInicial - 1)
    
    'Personal Temporal
    intPosInicial = intPosFinal
    intPosFinal = InStr(intPosInicial + 1, strCurrent, p_StrFieldSeperator)
    CurrentContact.PersonalTemporal = Mid(strCurrent, intPosInicial + 1, intPosFinal - intPosInicial - 1)
    
    'Alicuta
    intPosInicial = intPosFinal
    intPosFinal = InStr(intPosInicial + 1, strCurrent, p_StrFieldSeperator)
    CurrentContact.Alicuta = Mid(strCurrent, intPosInicial + 1, intPosFinal - intPosInicial - 1)
    
    'Fijo
    intPosInicial = intPosFinal
    intPosFinal = InStr(intPosInicial + 1, strCurrent, p_StrFieldSeperator)
    CurrentContact.Fijo = Mid(strCurrent, intPosInicial + 1, intPosFinal - intPosInicial - 1)
    
    'Pago Total
    intPosInicial = intPosFinal
    intPosFinal = InStr(intPosInicial + 1, strCurrent, p_StrFieldSeperator)
    CurrentContact.PagoTotal = Mid(strCurrent, intPosInicial + 1, intPosFinal - intPosInicial - 1)
    
    'ART
    intPosInicial = intPosFinal
    intPosFinal = InStr(intPosInicial + 1, strCurrent, p_StrFieldSeperator)
    CurrentContact.CodigoART = Mid(strCurrent, intPosInicial + 1, intPosFinal - intPosInicial - 1)
    
    
    Set SplitStringInObject = CurrentContact
End Function


Public Function GenerateAccountno(strNombre As String) As String
    GenerateAccountno = Format(Now, "yymmdd") & Format(DateDiff("s", Format(Format(Now, "yyyy/mm/dd") & " 00:00:00", "yyyy/mm/dd hh:nn:ss"), Now), "00000") & GenerateRandomString(6) & Left(strNombre, 3)

End Function

Public Function GenerateRecID() As String
    GenerateRecID = GenerateRandomString(4) & Format(Now, "yymmdd") & Format(DateDiff("s", Format(Format(Now, "yyyy/mm/dd") & " 00:00:00", "yyyy/mm/dd hh:nn:ss"), Now), "00000")

End Function

Public Function GenerateRandomString(intLen As Integer) As String
Dim i As Integer
Dim intAscii As Integer
Dim strTemp As String

    Randomize
    strTemp = ""
    For i = 1 To intLen
        intAscii = Int((85 - 41 + 1) * Rnd + 41)
        strTemp = strTemp & Chr(intAscii)
    Next i
    GenerateRandomString = strTemp
End Function

Public Function ConvertStringToBinary(strString As String) As Variant
Dim recTemp As ADODB.Recordset
        Set recTemp = New ADODB.Recordset
        recTemp.Open "SELECT CONVERT(binary(60), '" & strString & "') AS Binario", p_ADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not recTemp.EOF Then
            'recHist("NOTES").Value = Trim(xlSht.Cells(i, 5))
            'recHist("NOTES").Value = recTemp("Binario").Value
            ConvertStringToBinary = recTemp("Binario").Value
        End If
        Set recTemp = Nothing

End Function

Public Sub AddEmailAddress(AccountNo As String, Email As String)
Dim recContSupp As New ADODB.Recordset

    recContSupp.Open "SELECT * from ContSupp where Accountno='AA'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    recContSupp.AddNew
        'recContSupp("").Value
        recContSupp("AccountNo").Value = AccountNo
        recContSupp("RECTYPE").Value = "P"
        recContSupp("Title").Value = ""
        recContSupp("Contact").Value = "E-mail Address"
        recContSupp("CONTSUPREF").Value = Email
        recContSupp("ZIP").Value = "011"
        recContSupp("U_CONTACT").Value = "E-MAIL ADDRESS"
        recContSupp("U_CONTSUPREF").Value = Email
        recContSupp("U_ADDRESS1").Value = ""
        recContSupp("recid").Value = GenerateRecID
    recContSupp.Update
    Set recContSupp = Nothing
End Sub

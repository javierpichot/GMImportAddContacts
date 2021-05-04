VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fmrMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Goldmine additional contacts import tool - Vladimir"
   ClientHeight    =   5475
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   12195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLimpiezaPoliza 
      Caption         =   "Procesar XLS de Contactos Adicionales"
      Height          =   1005
      Left            =   240
      TabIndex        =   3
      Top             =   1500
      Width           =   3645
   End
   Begin VB.TextBox txtLog 
      Height          =   4695
      Left            =   4020
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   660
      Width           =   8115
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4050
      Top             =   630
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblRPS 
      Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   420
      Width           =   1815
   End
   Begin VB.Label lblInfo 
      Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   30
      Width           =   9255
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Salir"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "fmrMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngLastRecord As Long
Dim lngCurrentRecord As Long
Dim blnProcesaLimpieza As Boolean



Private Sub cmdLimpiezaPoliza_Click()
Dim xlTmp As Excel.Application
Dim strFile As String
Dim xlSht As Excel.Worksheet
Dim recContact1 As New ADODB.Recordset
Dim recContact2 As New ADODB.Recordset
Dim recPending As New ADODB.Recordset
Dim recTemp As New ADODB.Recordset
Dim recNewPending As ADODB.Recordset
Dim i As Integer
Dim strCampo As String
Dim strAccountNo As String
Dim recAdditional As New ADODB.Recordset
Dim strRecIDAddContact As String
Dim strNotes As String
    On Error GoTo ERRORHANDLER
    
    CommonDialog1.ShowOpen
    strFile = CommonDialog1.FileName
    If strFile <> "" Then
        Set p_ADODBConnection = New ADODB.Connection
        p_ADODBConnection.ConnectionString = GetIni("DataSourceName", "CS")
        p_ADODBConnection.Open GetIni("DataSourceName", "CS"), GetIni("DataSourceName", "UID"), GetIni("DataSourceName", "PWD")
        lngCurrentRecord = 0
        lngLastRecord = 0
        'tomo el campo de combinacion
        strCampo = frmSelectUser.getvalue
    
        Screen.MousePointer = vbHourglass
        Set xlTmp = New Excel.Application
        xlTmp.Workbooks.Open strFile
        Set xlSht = xlTmp.Sheets(1)
        i = Val(InputBox("Desde que linea del XLS empezamos?", App.Title, "3"))
        txtLog.Text = ""
        Screen.MousePointer = vbDefault
        

        Do While Trim(xlSht.Cells(i, 1)) <> ""
            
            'Debug.Print xlSht.Cells(i, 2) & " - " & i
            recContact1.Open "SELECT * from contact1 WHERE " & strCampo & "='" & Trim(xlSht.Cells(i, 1)) & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
            If Not recContact1.EOF Then
                lblInfo.Caption = "Contact: " & recContact1("Contact").Value
                DoEvents
                'ahora que tengo el contact1 tomo el accountno
                strAccountNo = recContact1("ACCOUNTNO").Value & ""
                recAdditional.Open "SELECT * FROM CONTSUPP WHERE Accountno='aa'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
                recAdditional.AddNew
                
                'ahora inserto el contacto adicional
                    strRecIDAddContact = GenerateRecID
                    recAdditional("RECID").Value = strRecIDAddContact
                    recAdditional("AccountNo").Value = strAccountNo
                    recAdditional("RECTYPE").Value = "C"
                    
                    recAdditional("Contact").Value = Left(Trim(xlSht.Cells(i, 2)), 40)
                    recAdditional("Title").Value = Left(Trim(xlSht.Cells(i, 3)), 35)
                    recAdditional("CONTSUPREF").Value = Left(Trim(xlSht.Cells(i, 4) & ""), 35)
                    recAdditional("DEAR").Value = Trim(xlSht.Cells(i, 5))
                    recAdditional("PHONE").Value = Left(Trim(xlSht.Cells(i, 6)), 20)
                    recAdditional("EXT").Value = Left(Trim(xlSht.Cells(i, 7)), 6)
                    recAdditional("FAX").Value = Left(Trim(xlSht.Cells(i, 8)), 20)
                    recAdditional("ADDRESS1").Value = Trim(xlSht.Cells(i, 9))
                    recAdditional("ADDRESS2").Value = Trim(xlSht.Cells(i, 10))
                    recAdditional("ADDRESS3").Value = Trim(xlSht.Cells(i, 11))
                    recAdditional("CITY").Value = Trim(xlSht.Cells(i, 12))
                    recAdditional("State").Value = Trim(xlSht.Cells(i, 13))
                    recAdditional("ZIP").Value = Trim(xlSht.Cells(i, 14))
                    recAdditional("COUNTRY").Value = Trim(xlSht.Cells(i, 15))
                    recAdditional("MERGECODES").Value = Trim(xlSht.Cells(i, 16))
                    recAdditional("U_CONTACT").Value = Left(UCase(Trim(xlSht.Cells(i, 2))), 40)
                    recAdditional("U_CONTSUPREF").Value = Left(UCase(Trim(xlSht.Cells(i, 4))), 35)
                    recAdditional("U_ADDRESS1").Value = UCase(Trim(xlSht.Cells(i, 9)))
                                    
                    recAdditional("LINKACCT").Value = ""
                    recAdditional("Notes").Value = ""
                    recAdditional("Status").Value = "10"
                    recAdditional("LASTUSER").Value = "MASTER"
                    recAdditional("LASTDATE").Value = Format(Now, "YYYY/MM/DD")
                    recAdditional("LASTTIME").Value = Format(Now, "HH:NN")
                    'guardo en las notas los datos
                    strNotes = "Contact: " & Trim(xlSht.Cells(i, 2)) & "  -  " & "Tel: " & Trim(xlSht.Cells(i, 6)) & "  -  " & Trim(xlSht.Cells(i, 8)) & " - " & Trim(xlSht.Cells(i, 18)) & " - " & Trim(xlSht.Cells(i, 19))
                    If strNotes <> "" Then
                        recAdditional("NOTES") = strNotes
                    End If
                recAdditional.Update
                
                'ahora inserto el email si tiene
                If Trim(xlSht.Cells(i, 17)) <> "" Then
                    recAdditional.AddNew
                        recAdditional("RECID").Value = GenerateRecID
                        recAdditional("AccountNo").Value = strAccountNo
                        recAdditional("RECTYPE").Value = "P"
                        
                        recAdditional("Contact").Value = "E-mail Address"
                        recAdditional("Title").Value = ""
                        recAdditional("CONTSUPREF").Value = Left(Trim(xlSht.Cells(i, 17)), 35)
                        recAdditional("ADDRESS1").Value = ""
                        recAdditional("ADDRESS2").Value = Left(Trim(xlSht.Cells(i, 2)), 40)
                        recAdditional("CITY").Value = "MASTER"
                        recAdditional("ZIP").Value = "101"
                        recAdditional("MERGECODES").Value = ""
                        recAdditional("U_CONTACT").Value = UCase("E-mail Address")
                        recAdditional("U_CONTSUPREF").Value = Left(UCase(Trim(xlSht.Cells(i, 17))), 35)
                        recAdditional("U_ADDRESS1").Value = ""
                                        
                        'este es el recID del contacto
                        recAdditional("LINKACCT").Value = strRecIDAddContact
                        recAdditional("Notes").Value = ""
                        recAdditional("LASTUSER").Value = "MASTER"
                        recAdditional("LASTDATE").Value = Format(Now, "YYYY/MM/DD")
                        recAdditional("LASTTIME").Value = Format(Now, "HH:NN")
                    recAdditional.Update
                
                End If
                recAdditional.Close
            Else
                'si no encontre el cuit no pasa nada genero log
                'Generar archivo txt
                txtLog.Text = txtLog.Text & Trim(xlSht.Cells(i, 1)) & "No existe en goldmine" & vbCrLf
                WriteLog Trim(xlSht.Cells(i, 1)), "Noexiste.txt"
            End If
            recContact1.Close
            
            i = i + 1
        Loop
        
        Set recTemp = Nothing
        'xlTmp.Workbooks.Close
        xlTmp.Quit
        
    End If
    
    MsgBox "Proceso finalizado!", vbInformation, App.Title
    p_ADODBConnection.Close
    Exit Sub
    
ERRORHANDLER:
    MsgBox GetADODBErrorMessage(p_ADODBConnection.Errors), vbCritical, App.EXEName


End Sub

Private Sub Form_Load()
        
    'Levanto los parametros de la reg
    lblInfo.Caption = "Importacion aun no iniciada"
    blnProcesaLimpieza = True
    If Trim(Command$) = "/GMVC" Then
        'Corro el proceso de integracion con vocalcom
        GMVocalcommPreVenta
        End
    End If
End Sub

Private Sub ProcessDetalle1()
Dim recImportacion As New ADODB.Recordset
Dim recDetalle As Recordset
Dim strAccountNo As String
Dim strCampo2 As String
Dim strCampo3 As String

    'On Error GoTo ERRORHANDLER:
    'Primero Abro el recordset de la importacion
    recImportacion.Open "SELECT * FROM basecrm where campo1 is not null", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    Do While Not recImportacion.EOF
        strAccountNo = GetAccountno(Str(recImportacion("ID_Contact").Value))
        If strAccountNo <> "" Then
            lblInfo.Caption = Str(recImportacion("ID_Contact").Value)
            DoEvents
            'Inserto el primer detalle
            Set recDetalle = New ADODB.Recordset
            recDetalle.Open "SELECT * from Contsupp where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
            recDetalle.AddNew
                recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                recDetalle("recid").Value = GenerateRecID
                recDetalle("rectype").Value = "P"
                recDetalle("Contact").Value = "Cirugia OC"
                recDetalle("u_Contact").Value = "Cirugia OC"
                recDetalle("Contsupref") = recImportacion("Campo1").Value
                recDetalle("u_Contsupref") = recImportacion("Campo1").Value
                'Aca ahora segun lo que viene hacemos algo
                strCampo3 = ""
                strCampo2 = Left(recImportacion("Campo2").Value, 20)
                If InStr(1, UCase(strCampo2), "TRAN") Then
                    strCampo2 = "Transitoria"
                Else
                    If InStr(1, UCase(strCampo2), "DEF") Then
                        strCampo2 = "Definitiva"
                    Else
                        strCampo3 = strCampo2
                        strCampo2 = ""
                    End If
                End If
                recDetalle("Title") = strCampo2
                recDetalle.Update
                
                'Ahora chequeo si estos putos pusieron lo que va en el campo3 ademas en el 2
                If strCampo3 <> "" Then
                    'inserto un detalle adicional, que puede ser el unico
                    recDetalle.AddNew
                    recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                    recDetalle("recid").Value = GenerateRecID
                    recDetalle("rectype").Value = "P"
                    recDetalle("Contact").Value = "Patologia OC"
                    recDetalle("u_Contact").Value = "Patologia OC"
                    recDetalle("Contsupref") = Left(strCampo3, 20)
                    recDetalle("u_Contsupref") = Left(strCampo3, 20)
                    recDetalle.Update
                End If
                    
                If Not IsNull(recImportacion("Campo3").Value) Then
                    recDetalle.AddNew
                    recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                    recDetalle("recid").Value = GenerateRecID
                    recDetalle("rectype").Value = "P"
                    recDetalle("Contact").Value = "Patologia OC"
                    recDetalle("u_Contact").Value = "Patologia OC"
                    recDetalle("Contsupref") = Left(recImportacion("Campo3").Value, 20)
                    recDetalle("u_Contsupref") = Left(recImportacion("Campo3").Value, 20)
                    recDetalle.Update
                End If
            recDetalle.Close
            Set recDetalle = Nothing
        End If
        recImportacion.MoveNext
    Loop
    

    Set p_ADODBConnection = Nothing
    Exit Sub
    
ERRORHANDLER:
    On Error Resume Next
    MsgBox "Error: " & Err.Description, vbCritical, App.EXEName

End Sub

Private Sub InsertOrUpdateContact(objContact As Contact)
Dim recContact1 As ADODB.Recordset
Dim recContsup As ADODB.Recordset

    On Error GoTo ERRORHANDLER
    'Primero busco el contacto si no esta lo creo
    Set recContact1 = New ADODB.Recordset
    recContact1.Open "SELECT * FROM Contact1 WHERE Key1='" & objContact.CUIT & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    If recContact1.EOF Then
        'Inserto contacto
        recContact1.AddNew
        recContact1("ACCOUNTNO").Value = GenerateAccountno(objContact.RazonSocial)
        recContact1("COMPANY").Value = Left(Trim(objContact.RazonSocial), 40)
        recContact1("U_COMPANY").Value = Left(Trim(objContact.RazonSocial), 40)
        recContact1("U_CONTACT").Value = Left(Trim(objContact.RazonSocial), 40)
        'recContact1("Contact").Value
        'recContact1("LASTNAME").Value
        'recContact1("DEPARTMENT").Value
        'recContact1("Title").Value
        'recContact1("SECR").Value
        recContact1("PHONE1").Value = " "
        recContact1("ADDRESS1").Value = Left(objContact.DomicilioAddress, 40)
        'recContact1("ADDRESS2").Value
        'recContact1("ADDRESS3").Value
        recContact1("CITY").Value = objContact.DomicilioCity
        recContact1("U_CITY").Value = objContact.DomicilioCity
        'recContact1("State").Value
        recContact1("U_State").Value = ""
        recContact1("ZIP").Value = objContact.DomicilioZip
        recContact1("COUNTRY").Value = "Argentina"
        recContact1("U_COUNTRY").Value = "Argentina"
        recContact1("KEY1").Value = objContact.CUIT
        recContact1("KEY2").Value = ""
        recContact1("KEY3").Value = ""
        recContact1("KEY4").Value = ""
        recContact1("KEY5").Value = ""
        recContact1("U_KEY1").Value = objContact.CUIT
        recContact1("U_KEY2").Value = ""
        recContact1("U_KEY3").Value = ""
        recContact1("U_KEY4").Value = ""
        recContact1("U_KEY5").Value = ""
        recContact1("NOTES").Value = objContact.Domicilio
        'recContact1("CREATEBY").Value
        recContact1("CREATEON").Value = Format(Now, "yyyy/mm/dd")
        recContact1("RECID").Value = GenerateRecID
        recContact1("U_Lastname").Value = " "
        recContact1.Update
    Else
    End If
'Siempre inserto el detalle
    'Inserto presentacion
    Set recContsup = New ADODB.Recordset
    recContsup.Open "SELECT * FROM Contsupp  WHERE accountno='" & recContact1("ACCOUNTNO").Value & "' AND contsupref='" & objContact.Periodo & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    If recContsup.EOF Then
        recContsup.AddNew
        recContsup("ACCOUNTNO").Value = recContact1("ACCOUNTNO").Value
        recContsup("REcID").Value = GenerateRecID
    End If
    
    recContsup("Contact").Value = "Presentacion"
    recContsup("u_Contact").Value = "Presentacion"
    recContsup("Contsupref").Value = objContact.Periodo 'Campo1
    recContsup("u_Contsupref").Value = objContact.Periodo
    recContsup("rectype").Value = "P"
    recContsup("linkacct").Value = TrimZeros(objContact.CodigoActividad) 'Campo2
    recContsup("country").Value = TrimZeros(objContact.CodigoART) 'Campo3
    recContsup("zip").Value = AddComma(TrimZeros(objContact.PagoTotal)) 'Campo4
    recContsup("ext").Value = TrimZeros(objContact.Alicuta) 'Campo5
    recContsup("state").Value = TrimZeros(objContact.Fechapresentacion) 'Campo6
    recContsup("Address1").Value = AddComma(TrimZeros(objContact.Fijo)) 'Campo7
    recContsup("Address2").Value = AddComma(TrimZeros(objContact.MasaSalarial)) 'Campo8
    
    recContsup.Update

    
    recContact1.Close
    recContsup.Close
    
    Set recContact1 = Nothing
    Set recContsup = Nothing
    Exit Sub
ERRORHANDLER:
    MsgBox GetADODBErrorMessage(p_ADODBConnection.Errors), vbCritical, App.EXEName
    On Error Resume Next
    Set recContact1 = Nothing
    Set recContsup = Nothing
End Sub

Private Function GetAccountno(strRegistro As String) As String
Dim recSearch As New ADODB.Recordset

    On Error GoTo ERRORHANDLER:
    GetAccountno = ""
    recSearch.Open "select Accountno from Contact2 where Userdef01='" & Trim(strRegistro) & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    If Not recSearch.EOF Then
        GetAccountno = recSearch("Accountno").Value & ""
        
    End If
    
    Set recSearch = Nothing
    Exit Function
    
ERRORHANDLER:
    On Error Resume Next
    MsgBox "Error: " & Err.Description, vbCritical, App.EXEName


End Function


Private Sub ProcessDetalle2()
Dim recImportacion As New ADODB.Recordset
Dim recDetalle As Recordset
Dim strAccountNo As String
Dim strCampo2 As String
Dim strCampo3 As String

    'On Error GoTo ERRORHANDLER:
    'Primero Abro el recordset de la importacion
    recImportacion.Open "SELECT * FROM basecrm  where campo1 is not null", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    Do While Not recImportacion.EOF
        strAccountNo = GetAccountno(Str(recImportacion("ID_Contact").Value))
        If strAccountNo <> "" Then
            lblInfo.Caption = Str(recImportacion("ID_Contact").Value)
            DoEvents
            'Inserto el primer detalle que es el del primer llamado
            Set recDetalle = New ADODB.Recordset
            recDetalle.Open "SELECT * from Contsupp where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
            recDetalle.AddNew
                On Error GoTo retry
                recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                
                recDetalle("recid").Value = GenerateRecID
                recDetalle("rectype").Value = "P"
                recDetalle("Contact").Value = "Primer Llamado OC"
                recDetalle("u_Contact").Value = "Primer Llamado OC"
                recDetalle("Contsupref") = ""
                recDetalle("u_Contsupref") = ""
                
                recDetalle("Title").Value = Left(recImportacion("SERVICIO_ASESORIA").Value & "", 20) 'Campo1
                recDetalle("linkacct").Value = Left(recImportacion("CANT_VISITAS").Value & "", 20) 'Campo2
                recDetalle("country").Value = Left(recImportacion("VISITO_COMPETENCIA").Value & "", 20) 'Campo3
                recDetalle("zip").Value = Left(recImportacion("EMPRESA_COMPETENCIA").Value & "", 10) 'Campo4
                'recDetalle("ext").Value = TrimZeros(objContact.Alicuta) 'Campo5
                'recDetalle("state").Value = Left(recImportacion("DONDE_COMPRA").Value & "", 20) 'Campo6
                recDetalle("Address1").Value = Left(recImportacion("OBSERVACIONES").Value & "", 40) 'Campo7
                'recDetalle("Address2").Value = Left(recImportacion("OBTIENE_PRODUCTO").Value & "", 40) 'Campo8
                
                recDetalle.Update
                
                If Not IsNull(recImportacion("MARCA_BOLSA").Value) Then
                    recDetalle.AddNew
                    recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                    recDetalle("recid").Value = GenerateRecID
                    recDetalle("rectype").Value = "P"
                    recDetalle("Contact").Value = "Productos Utilizados OC"
                    recDetalle("u_Contact").Value = "Productos Utilizados OC"
                    recDetalle("Contsupref") = Left(recImportacion("MODELO_COLOPLAST").Value & "", 20)
                    recDetalle("u_Contsupref") = Left(recImportacion("MODELO_COLOPLAST").Value & "", 20)
                    recDetalle("Address1") = Left(recImportacion("MARCA_BOLSA").Value, 40)
                    If Trim(recImportacion("OTRA_COMPETENCIA").Value & "") <> "" Then
                        recDetalle("Address1") = Left(recImportacion("OTRA_COMPETENCIA").Value, 40)
                    End If

                    recDetalle.Update
                End If
                'Para la marca de la competencia ahora lo ponemos en el mismo
'                If Not IsNull(recImportacion("OTRA_COMPETENCIA").Value) Then
'                    recDetalle.AddNew
'                    recDetalle("Accountno").Value = ReplaceQuote(strAccountno)
'                    recDetalle("recid").Value = GenerateRecID
'                    recDetalle("rectype").Value = "P"
'                    recDetalle("Contact").Value = "Productos Utilizados OC"
'                    recDetalle("u_Contact").Value = "Productos Utilizados OC"
'                    recDetalle("Contsupref") = ""
'                    recDetalle("u_Contsupref") = ""
'                    recDetalle("Title") = Left(recImportacion("OTRA_COMPETENCIA").Value, 20)
'                    recDetalle.Update
'                End If
                'Producto recomendad
                If Not IsNull(recImportacion("PRODUCT").Value) Then
                    recDetalle.AddNew
                    recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                    recDetalle("recid").Value = GenerateRecID
                    recDetalle("rectype").Value = "P"
                    recDetalle("Contact").Value = "Producto Recomendado OC"
                    recDetalle("u_Contact").Value = "Producto Recomendado OC"
                    recDetalle("Contsupref") = Left(recImportacion("PRODUCT").Value, 20)
                    recDetalle("u_Contsupref") = Left(recImportacion("PRODUCT").Value, 20)
                    
                    recDetalle.Update
                End If
            
            
            recDetalle.Close
            Set recDetalle = Nothing
        End If
retry:
        recImportacion.MoveNext
    Loop
    

    Set p_ADODBConnection = Nothing
    Exit Sub
    
ERRORHANDLER:
    On Error Resume Next
    MsgBox "Error: " & Err.Description, vbCritical, App.EXEName

End Sub

Private Sub ProcessReferencias(Base As String)
Dim recImportacion As New ADODB.Recordset
Dim recReferencia As New ADODB.Recordset
Dim recContacto As New ADODB.Recordset
Dim recDetalle As Recordset
Dim strAccountNo As String
Dim strRecID1 As String
Dim strRecID2 As String
Dim strCampo2 As String
Dim strCampo3 As String

    'On Error GoTo ERRORHANDLER:
    'Primero Abro el recordset de la importacion con las entidades
    If Base = "OC" Then
        recImportacion.Open "select * from basecrm inner join instituciones on basecrm.Entity=instituciones.Institucion", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    Else
        recImportacion.Open "select * from basecc inner join instituciones on basecc.Entity=instituciones.Institucion", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    End If
    
    Do While Not recImportacion.EOF
        strAccountNo = GetAccountno(Str(recImportacion("ID_Contact").Value))
        If strAccountNo <> "" Then
            lblInfo.Caption = Str(recImportacion("ID_Contact").Value)
            DoEvents
            'Obtengo el contacto
            recContacto.Open "SELECT * from Contact1 where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
            'Primero obtengo la referencia, la referencia la vamos a buscar por el campo completo
            recReferencia.Open "SELECT * from contact1 inner join contact2 on contact1.accountno=contact2.accountno where uinstlargo='" & recImportacion("Institucion").Value & "'", p_ADODBConnection, adOpenKeyset, adLockReadOnly, adcmdtect
            'Inserto la referencia
            If Not recReferencia.EOF Then
                'Inserto la referencia de ida
                Set recDetalle = New ADODB.Recordset
                recDetalle.Open "SELECT * from Contsupp where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
                recDetalle.AddNew
                    On Error GoTo retry
                    recDetalle("Accountno").Value = recReferencia("Accountno").Value
                    strRecID1 = GenerateRecID
                    recDetalle("recid").Value = strRecID1
                    recDetalle("rectype").Value = "R"
                    recDetalle("Contact").Value = Left("A:" & recContacto("Contact").Value, 30)
                    recDetalle("u_Contact").Value = Left("A:" & recContacto("Contact").Value, 30)
                    recDetalle("Contsupref") = "Institucion donde se atendio" 'Aca va el detalle
                    recDetalle("u_Contsupref") = "Institucion donde se atendio"
                    recDetalle("Title").Value = recContacto("Accountno").Value 'aca el accountno del otro
                    strRecID2 = GenerateRecID ' Lo genero aca para tenerlo abajo
                    recDetalle("linkacct").Value = strRecID2 'Aca va a ir el recid del otro registro
                    recDetalle("ext").Value = "T"
                    recDetalle.Update
                recDetalle.Close
                'inserto la referencia de vuelta
                Set recDetalle = New ADODB.Recordset
                recDetalle.Open "SELECT * from Contsupp where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
                recDetalle.AddNew
                    On Error GoTo retry
                    recDetalle("Accountno").Value = recContacto("Accountno").Value
                    'strRecID2 = GenerateRecID
                    recDetalle("recid").Value = strRecID2
                    recDetalle("rectype").Value = "R"
                    recDetalle("Contact").Value = Left("Para:" & recReferencia("Company").Value, 30)
                    recDetalle("u_Contact").Value = Left("Para:" & recContacto("Company").Value, 30)
                    recDetalle("Contsupref") = "Institucion donde se atendio" 'Aca va el detalle
                    recDetalle("u_Contsupref") = "Institucion donde se atendio"
                    recDetalle("Title").Value = recReferencia("Accountno").Value 'aca el accountno del otro
                    recDetalle("linkacct").Value = strRecID1 'Aca va a ir el recid del otro registro
                    recDetalle("ext").Value = "R"
                    recDetalle.Update
                recDetalle.Close
                                
                Set recDetalle = Nothing
                recContacto.Close
                recReferencia.Close
            End If
        End If
retry:
        recImportacion.MoveNext
    Loop
    

    Set p_ADODBConnection = Nothing
    Exit Sub
    
ERRORHANDLER:
    On Error Resume Next
    MsgBox "Error: " & Err.Description, vbCritical, App.EXEName

End Sub


Private Sub ProcessReferenciasOS(Base As String)
Dim recImportacion As New ADODB.Recordset
Dim recReferencia As New ADODB.Recordset
Dim recContacto As New ADODB.Recordset
Dim recDetalle As Recordset
Dim strAccountNo As String
Dim strRecID1 As String
Dim strRecID2 As String
Dim strCampo2 As String
Dim strCampo3 As String

    'On Error GoTo ERRORHANDLER:
    'Primero Abro el recordset de la importacion con las entidades
    If Base = "OC" Then
        recImportacion.Open "select * from basecrm inner join os on basecrm.Membership=os.obrasocialnl", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    Else
        recImportacion.Open "select * from basecc inner join os on basecc.Membership=os.obrasocialnl", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    End If
    
    Do While Not recImportacion.EOF
        strAccountNo = GetAccountno(Str(recImportacion("ID_Contact").Value))
        If strAccountNo <> "" Then
            lblInfo.Caption = Str(recImportacion("ID_Contact").Value)
            DoEvents
            'Obtengo el contacto
            recContacto.Open "SELECT * from Contact1 where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
            'Primero obtengo la referencia, la referencia la vamos a buscar por el campo completo
            recReferencia.Open "SELECT * from contact1 inner join contact2 on contact1.accountno=contact2.accountno where uinstlargo='" & recImportacion("obrasocialnl").Value & "'", p_ADODBConnection, adOpenKeyset, adLockReadOnly, adcmdtect
            'Inserto la referencia
            If Not recReferencia.EOF Then
                'Inserto la referencia de ida
                Set recDetalle = New ADODB.Recordset
                recDetalle.Open "SELECT * from Contsupp where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
                recDetalle.AddNew
                    'On Error GoTo retry
                    recDetalle("Accountno").Value = recReferencia("Accountno").Value
                    strRecID1 = GenerateRecID
                    recDetalle("recid").Value = strRecID1
                    recDetalle("rectype").Value = "R"
                    recDetalle("Contact").Value = Left("A:" & recContacto("Contact").Value, 30)
                    recDetalle("u_Contact").Value = Left("A:" & recContacto("Contact").Value, 30)
                    recDetalle("Contsupref") = "Obra social de pertenencia" 'Aca va el detalle
                    recDetalle("u_Contsupref") = "Obra social de pertenencia"
                    recDetalle("Title").Value = recContacto("Accountno").Value 'aca el accountno del otro
                    strRecID2 = GenerateRecID ' Lo genero aca para tenerlo abajo
                    recDetalle("linkacct").Value = strRecID2 'Aca va a ir el recid del otro registro
                    recDetalle("ext").Value = "T"
                    recDetalle.Update
                recDetalle.Close
                'inserto la referencia de vuelta
                Set recDetalle = New ADODB.Recordset
                recDetalle.Open "SELECT * from Contsupp where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
                recDetalle.AddNew
                    'On Error GoTo retry
                    recDetalle("Accountno").Value = recContacto("Accountno").Value
                    'strRecID2 = GenerateRecID
                    recDetalle("recid").Value = strRecID2
                    recDetalle("rectype").Value = "R"
                    recDetalle("Contact").Value = Left("Para:" & recReferencia("Company").Value, 30)
                    recDetalle("u_Contact").Value = Left("Para:" & recContacto("Company").Value, 30)
                    recDetalle("Contsupref") = "Obra social de pertenencia" 'Aca va el detalle
                    recDetalle("u_Contsupref") = "Obra social de pertenencia"
                    recDetalle("Title").Value = recReferencia("Accountno").Value 'aca el accountno del otro
                    recDetalle("linkacct").Value = strRecID1 'Aca va a ir el recid del otro registro
                    recDetalle("ext").Value = "R"
                    recDetalle.Update
                recDetalle.Close
                                
                Set recDetalle = Nothing
                recContacto.Close
                recReferencia.Close
            Else
                recContacto.Close
                recReferencia.Close
            
            End If
        Else
            recContacto.Close
            recReferencia.Close
        End If
retry:
        recImportacion.MoveNext
    Loop
    

    Set p_ADODBConnection = Nothing
    Exit Sub
    
ERRORHANDLER:
    On Error Resume Next
    MsgBox "Error: " & Err.Description, vbCritical, App.EXEName

End Sub


Private Sub ProcessDetalleTMK2()
Dim recImportacion As New ADODB.Recordset
Dim recDetalle As Recordset
Dim strAccountNo As String
Dim strCampo2 As String
Dim strCampo3 As String

    'On Error GoTo ERRORHANDLER:
    'Primero Abro el recordset de la importacion
    recImportacion.Open "SELECT * FROM TMK2", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    Do While Not recImportacion.EOF
        strAccountNo = GetAccountno(Str(recImportacion("ID_Contact").Value))
        If strAccountNo <> "" Then
            lblInfo.Caption = Str(recImportacion("ID_Contact").Value)
            'Importo los campos del contact1 primero
            Set recDetalle = New ADODB.Recordset
            recDetalle.Open "SELECT * FROM Contact1 where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
                recDetalle("Notes").Value = recImportacion("observaciones").Value & ""
                recDetalle("Key2").Value = recImportacion("Estado").Value & ""
                recDetalle("Key5").Value = recImportacion("Comprador").Value & ""
            recDetalle.Update
            recDetalle.Close
            'Importo los campos del contact2 primero
            recDetalle.Open "SELECT * FROM Contact2 where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
                recDetalle("Ucono1").Value = recImportacion("Informacion").Value & ""
                recDetalle("UMedicorec").Value = recImportacion("MEDICO_RECETA").Value & ""
                recDetalle("UFechaVis2").Value = IIf(IsNull(recImportacion("Fecha_Visita").Value), Null, recImportacion("Fecha_Visita").Value)
            recDetalle.Update
            recDetalle.Close
            
            DoEvents
            'Inserto el primer detalle que es el del segundo llamado
            
            recDetalle.Open "SELECT * from Contsupp where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
            recDetalle.AddNew
                'On Error GoTo retry
                recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                
                recDetalle("recid").Value = GenerateRecID
                recDetalle("rectype").Value = "P"
                recDetalle("Contact").Value = "Segundo Llamado OC"
                recDetalle("u_Contact").Value = "Segundo Llamado OC"
                recDetalle("Contsupref") = ""
                recDetalle("u_Contsupref") = ""
                
                recDetalle("Title").Value = Left(recImportacion("OBTENCION_PRODUCTO").Value & "", 20) 'Campo1
                recDetalle("linkacct").Value = Left(recImportacion("CANT_BOLSAS_MES_OS").Value & "", 20) 'Campo2
                recDetalle("country").Value = Left(recImportacion("CANT_BOLSAS_MES_USA").Value & "", 20) 'Campo3
                recDetalle("zip").Value = Left(recImportacion("CONFORMIDAD_BOLSA").Value & "", 10) 'Campo4
                recDetalle("ext").Value = Left(recImportacion("PROBO_OTRA_MARCA").Value & "", 6) 'Campo5
                recDetalle("state").Value = Left(recImportacion("CUAL_OTRA_MARCA").Value & "", 20) 'Campo6
                recDetalle("Address1").Value = Left(recImportacion("MOTIVO_DECISION_MARCA").Value & "", 40) 'Campo7
                recDetalle("Address2").Value = Left(recImportacion("MOTIVO_DECISION_MARCA2").Value & "", 40) 'Campo8
                recDetalle("city").Value = Format(recImportacion("FECHA_LLAMADO").Value & "", "YYYYMMDD")
                
                recDetalle.Update
                
'                If Not IsNull(recImportacion("PRODUCT").Value) Then
'                    recDetalle.AddNew
'                    recDetalle("Accountno").Value = ReplaceQuote(strAccountno)
'                    recDetalle("recid").Value = GenerateRecID
'                    recDetalle("rectype").Value = "P"
'                    recDetalle("Contact").Value = "Producto Recomendado OC"
'                    recDetalle("u_Contact").Value = "Producto Recomendado OC"
'                    recDetalle("Contsupref") = Left(recImportacion("PRODUCT").Value & "", 20)
'                    recDetalle("u_Contsupref") = Left(recImportacion("PRODUCT").Value & "", 20)
'
'                    recDetalle.Update
'                End If
                
                'Producto Utilizado
                If Not IsNull(recImportacion("CODIGO_BOLSA").Value) Then
                    recDetalle.AddNew
                    recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                    recDetalle("recid").Value = GenerateRecID
                    recDetalle("rectype").Value = "P"
                    recDetalle("Contact").Value = "Productos Utilizados OC"
                    recDetalle("u_Contact").Value = "Productos Utilizados OC"
                    recDetalle("Contsupref") = ""
                    recDetalle("u_Contsupref") = ""
                    recDetalle("ext").Value = Left(recImportacion("CODIGO_BOLSA").Value & "", 6) 'Campo5
                    recDetalle("state").Value = Left(recImportacion("MARCA_BOLSA").Value & "", 20) 'Campo6
                    recDetalle("Address1").Value = Left(recImportacion("MODELO_BOLSA").Value & " " & recImportacion("DESCR_BOLSA").Value, 40) 'Campo7
                    recDetalle.Update
                End If
            
                If Not IsNull(recImportacion("cual_otro_producto_cuidado").Value) Then
                    recDetalle.AddNew
                    recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                    recDetalle("recid").Value = GenerateRecID
                    recDetalle("rectype").Value = "P"
                    recDetalle("Contact").Value = "Productos Utilizados OC"
                    recDetalle("u_Contact").Value = "Productos Utilizados OC"
                    recDetalle("Contsupref") = ""
                    recDetalle("u_Contsupref") = ""
                    recDetalle("ext").Value = Left(recImportacion("cual_otro_producto_cuidado").Value & "", 6) 'Campo5
                    recDetalle.Update
                End If
            
            
            recDetalle.Close
            Set recDetalle = Nothing
        End If
retry:
        recImportacion.MoveNext
    Loop
    

    Set p_ADODBConnection = Nothing
    Exit Sub
    
ERRORHANDLER:
    On Error Resume Next
    MsgBox "Error: " & Err.Description, vbCritical, App.EXEName

End Sub

Private Sub ProcessDetalle1CC()
Dim recImportacion As New ADODB.Recordset
Dim recDetalle As Recordset
Dim strAccountNo As String
Dim strCampo2 As String
Dim strCampo3 As String

    'On Error GoTo ERRORHANDLER:
    'Primero Abro el recordset de la importacion
    recImportacion.Open "SELECT * FROM basecc", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    Do While Not recImportacion.EOF
        strAccountNo = GetAccountno(Str(recImportacion("ID_Contact").Value))
        If strAccountNo <> "" Then
            lblInfo.Caption = Str(recImportacion("ID_Contact").Value)
            DoEvents
            'Inserto el primer detalle
            Set recDetalle = New ADODB.Recordset
            recDetalle.Open "SELECT * from Contsupp where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
            recDetalle.AddNew
            recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
            recDetalle("recid").Value = GenerateRecID
            recDetalle("rectype").Value = "P"
            recDetalle("Contact").Value = "Patologia CC"
            recDetalle("u_Contact").Value = "Patologia CC"
            recDetalle("Contsupref") = recImportacion("Campo1").Value & ""
            recDetalle("u_Contsupref") = recImportacion("Campo1").Value & ""
            recDetalle("Title") = Left(recImportacion("Campo2").Value & "", 20)
            recDetalle.Update
            recDetalle.Close
            Set recDetalle = Nothing
        End If
        recImportacion.MoveNext
    Loop
    

    Set p_ADODBConnection = Nothing
    Exit Sub
    
ERRORHANDLER:
    On Error Resume Next
    MsgBox "Error: " & Err.Description, vbCritical, App.EXEName

End Sub

Private Sub ProcessDetalleCCLlamado()
Dim recImportacion As New ADODB.Recordset
Dim recDetalle As Recordset
Dim strAccountNo As String
Dim strCampo2 As String
Dim strCampo3 As String

    'On Error GoTo ERRORHANDLER:
    'Primero Abro el recordset de la importacion
    recImportacion.Open "SELECT * FROM BaseCC", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    Do While Not recImportacion.EOF
        strAccountNo = GetAccountno(Str(recImportacion("ID_Contact").Value))
        If strAccountNo <> "" Then
            lblInfo.Caption = Str(recImportacion("ID_Contact").Value)
            Set recDetalle = New ADODB.Recordset
            DoEvents
            'Inserto el primer detalle que es el del segundo llamado
            
            recDetalle.Open "SELECT * from Contsupp where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
            recDetalle.AddNew
                'On Error GoTo retry
                recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                recDetalle("recid").Value = GenerateRecID
                recDetalle("rectype").Value = "P"
                recDetalle("Contact").Value = "Primer Llamado CC"
                recDetalle("u_Contact").Value = "Primer Llamado CC"
                recDetalle("Contsupref") = ""
                recDetalle("u_Contsupref") = ""
                
                recDetalle("Title").Value = Left(recImportacion("CATETERISMO_INTERMITENTE").Value & "", 20) 'Campo1
                recDetalle("linkacct").Value = Left(recImportacion("MARCA").Value & "", 20) 'Campo2
                recDetalle("country").Value = Left(recImportacion("CONFORMIDAD").Value & "", 20) 'Campo3
                recDetalle("zip").Value = Left(recImportacion("USUARIO_EASICATH").Value & "", 10) 'Campo4
                recDetalle("ext").Value = Left(recImportacion("ENTREGA_OS").Value & "", 6) 'Campo5
                recDetalle("state").Value = Left(recImportacion("RECONOCIMIENTO").Value & "", 20) 'Campo6
                'recDetalle("Address1").Value = Left(recImportacion("MOTIVO_DECISION_MARCA").Value & "", 40) 'Campo7
                'recDetalle("Address2").Value = Left(recImportacion("MOTIVO_DECISION_MARCA2").Value & "", 40) 'Campo8
                'recDetalle("city").Value = Format(recImportacion("FECHA_LLAMADO").Value & "", "YYYYMMDD")
                
                recDetalle.Update
                
                If Not IsNull(recImportacion("PRODUCT").Value) Then
                    recDetalle.AddNew
                    recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                    recDetalle("recid").Value = GenerateRecID
                    recDetalle("rectype").Value = "P"
                    recDetalle("Contact").Value = "Producto Utilizado CC"
                    recDetalle("u_Contact").Value = "Producto Utilizado CC"
                    recDetalle("Contsupref") = Left(recImportacion("PRODUCT").Value & "", 20)
                    recDetalle("u_Contsupref") = Left(recImportacion("PRODUCT").Value & "", 20)

                    recDetalle.Update
                End If
                            
            recDetalle.Close
            Set recDetalle = Nothing
        End If
retry:
        recImportacion.MoveNext
    Loop
    

    Set p_ADODBConnection = Nothing
    Exit Sub
    
ERRORHANDLER:
    On Error Resume Next
    MsgBox "Error: " & Err.Description, vbCritical, App.EXEName

End Sub

Public Sub CompleteCalendarActivity(CalRecID As String, RESCode As String, Notes As String)
Dim recCalendar As New ADODB.Recordset
Dim recNewHist As New ADODB.Recordset
    
    'Primero busco la actividad en el CAL
    'La borro y creo el conthist
    On Error GoTo ERRORHANDLER
    'p_ADODBConnection.BeginTrans
    recCalendar.Open "SELECT * from CAL WHERE RECID='" & CalRecID & "'", p_ADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    If recCalendar.EOF Then
        'p_ADODBConnection.RollbackTrans
        Exit Sub
    Else
        recNewHist.Open "SELECT * from CONTHIST WHERE Accountno='" & recCalendar("Accountno").Value & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
        'inserto el registro nuevo
        recNewHist.AddNew
            'recNewHist("USERID").Value = "MASTER"
            recNewHist("USERID").Value = recCalendar("userid").Value
            recNewHist("ACCOUNTNO").Value = recCalendar("Accountno").Value
            recNewHist("SRECTYPE").Value = recCalendar("RECTYPE").Value & ""
            recNewHist("RECTYPE").Value = recCalendar("RECTYPE").Value & ""
            recNewHist("ONDATE").Value = recCalendar("ONDATE").Value & ""
            recNewHist("ONTIME").Value = recCalendar("ONTIME").Value & ""
            recNewHist("ACTVCODE").Value = recCalendar("ACTVCODE").Value & ""
            recNewHist("RESULTCODE").Value = RESCode
            recNewHist("STATUS").Value = recCalendar("STATUS").Value & ""
            recNewHist("DURATION").Value = recCalendar("DURATION").Value & ""
            recNewHist("UNITS").Value = ""
            recNewHist("REF").Value = recCalendar("REF").Value & ""
            'recNewHist("NOTES").Value = Notes
            'recNewHist("NOTES").Value = ConvertStringToBinary(Notes)
            recNewHist("NOTES").Value = recCalendar("NOTES").Value
            recNewHist("CREATEBY").Value = recCalendar("CREATEBY").Value & ""
            recNewHist("CREATEON").Value = recCalendar("CREATEON").Value & ""
            recNewHist("CREATEAT").Value = recCalendar("CREATEAT").Value & ""
            recNewHist("LASTUSER").Value = recCalendar("LASTUSER").Value & ""
            If Not IsNull(recCalendar("LASTDATE").Value) Then
                recNewHist("LASTDATE").Value = recCalendar("LASTDATE").Value & ""
                recNewHist("LASTTIME").Value = recCalendar("LASTTIME").Value & ""
            End If
            'recNewHist("EXT").Value = recOldHist("EXT").Value & ""
            recNewHist("recid").Value = GenerateRecID
            recNewHist("loprecid").Value = ""
            recNewHist("completedid").Value = ""
        recNewHist.Update
        recCalendar.Close
        p_ADODBConnection.Execute "DELETE FROM CAL where recid='" & CalRecID & "'"
        'p_ADODBConnection.CommitTrans
    End If
    
    
    Exit Sub
ERRORHANDLER:

    On Error Resume Next
    'p_ADODBConnection.RollbackTrans
    MsgBox "Error: " & Err.Description, vbCritical, App.EXEName

End Sub

Private Sub GMVocalcomm()
Dim recEntradas As New ADODB.Recordset
Dim recContact1 As ADODB.Recordset
Dim recContact2 As ADODB.Recordset
Dim recContSupp As ADODB.Recordset
Dim strNewAccountno As String
Dim strNewRecid As String
Dim strTemp As String

    On Error GoTo ERRORHANDLER
    Set p_ADODBConnection = New ADODB.Connection
    p_ADODBConnection.ConnectionString = GetIni("DataSourceName", "CS")
    p_ADODBConnection.Open
    
    recEntradas.Open "SELECT * FROM MIDWARE.DBO.GMVC where PasoGM is null", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    Do While Not recEntradas.EOF
        'Busco si existe el contact1
        Set recContact1 = New ADODB.Recordset
        recContact1.Open "SELECT * FROM Contact1 where phone1='" & recEntradas("CUIT").Value & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
        If Not recContact1.EOF Then
            'Existe el contacto
            WriteLog "GMVC: Existe CUIT " & recEntradas("CUIT").Value
        Else
            'No existe el contacto, lo creo
            strNewAccountno = GenerateAccountno(recEntradas("RazonSocial").Value & "")
            recContact1.AddNew
                recContact1("ACCOUNTNO").Value = strNewAccountno
                recContact1("RECID").Value = GenerateRecID
                recContact1("PHONE1").Value = recEntradas("CUIT").Value
                recContact1("contact").Value = Left(recEntradas("RazonSocial").Value, 40)
                recContact1("Address1").Value = Left(recEntradas("Domicilio").Value, 40)
                recContact1("City").Value = Left(recEntradas("Localidad").Value, 30)
                recContact1("Phone2").Value = Left(recEntradas("Telefono").Value, 25)
                recContact1("ZIP").Value = Left(recEntradas("CP").Value & "", 4)
                recContact1("Status").Value = "I0"
                recContact1("Source").Value = "NET BROKER"
                recContact1("OWNER").Value = ""
                recContact1("U_COMPANY").Value = ""
                recContact1("U_CONTACT").Value = Left(recEntradas("RazonSocial").Value, 40)
                recContact1("U_LASTNAME").Value = ""
                recContact1("u_CITY").Value = Left(recEntradas("Localidad").Value, 30)
                recContact1("U_STATE").Value = ""
                recContact1("u_COUNTRY").Value = ""
                recContact1("U_KEY1").Value = ""
                recContact1("U_KEY2").Value = ""
                recContact1("KEY3").Value = Left(recEntradas("CIIU").Value & "", 20)
                recContact1("U_KEY3").Value = Left(recEntradas("CIIU").Value & "", 20)
                strTemp = Trim(GetIni("Agentes", recEntradas("EjecutivoAsignado").Value & ""))
                If strTemp <> "" Then
                    recContact1("KEY4").Value = strTemp
                    recContact1("U_KEY4").Value = strTemp
                Else
                    recContact1("KEY4").Value = recEntradas("EjecutivoAsignado").Value & ""
                    recContact1("U_KEY4").Value = recEntradas("EjecutivoAsignado").Value & ""
                End If
                recContact1("U_KEY5").Value = ""
                recContact1("CREATEBY").Value = "VOC2GLMD"
                recContact1("CREATEON").Value = Format(Now, "yyyy/mm/dd")
            recContact1.Update
            'ahora agrego el email
            If Trim(recEntradas("mail").Value & "") <> "" Then
                AddEmailAddress strNewAccountno, Left(recEntradas("mail").Value & "", 35)
            End If
            Set recContact2 = New ADODB.Recordset
            recContact2.Open "SELECT * from contact2 where accountno='" & strNewAccountno & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
            recContact2.AddNew
                recContact2("accountno").Value = strNewAccountno
                recContact2("recid").Value = GenerateRecID
                recContact2("ucapitas").Value = recEntradas("capitas").Value & ""
                ' si son mas de 50 capitas pongo el tipo de cliente en segmento
                If Val("0" & recEntradas("capitas").Value) >= 50 Then
                    recContact2("USEGMENTO").Value = "CLIENTE CORPORATIVO"
                End If
                recContact2("umasasrial").Value = recEntradas("masasalarial").Value & ""
                'art actual
                recContact2("userdef04").Value = Left(recEntradas("ART").Value & "", 12)
                recContact2("uav").Value = recEntradas("AV").Value & ""
                recContact2("uaf").Value = recEntradas("AF").Value & ""
                strTemp = Trim(GetIni("Agentes", recEntradas("agente").Value & ""))
                If strTemp <> "" Then
                    recContact2("utlmk").Value = strTemp
                Else
                    recContact2("utlmk").Value = recEntradas("agente").Value & ""
                End If
            recContact2.Update
            'Ahora cargo las cosas de contact2
'            Agente
'            EjecutivoAsignado
'            Campana
'            ART
'            CIIU
'            MasaSalarial
'            Capitas
'            AV
'            AF

            'Ahora marco como actualizado
            recEntradas("PasoGM").Value = Now
            recEntradas.Update
        End If
        recEntradas.MoveNext
    Loop
    Exit Sub
ERRORHANDLER:

    On Error Resume Next
    WriteLog "GMVC: " & GetADODBErrorMessage(p_ADODBConnection.Errors)
End Sub

Private Sub GMVocalcommPreVenta()
Dim recEntradas As New ADODB.Recordset
Dim recContact1 As ADODB.Recordset
Dim recContact2 As ADODB.Recordset
Dim recContSupp As ADODB.Recordset
Dim strNewAccountno As String
Dim strNewRecid As String
Dim strTemp As String
Dim strUserIDGM As String
Dim recPending As ADODB.Recordset
Dim recNewPending As ADODB.Recordset
Dim InsertarPendiente As Boolean

    On Error GoTo ERRORHANDLER
    Set p_ADODBConnection = New ADODB.Connection
    p_ADODBConnection.ConnectionString = GetIni("DataSourceName", "CS")
    p_ADODBConnection.Open
    
    recEntradas.CursorLocation = adUseClient
    recEntradas.Open "SELECT * FROM MIDWARE.DBO.GMVC where EstadoEntrevista in ('G','B','R') order by EstadoEntrevista DESC ", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    Do While Not recEntradas.EOF
        'Busco si existe el contact1
        Set recContact1 = New ADODB.Recordset
        Me.lblInfo.Caption = "Proceso: " & recEntradas("CUIT").Value
        DoEvents
        recContact1.Open "SELECT * FROM Contact1 where phone1='" & recEntradas("CUIT").Value & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
        If recContact1.EOF And recEntradas("EstadoEntrevista").Value = "G" Then
            'No existe el contacto, lo creo
            'p_ADODBConnection.BeginTrans
            strNewAccountno = GenerateAccountno(recEntradas("RazonSocial").Value & "")
            recContact1.AddNew
                recContact1("ACCOUNTNO").Value = strNewAccountno
                recContact1("RECID").Value = GenerateRecID
                recContact1("PHONE1").Value = recEntradas("CUIT").Value
                recContact1("contact").Value = Left(recEntradas("RazonSocial").Value, 40)
                recContact1("Address1").Value = Left(recEntradas("Domicilio").Value, 40)
                recContact1("City").Value = Left(recEntradas("Localidad").Value, 30)
                recContact1("Phone2").Value = Left(recEntradas("Telefono").Value, 25)
                recContact1("ZIP").Value = Left(recEntradas("CP").Value & "", 4)
                recContact1("Status").Value = "I0"
                recContact1("Source").Value = "NET BROKER"
                recContact1("OWNER").Value = ""
                recContact1("U_COMPANY").Value = ""
                recContact1("U_CONTACT").Value = Left(recEntradas("RazonSocial").Value, 40)
                recContact1("U_LASTNAME").Value = ""
                recContact1("u_CITY").Value = Left(recEntradas("Localidad").Value, 30)
                recContact1("U_STATE").Value = ""
                recContact1("u_COUNTRY").Value = ""
                recContact1("U_KEY1").Value = ""
                recContact1("U_KEY2").Value = ""
                recContact1("KEY3").Value = Left(recEntradas("CIIU").Value & "", 20)
                recContact1("U_KEY3").Value = Left(recEntradas("CIIU").Value & "", 20)
                strTemp = Trim(GetIni("Agentes", recEntradas("EjecutivoAsignado").Value & ""))
                If strTemp <> "" Then
                    recContact1("KEY4").Value = strTemp
                    recContact1("U_KEY4").Value = strTemp
                Else
                    recContact1("KEY4").Value = recEntradas("EjecutivoAsignado").Value & ""
                    recContact1("U_KEY4").Value = recEntradas("EjecutivoAsignado").Value & ""
                End If
                recContact1("U_KEY5").Value = ""
                recContact1("CREATEBY").Value = "VOC2GLMD"
                recContact1("CREATEON").Value = Format(Now, "yyyy/mm/dd")
            recContact1.Update
            'ahora agrego el email
            If Trim(recEntradas("mail").Value & "") <> "" Then
                AddEmailAddress strNewAccountno, Left(recEntradas("mail").Value & "", 35)
            End If
            Set recContact2 = New ADODB.Recordset
            recContact2.Open "SELECT * from contact2 where accountno='" & strNewAccountno & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
            recContact2.AddNew
                recContact2("accountno").Value = strNewAccountno
                recContact2("recid").Value = GenerateRecID
                recContact2("ucapitas").Value = recEntradas("capitas").Value & ""
                ' si son mas de 50 capitas pongo el tipo de cliente en segmento
                If Val("0" & recEntradas("capitas").Value) >= 50 Then
                    recContact2("USEGMENTO").Value = "CLIENTE CORPORATIVO"
                End If
                recContact2("umasasrial").Value = recEntradas("masasalarial").Value & ""
                'art actual
                recContact2("userdef04").Value = Left(recEntradas("ART").Value & "", 12)
                recContact2("uav").Value = recEntradas("AV").Value & ""
                recContact2("uaf").Value = recEntradas("AF").Value & ""
                strTemp = Trim(GetIni("Agentes", recEntradas("agente").Value & ""))
                If strTemp <> "" Then
                    recContact2("utlmk").Value = strTemp
                Else
                    recContact2("utlmk").Value = recEntradas("agente").Value & ""
                End If
                'los campos nuevos
                recContact2("UEPERSONA").Value = recEntradas("PersonaEntrevista").Value & ""
                recContact2("UEFECHA").Value = recEntradas("FechaEntrevista").Value
                recContact2("UEDOMICILI").Value = recEntradas("DireccionEntrevista").Value & ""
                recContact2("UECP").Value = recEntradas("CP").Value & ""
            
            recContact2.Update
            'p_ADODBConnection.CommitTrans
        End If
        
        'Genero el pendiente
        'ya sea lo encontre o no
        'Busco si tiene un pending
        If Not recContact1.EOF Then
            Set recPending = New ADODB.Recordset
            recPending.Open "SELECT * FROM CAL WHERE AccountNo='" & recContact1("ACCOUNTNO").Value & "' AND (REF ='ENTREVISTA' OR REF ='COTIZACION POR MAIL')", p_ADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not recPending.EOF Then
                'Solo si estoy reprogramando hago algo
                If recEntradas("EstadoEntrevista").Value = "B" Then
                    p_ADODBConnection.Execute "DELETE FROM CAL WHERE recID=" & FormatToSQL(recPending("recID").Value, gsdtString)
                    InsertarPendiente = False
                Else
                    If recEntradas("EstadoEntrevista").Value = "R" Then
                        p_ADODBConnection.Execute "DELETE FROM CAL WHERE recID=" & FormatToSQL(recPending("recID").Value, gsdtString)
                        InsertarPendiente = True
                    End If
                End If
            Else
                If recEntradas("EstadoEntrevista").Value = "G" Then
                    InsertarPendiente = True
                End If
            End If
            If InsertarPendiente Then
                Set recNewPending = New ADODB.Recordset
                recNewPending.Open "SELECT * FROM CAL WHERE Accountno='aa'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
                recNewPending.AddNew
                recNewPending("RECID").Value = GenerateRecID
                recNewPending("Accountno").Value = recContact1("Accountno").Value & ""
                recNewPending("Company").Value = recContact1("Contact").Value & ""
                strUserIDGM = Trim(GetIni("UsersGM", recEntradas("EjecutivoAsignado").Value & ""))
                recNewPending("USERID").Value = Left(strUserIDGM, 8)
                'recNewPending("ONDATE").Value = Format(DateAdd("d", 2, Now), "DD/MM/YYYY")
                'recNewPending("ENDDATE").Value = Format(DateAdd("d", 2, Now), "DD/MM/YYYY")
                recNewPending("ONDATE").Value = recEntradas("FechaEntrevista").Value
                recNewPending("ENDDATE").Value = recEntradas("FechaEntrevista").Value
                
                If recEntradas("TipoContacto").Value & "" = "M" Then
                    recNewPending("ACTVCODE").Value = "ML"
                    recNewPending("REF").Value = "ENVIAR COT POR MAIL"
                    recNewPending("RECTYPE").Value = "D" 'A: Appointment, C: Call Back, T: Next Action, D: To-Do M: Message, S: Forecasted Sale, O: Other, E: Event
                Else
                    recNewPending("ACTVCODE").Value = "ENT"
                    recNewPending("REF").Value = "ENTREVISTA"
                    recNewPending("Notes").Value = ConvertStringToBinary("Direccion de entrevista" & recEntradas("DireccionEntrevista").Value & "  Persona contacto: " & recEntradas("PersonaEntrevista").Value)
                    recNewPending("RECTYPE").Value = "A" 'A: Appointment, C: Call Back, T: Next Action, D: To-Do M: Message, S: Forecasted Sale, O: Other, E: Event
                End If
                
                recNewPending("CREATEBY").Value = Left(strUserIDGM, 8)
                recNewPending("LASTUSER").Value = Left(strUserIDGM, 8)
                recNewPending("CREATEON").Value = Format(Now, "DD/MM/YYYY")
                'rellenos
                recNewPending("ONTIME").Value = Format(recEntradas("FechaEntrevista").Value, "HH:NN")
                recNewPending("ALARMFLAG").Value = "N"
                recNewPending("ALARMTIME").Value = ""
                
                recNewPending("LINKRECID").Value = ""
                recNewPending("LOPRECID").Value = ""
                
                recNewPending.Update
                recNewPending.Close
                recPending.Close
            End If
        End If
        'Ahora marco como actualizado
        recEntradas("PasoGM").Value = Now
        recEntradas("EstadoEntrevista").Value = "T"
        recEntradas.Update
    
        
        recEntradas.MoveNext
    Loop
    
    Exit Sub
ERRORHANDLER:

    On Error Resume Next
    p_ADODBConnection.RollbackTrans
    WriteLog "GMVC: " & GetADODBErrorMessage(p_ADODBConnection.Errors)
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuHelp_Click()
Dim strHelp As String
strHelp = " Column distribution in input XLS: (mandatory) "
strHelp = strHelp & " 1 - Identificador" & vbCrLf
strHelp = strHelp & "2 - Contact" & vbCrLf
strHelp = strHelp & "3 - Title" & vbCrLf
strHelp = strHelp & "4 - CONTSUPREF" & vbCrLf
strHelp = strHelp & "5 - DEAR" & vbCrLf
strHelp = strHelp & "6 - PHONE" & vbCrLf
strHelp = strHelp & "7 - EXT" & vbCrLf
strHelp = strHelp & "8 - FAX" & vbCrLf
strHelp = strHelp & "9 - ADDRESS1" & vbCrLf
strHelp = strHelp & "10 - ADDRESS2" & vbCrLf
strHelp = strHelp & "11 - ADDRESS3" & vbCrLf
strHelp = strHelp & "12 - CITY" & vbCrLf
strHelp = strHelp & "13 - State" & vbCrLf
strHelp = strHelp & "14 - ZIP" & vbCrLf
strHelp = strHelp & "15 - COUNTRY" & vbCrLf
strHelp = strHelp & "16 - MERGECODES" & vbCrLf
strHelp = strHelp & "18 - NOTES" & vbCrLf
strHelp = strHelp & "19 - NOTES" & vbCrLf

MsgBox strHelp, vbOKOnly + vbInformation, App.Title
End Sub

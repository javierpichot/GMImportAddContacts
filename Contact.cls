VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Contact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private vDomicilio As String
Public CUIT As String
Public RazonSocial As String
Public CodigoActividad As String
Public CodigoART As String
Public Periodo  As String
Public Empleados As String
Public MasaSalarial As String
Public Fechapresentacion As String
Public PersonalTemporal As String
Public Alicuta As String
Public Fijo As String
Public PagoTotal As String


Public Property Get DomicilioAddress() As String
    Select Case FormatoDireccion
        Case Is = 1
            DomicilioAddress = Trim(Left(vDomicilio, 30))
        Case Is = 2
            DomicilioAddress = Trim(Left(vDomicilio, 31))
        Case Is = 3
            DomicilioAddress = Trim(vDomicilio)
    End Select
    
End Property

Public Property Get DomicilioCity() As String
    Select Case FormatoDireccion
        Case Is = 1
            DomicilioCity = Trim(Mid(vDomicilio, 31, 55 - 31))
        Case Is = 2
            DomicilioCity = Trim(Mid(vDomicilio, 43))
        Case Is = 3
            DomicilioCity = ""
    End Select
End Property
Public Property Get DomicilioZip() As String
    Select Case FormatoDireccion
        Case Is = 1
            DomicilioZip = Mid(vDomicilio, 56, 4)
        Case Is = 2
            DomicilioZip = Mid(vDomicilio, 32, 4)
        Case Is = 3
            DomicilioZip = ""
    End Select
End Property
    
Public Property Let Domicilio(vData As String)
    vDomicilio = vData
End Property

Public Property Get Domicilio() As String
    Domicilio = vDomicilio
End Property


Private Function FormatoDireccion() As Integer
    If IsNumeric(Right(vDomicilio, 6)) Then
        'Codigo Postal al final, Localidad en posicion 30 a 55
        FormatoDireccion = 1
    Else
        If IsNumeric(Mid(vDomicilio, 32, 4)) Then
            'CodigoPostal en posicion 31 a 35 y localidad en posicion 42 a final
            FormatoDireccion = 2
        Else
            'Tomo todo
            FormatoDireccion = 3
        End If
    End If
    
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRockwell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'N354:0>011101
'N401:126
Dim sAddrBase As String, lLow As Long, lHigh As Long, bIsLogix As Boolean, lAddrBit As Long
Dim sPT_TYPE As String

Public Property Let PT_TYPE(S As String)
    sPT_TYPE = S
End Property

Public Sub Address(ByVal Addr As String, ByVal AddrBit As Long, ByVal nElems As Long)
    Dim P1 As Long, P2 As Long
    P1 = InStr(Addr, ":")
    P2 = InStr(Addr, ">")
    bIsLogix = (P1 > 0)
    If (P1 > 0) And (P2 > 0) Then
        lLow = Mid(Addr, P1 + 1, P2 - P1 - 1)
        sAddrBase = Left(Addr, P1 - 1)
    ElseIf bIsLogix Then
        lLow = Mid(Addr, P1 + 1)
        sAddrBase = Left(Addr, P1 - 1)
    End If
    lAddrBit = AddrBit
    lHigh = lLow + (nElems - 1)
End Sub

Public Property Get AddrBase() As String
    AddrBase = sAddrBase
End Property

Public Property Get IsLogix() As Boolean
    IsLogix = bIsLogix
End Property

Public Property Get Word() As Long
    Word = lLow
End Property

Public Property Get Bit() As Long
    Bit = lAddrBit
End Property

Public Property Get LastWord() As Long
    LastWord = lHigh
End Property

Public Property Get TypeFamily() As String
    Select Case sPT_TYPE
    Case "BOOL"
        TypeFamily = "BOOL"
    Case "DINT", "SINT", "INT", "UDINT", "UINT", "USINT"
        TypeFamily = "INT"
    Case "REAL"
        TypeFamily = "REAL"
    Case "STRING", "STRING_1600", "STRING_20", "STRING_75", "STRING_8", "STRING_80"
        TypeFamily = "STRING"
    Case Else
        Stop
    End Select
End Property

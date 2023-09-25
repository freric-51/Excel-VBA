VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArrayByCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const sSeparator = "|"
'Dim cData As Collection
Dim cData() As String
Dim iLastPos As Long

Public Property Get DataM(ID As Long) As String()
    DataM = Split(Trim(cData(ID)), sSeparator)
End Property

Public Property Let DataSep(ByVal Value As String)
    'cData.Add Trim(Value)
    Dim L As Long
    L = 130000
    If iLastPos = UBound(cData) Then
        ReDim Preserve cData(0 To L + UBound(cData))
        If 1 = Workbooks.Count Then
            Workbooks(1).Save
        End If
    End If
    If cData(iLastPos) <> "" Then iLastPos = iLastPos + 1
    cData(iLastPos) = Trim(Value)
End Property

Public Property Get Max() As Long
    Max = iLastPos 'cData.Count
End Property

Public Property Get Separator() As String
    Separator = sSeparator
End Property

Private Sub Class_Initialize()
    'Set cData = New Collection
    Dim L As Long
    L = 100000
    ReDim cData(0 To L)
    iLastPos = 0
End Sub
Private Sub Class_Terminate()
    'Set cData = Nothing
    Dim L As Long
    L = 0
    ReDim cData(L To L)
    
    iLastPos = 0
End Sub

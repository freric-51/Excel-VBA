VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public R As Long
Public G As Long
Public B As Long
Private loc_ID As Long

'=============================================================
Public Property Let ID(Valor As Long)
    Dim c As Currency, v As Currency, Resto As Long

    loc_ID = Valor

    Resto = loc_ID - 19 * (loc_ID \ 19)

    Select Case Resto
    Case 0
        v = RGB(0, 0, 90)
    Case 1
        v = RGB(0, 90, 0)
    Case 2
        v = RGB(90, 0, 0)
    Case 3
        v = RGB(0, 90, 90)
    Case 4
        v = RGB(90, 0, 90)
    Case 5
        v = RGB(90, 90, 0)
    Case 6
        v = RGB(0, 0, 180)
    Case 7
        v = RGB(0, 180, 0)
    Case 8
        v = RGB(180, 0, 0)
    Case 9
        v = RGB(0, 180, 180)
    Case 10
        v = RGB(180, 0, 180)
    Case 11
        v = RGB(180, 180, 0)
    Case 12
        v = RGB(0, 0, 250)
    Case 13
        v = RGB(0, 250, 0)
    Case 14
        v = RGB(250, 0, 0)
    Case 15
        v = RGB(0, 250, 250)
    Case 16
        v = RGB(250, 0, 250)
    Case 17
        v = RGB(250, 250, 0)
    Case 18
        v = RGB(90, 180, 90)
    Case Else
        Stop    'error
    End Select

    c = 256@ * 256@    ' * 256@

    R = Int(v \ c)
    v = v - R * c

    c = 256@    '* 256@
    G = Int(v \ c)
    v = v - G * c

    c = 1    '256@
    B = Int(v \ c)
    v = v - B * c

    If v <> 0 Then Stop    'found a error

End Property

Public Property Get ID() As Long
    ID = loc_ID
End Property

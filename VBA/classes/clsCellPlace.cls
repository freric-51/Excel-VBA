VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCellPlace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'/**
' * first release in 9/9/2013
' */
Public FirstColumn As String
Public LastColumn As String
Public FirstLine As Long
Public LastLine As Long

Public Property Let Range(Info As String)
    Dim i As Long, j As Long
    Dim t1 As String, t2 As String
    i = InStr(Info, ":")
    t1 = Left(Info, i - 1)
    t2 = Mid(Info, i + 1)
    '**
    FirstColumn = ""
    For i = 1 To Len(t1)
        If Not IsNumeric(Mid(t1, i, 1)) Then
            FirstColumn = FirstColumn & Mid(t1, i, 1)
        Else
            FirstLine = Mid(t1, i)
            Exit For
        End If
    Next i

    '**
    LastColumn = ""
    For i = 1 To Len(t2)
        If Not IsNumeric(Mid(t2, i, 1)) Then
            LastColumn = LastColumn & Mid(t2, i, 1)
        Else
            LastLine = Mid(t2, i)
            Exit For
        End If
    Next i
End Property

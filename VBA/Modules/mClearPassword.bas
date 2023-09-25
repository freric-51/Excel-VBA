Attribute VB_Name = "mClearPassword"
Option Explicit

Private Const PAGE_EXECUTE_READWRITE = &H40

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
                               (Destination As Long, Source As Long, ByVal length As Long)

Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Long, _
                                                        ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
                                                        ByVal lpProcName As String) As Long

Private Declare Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" (ByVal hInstance As Long, _
                                                                              ByVal pTemplateName As Long, ByVal hWndParent As Long, _
                                                                              ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer

Dim HookBytes(0 To 5) As Byte
Dim OriginBytes(0 To 5) As Byte
Dim pFunc As Long
Dim Flag As Boolean

Public Sub unprotected()
    If Hook Then
        PlayWAV "chord.wav"
        MsgBox "VBA Project is unprotected!", vbInformation, "*****"
    End If
End Sub

Private Function GetPtr(ByVal Value As Long) As Long
    GetPtr = Value
End Function

Public Sub RecoverBytes()
    If Flag Then MoveMemory ByVal pFunc, ByVal VarPtr(OriginBytes(0)), 6
End Sub

Public Function Hook() As Boolean
    Dim TmpBytes(0 To 5) As Byte
    Dim p As Long
    Dim OriginProtect As Long

    Hook = False

    pFunc = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA")


    If VirtualProtect(ByVal pFunc, 6, PAGE_EXECUTE_READWRITE, OriginProtect) <> 0 Then

        MoveMemory ByVal VarPtr(TmpBytes(0)), ByVal pFunc, 6
        If TmpBytes(0) <> &H68 Then

            MoveMemory ByVal VarPtr(OriginBytes(0)), ByVal pFunc, 6

            p = GetPtr(AddressOf MyDialogBoxParam)

            HookBytes(0) = &H68
            MoveMemory ByVal VarPtr(HookBytes(1)), ByVal VarPtr(p), 4
            HookBytes(5) = &HC3

            MoveMemory ByVal pFunc, ByVal VarPtr(HookBytes(0)), 6
            Flag = True
            Hook = True
        End If
    End If
End Function

Private Function MyDialogBoxParam(ByVal hInstance As Long, _
                                  ByVal pTemplateName As Long, ByVal hWndParent As Long, _
                                  ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer
    If pTemplateName = 4070 Then
        MyDialogBoxParam = 1
    Else
        RecoverBytes
        MyDialogBoxParam = DialogBoxParam(hInstance, pTemplateName, _
                                          hWndParent, lpDialogFunc, dwInitParam)
        Hook
    End If
End Function

Public Sub WorksheetPasswordBreaker()
    Dim i As Integer, j As Integer, k As Integer
    Dim L As Integer, m As Integer, n As Integer
    Dim i1 As Integer, i2 As Integer, i3 As Integer
    Dim i4 As Integer, i5 As Integer, i6 As Integer
    'Breaks worksheet password protection.
    'On Error Resume Next
    For i = 32 To 126
        For j = 32 To 126
            For k = 32 To 126
                For L = 32 To 126
                    For m = 32 To 126
                        For i1 = 32 To 126
                            For i2 = 32 To 126
                                For i3 = 32 To 126
                                    For i4 = 32 To 126
                                        For i5 = 32 To 126
                                            For i6 = 32 To 126
                                                For n = 32 To 126
                                                    ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(L) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
                                                    If ActiveSheet.ProtectContents = False Then
                                                        PlayWAV "chord.wav"
                                                        MsgBox "One usable password is " & Chr(i) & Chr(j) & Chr(k) & Chr(L) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
                                                        Exit Sub
                                                    End If
                                                Next
                                            Next
                                        Next
                                    Next
                                Next
                            Next
                        Next
                    Next
                Next
            Next
        Next
    Next
End Sub

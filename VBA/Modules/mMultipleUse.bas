Attribute VB_Name = "mMultipleUse"
Option Explicit
'##### MM/DD/YYYY #####
'## /2018 ##
'01/17  GetTruePath
'02/01  ClearGlobalVariables
'03/02  FindColumnByRange QAS. 'rever QAS.FindColumnByRange
'03/02  GetColumnByTable
'03/05  NewSheet +iPos + ColorTab
'03/06  AntiWord AntiByte AntiLog
'03/08  FixedLengthParameter
'03/17  IsPainted
'03/26  OneDrive
'05/17  multiple PDUs in import data
'07/05  SplitFixedAddress
'09/28  HighlightDuplicated

'## /2017 ##
'01/03  Getlnkpath GetTruePath
'02/13  ImportRSLinx
'02/17  HighlightDuplicated in QAS DTR values
'02/22  doevents in time
'02/24  desmonta ponto ANDON
'03/07  trim rtrim
'03/08  UTF8
'09/22  reverse string
'12/15  OPC Cimplicity code

'## /2016 ##
'04/08  ReCalc
'03/16  chr 160
'03/11  WorksheetExists
'03/02  InitVBs
'02/02  Log > LogSheet
'05/06  mColorAnima
'05/12  LimpaBeiradas
'05/16  FaixaBits
'05/17  MatrixReady
'06/29  GoodForXML
'07/18  RemoveNotAllowedSheetName
'07/20  IncFormula
'07/25  FindCorrelatedText
'08/02  Main_Workbook_Path
'08/05  removeVBCRLF 150
'08/11  ImportIDT
'08/13  NumberOfDimensions
'08/16  ImportExcel
'09/01  MaxWidth + xxxColorIndex + WorkbookExist
'09/05  HighlightDuplicated
'09/13  ShowFirstSheet
'09/15  xla
'09/26  ExisteNaColecao'
'
Private DecLetter(0 To 677) As String
Public Const LastPossibleColumn = 676 'YZ
Public Const MaxQasWords = 80 * 16

Public Const MaxLenAlarm = 80
Public Const MaxLenDesc = 40
Public Const MAXINT As Integer = (2 ^ 15) - 1
Public Const MAXLONG = (2 ^ 31) - 1
'
Public Const NotFoundYet = "-"
Public ColColumn As New Collection
'
Public Const LogTab = "Log"
Public Enum enLog
    Register
    Fault
    Warning
End Enum
'
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_FILENAME = &H20000
'
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'
Private mColorAnima(0 To 24, 0 To 2) As String
'
Public Main_Workbook_Path As String
Public Main_Workbook_Name As String
'
'GPMC Lengths
Public Const cLengthSysCode = 1
Public Const cLengthPointType = 2
'Resource
Public Const cLengthArea = 2
Public Const cLengthSubArea = 5
Public Const cLengthEquipmentType = 2
Public Const cLengthEquipmentIdentifier = 4
'
Public Const cLengthDescription = 12
'QAS
Public Const cLengthTeam = 2
Public Const cLengthDepartment = 2
Public Const cLengthSection = 2
Public Const cLengthOperation = 4
'
Public Enum ePT_TYPE
    T_BOOL = 0
    T_INT
    T_UINT
    T_UDINT
    T_REAL
    T_STRING
    T_STRING20
    T_STRING8
    T_STRING80
End Enum

Public Enum UpdateActions
    EnableFull
    NoneWork
    ReanimateScreen 'Speed Down
    FreezeScreen 'Speed Up
End Enum
Public LastSheetUpdates As UpdateActions

Public Enum SyncActions
    Freeze
    Release
End Enum

Public Enum LogActions
    A_New = 0
    A_Write
    A_Read
    A_OptimalColumnWidth
    A_RemoveDuplicates
End Enum

'GE Color Index
Public Enum GEColorIndex
    Black = 0
    Red
    Lime
    Blue
    Maroon
    Green
    Purple
    white
    Teal
    Gray
    Olive
    Dark
    Rose
    Fuchsia
    Aqua
    Yellow
End Enum

'Excel Color Index
Public Enum ExcelColorIndex
    NotUsed = -1
    NoFill = 0
    Black = 1
    white = 2
    Gray25 = 15
    Gray50 = 16

    Red = 3
    DarkRed = 9
    LightRed = 38

    Green = 10
    DarkGreen = 51
    LightGreen = 4

    Blue = 5
    DarkBlue = 11
    LightBlue = 41

    Yellow = 6
    DarkYellow = 12
    LightYellow = 36

    Pink = 7
    DarkPink = 13
    LightPink = 39

    Cyan = 8
    DarkCyan = 14
    LightCyan = 42

    Brown = 53
    Orange = 46
End Enum

Public Enum RGBColorIndex
    Black = 0
    Maroon = 128
    DarkRed = 139
    Red = 255
    OrangeRed = 17919
    DarkGreen = 25600
    Green = 32768
    Olive = 32896
    DarkOrange = 36095
    Orange = 42495
    Gold = 55295
    LawnGreen = 64636
    Lime = 65280
    Chartreuse = 65407
    Yellow = 65535
    DarkGoldenrod = 755384
    Goldenrod = 2139610
    FireBrick = 2237106
    ForestGreen = 2263842
    OliveDrab = 2330219
    Brown = 2763429
    Sienna = 2970272
    DarkOliveGreen = 3107669
    GreenYellow = 3145645
    LimeGreen = 3329330
    YellowGreen = 3329434
    Crimson = 3937500
    Peru = 4163021
    Tomato = 4678655
    DarkSlateGray = 5197615
    DarkSlateGrey = 5197615
    Coral = 5275647
    SeaGreen = 5737262
    IndianRed = 6053069
    SandyBrown = 6333684
    DimGray = 6908265
    DimGrey = 6908265
    DarkKhaki = 7059389
    PaleGoldenrod = 7071982
    MidnightBlue = 7346457
    MediumSeaGreen = 7451452
    Salmon = 7504122
    DarkSalmon = 8034025
    LightSalmon = 8036607
    SpringGreen = 8388352
    Navy = 8388608
    NavyBlue = 8388608
    Purple = 8388736
    Teal = 8421376
    Gray = 8421504
    Grey = 8421504
    LightCoral = 8421616
    Indigo = 8519755
    MediumVioletRed = 8721863
    BurlyWood = 8894686
    DarkBlue = 9109504
    DarkMagenta = 9109643
    DarkSlateBlue = 9125192
    DarkCyan = 9145088
    LightCyan = 9145088
    Tan = 9221330
    Khaki = 9234160
    RosyBrown = 9408444
    DarkSeaGreen = 9419919
    SlateGray = 9470064
    LightGreen = 9498256
    DeepPink = 9639167
    PaleVioletRed = 9662683
    PaleGreen = 10025880
    LightSlateGray = 10061943
    MediumSpringGreen = 10156544
    CadetBlue = 10526303
    DarkGray = 11119017
    DarkGrey = 11119017
    LightSeaGreen = 11186720
    MediumAquamarine = 11206502
    NavajoWhite = 11394815
    Wheat = 11788021
    HotPink = 11823615
    SteelBlue = 11829830
    Moccasin = 11920639
    PeachPuff = 12180223
    Silver = 12632256
    LightPink = 12695295
    Bisque = 12903679
    Pink = 13353215
    DarkOrchid = 13382297
    MediumTurquoise = 13422920
    MediumBlue = 13434880
    SlateBlue = 13458026
    BlanchedAlmond = 13495295
    LemonChiffon = 13499135
    Turquoise = 13688896
    DarkTurquoise = 13749760
    LightGoldenrodYellow = 13826810
    DarkViolet = 13828244
    MediumOrchid = 13850042
    LightGray = 13882323
    LightGrey = 13882323
    Aquamarine = 13959039
    PapayaWhip = 14020607
    Orchid = 14053594
    AntiqueWhite = 14150650
    Thistle = 14204888
    MediumPurple = 14381203
    Gainsboro = 14474460
    Beige = 14480885
    Cornsilk = 14481663
    Plum = 14524637
    LightSteelBlue = 14599344
    LightYellow = 14745599
    RoyalBlue = 14772545
    MistyRose = 14804223
    BlueViolet = 14822282
    LightBlue = 15128749
    PowderBlue = 15130800
    Linen = 15134970
    OldLace = 15136253
    SkyBlue = 15453831
    CornflowerBlue = 15570276
    MediumSlateBlue = 15624315
    Violet = 15631086
    PaleTurquoise = 15658671
    Seashell = 15660543
    FloralWhite = 15792895
    Honeydew = 15794160
    Ivory = 15794175
    LavenderBlush = 16118015
    WhiteSmoke = 16119285
    LightSkyBlue = 16436871
    Lavender = 16443110
    Snow = 16448255
    MintCream = 16449525
    Blue = 16711680
    Fuchsia = 16711935
    DodgerBlue = 16748574
    DeepSkyBlue = 16760576
    AliceBlue = 16775408
    GhostWhite = 16775416
    Aqua = 16776960
    Azure = 16777200
    white = 16777215
End Enum

''' WinApi function that maps a UTF-16 (wide character) string to a new character string
Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
                                             ByVal CodePage As Long, _
                                             ByVal dwFlags As Long, _
                                             ByVal lpWideCharStr As Long, _
                                             ByVal cchWideChar As Long, _
                                             ByVal lpMultiByteStr As Long, _
                                             ByVal cbMultiByte As Long, _
                                             ByVal lpDefaultChar As Long, _
                                             ByVal lpUsedDefaultChar As Long) As Long

' CodePage constant for UTF-8
Private Const CP_UTF8 = 65001

'#####

Public Sub PlayWAV(ByVal WAVFile As String)
    Dim TotalWAVFile As String
    TotalWAVFile = Environ("SystemRoot") & "\media\" & WAVFile
    If Dir(TotalWAVFile) <> "" Then
        Call PlaySound(TotalWAVFile, 0&, SND_ASYNC Or SND_FILENAME)
    End If
End Sub

Function RemoveNotAllowedSheetName(ByVal old As String) As String
    Static ID As Long
    old = Replace(old, ":", "")
    old = Replace(old, "\", "")
    old = Replace(old, "/", "")
    old = Replace(old, "?", "")
    old = Replace(old, "*", "")
    old = Replace(old, "[", "")
    old = Replace(old, "]", "")
    old = Trim(old)
    If Len(old) > 31 Then
        old = Left(old, 27)
        old = Trim(old)
        old = old & "-" & ID
        ID = ID + 1
        If ID > 99 Then ID = 0
    End If
    RemoveNotAllowedSheetName = old
End Function

Public Sub BreakLinks(ByRef wb As Workbook)
    'http://msdn.microsoft.com/en-us/library/office/ff194915.aspx
    Dim Links As Variant
    Dim i As Long
    On Error Resume Next
    Links = wb.LinkSources(Type:=xlLinkTypeExcelLinks)
    On Error GoTo 0
    If Not IsEmpty(Links) Then
        For i = 1 To UBound(Links)
            wb.BreakLink Name:=Links(i), Type:=xlLinkTypeExcelLinks
            EnableEvents
        Next i
    End If
End Sub

Function LastDay(dtDate As Date)
    Dim dtBase As Date
    'mes futuro
    dtBase = DateAdd("m", 1, DateSerial(Year(dtDate), Month(dtDate), -1))
    'menos 1 dia
    LastDay = Day(DateAdd("D", -1, dtBase))
End Function
'
Public Function IdToColumn(ByVal ID As Long) As String
    IdToColumn = DecLetter(ID)
    If IdToColumn = "" Then
        Call Fill_IdToColumn
        IdToColumn = DecLetter(ID)
    End If
    If IdToColumn = "" Then Stop 'There is any issue here
End Function

Private Sub Fill_IdToColumn()
    Dim i1 As Long
    For i1 = 1 To 1 + 26 * 26
        DecLetter(i1) = IdToColumn_OBO(i1)
    Next i1
End Sub

Private Function IdToColumn_OBO(ByVal ID As Long) As String
    If ID > LastPossibleColumn Then
        IdToColumn_OBO = ""
        Exit Function
    End If
    '
    Dim n(0 To 2) As Long, i As Long, p As Long
    IdToColumn_OBO = ""
    ID = ID - 1
    'convert to 26 base
    For i = 2 To 0 Step -1
        p = 26 ^ i
        n(i) = ID \ (p)
        ID = ID - (n(i) * (p))
    Next i
    If ID <> 0 Then Stop 'error found

    'transfer to letters
    For i = 2 To 0 Step -1
        If i > 0 And n(i) > 0 Then IdToColumn_OBO = IdToColumn_OBO & Chr(Asc("A") + n(i) - 1)
        If i = 0 Then IdToColumn_OBO = IdToColumn_OBO & Chr(Asc("A") + n(i))
    Next i
End Function
'
Public Sub DestroySpreadsheet(ByVal WorkName As String, ByVal SheetName As String)
    Application.DisplayAlerts = False
    On Error Resume Next
    Workbooks(WorkName).Sheets(SheetName).Delete
    Sleep 500
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

Public Sub SetMainWorkbookName()
    Dim i1 As Long
    If Main_Workbook_Path = "" Then
        For i1 = 1 To Workbooks.Count
            If Workbooks.Item(i1).Name = ActiveWorkbook.Name Then
                Main_Workbook_Path = Workbooks.Item(i1).Path
                Main_Workbook_Name = Workbooks.Item(i1).Name
            End If
        Next i1
    End If
    If Main_Workbook_Path = "" Then Stop
End Sub

Public Function WorkbookExist(ByVal sWorkbookName As String) As Boolean
    Dim ok As Boolean, sPathFull As String
    '/**
    ' * verificar nome de arquivo trocado, renomeado, apagado, ...!
    ' */
    Call InitVBs
    sWorkbookName = Replace(sWorkbookName, Main_Workbook_Path, "")

    Dim obFile As New clsFile
    obFile.Arq = sWorkbookName
    obFile.BasePath = Main_Workbook_Path
    sPathFull = obFile.FullPath

    On Error Resume Next
    'OK = Dir(Main_Workbook_Path & "\" & sWorkbookName) <> ""
    ok = Dir(obFile.FullPath & "\" & obFile.FileAndExtension) <> ""
    If Err.Number <> 0 Then ok = False
    On Error GoTo 0

    If Not ok Then
        PlayWAV "chord.wav"
        'MsgBox "Arquivo " & sWorkbookName & " inexistente!"
        LogSheet "WorkbookExist", sWorkbookName, "Arquivo inexistente!", getLogType(enLog.Register), LogActions.A_Write
        WorkbookExist = False
        Exit Function
    End If

    WorkbookExist = True
End Function

Public Function FillChr(ByVal nQ As Long, ByVal sC As String) As String
    Dim nCount As Long
    sC = Left(sC & " ", 1)
    For nCount = 1 To nQ
        FillChr = FillChr & sC
    Next nCount
End Function

Public Function FixedLengthParameter(ByRef Par As String, ByRef Size As Long) As String
    Dim ParBkp As String
    ParBkp = Par
    '
    If Len(ParBkp) > Size Then
        'Stop 'never did before
        If "0" = (Left(ParBkp, 1)) Then
            ParBkp = Mid(ParBkp, 2)
        ElseIf "_" = (Left(ParBkp, 1)) Then
            ParBkp = Mid(ParBkp, 2)
        ElseIf space(1) = (Left(ParBkp, 1)) Then
            ParBkp = Mid(ParBkp, 2)
        Else
            Stop 'it is not possible
        End If
        ParBkp = FixedLengthParameter(ParBkp, Size)
        '
    ElseIf Len(ParBkp) < Size Then
        If IsNumeric(Left(ParBkp, 1)) Then
            ParBkp = FillChr(Size, "0") & ParBkp
        Else
            ParBkp = FillChr(Size, "_") & ParBkp
        End If
        ParBkp = Right(ParBkp, Size)
    End If
    FixedLengthParameter = ParBkp
End Function

Public Function OnlyNumberLeterUnderscore(ByVal sInput As String) As String
    Dim iLoop As Long
    'Good = 48-57 65-90 _ 97-122

    sInput = Trim(sInput)

    For iLoop = 1 To 47
        sInput = Replace(sInput, Chr(iLoop), "_")
    Next iLoop
    For iLoop = 58 To 64
        sInput = Replace(sInput, Chr(iLoop), "_")
    Next iLoop
    For iLoop = 91 To 96
        sInput = Replace(sInput, Chr(iLoop), "_")
    Next iLoop
    For iLoop = 123 To 255
        sInput = Replace(sInput, Chr(iLoop), "_")
    Next iLoop
    OnlyNumberLeterUnderscore = sInput
End Function

Public Function ReplaceAcentos(ByVal sInput As String) As String
    sInput = Replace(sInput, "�", "A")
    sInput = Replace(sInput, "�", "a")

    sInput = Replace(sInput, "�", "E")
    sInput = Replace(sInput, "�", "e")

    sInput = Replace(sInput, "�", "I")
    sInput = Replace(sInput, "�", "i")

    sInput = Replace(sInput, "�", "O")
    sInput = Replace(sInput, "�", "o")

    sInput = Replace(sInput, "�", "U")
    sInput = Replace(sInput, "�", "u")
    '=================================
    sInput = Replace(sInput, "�", "A")
    sInput = Replace(sInput, "�", "a")

    sInput = Replace(sInput, "�", "O")
    sInput = Replace(sInput, "�", "o")

    sInput = Replace(sInput, "�", "N")
    sInput = Replace(sInput, "�", "n")
    '=================================
    sInput = Replace(sInput, "�", "C")
    sInput = Replace(sInput, "�", "c")
    '=================================
    sInput = Replace(sInput, "�", "A")
    sInput = Replace(sInput, "�", "a")

    sInput = Replace(sInput, "�", "E")
    sInput = Replace(sInput, "�", "e")

    sInput = Replace(sInput, "�", "I")
    sInput = Replace(sInput, "�", "i")

    sInput = Replace(sInput, "�", "O")
    sInput = Replace(sInput, "�", "o")

    sInput = Replace(sInput, "�", "U")
    sInput = Replace(sInput, "�", "u")
    '=================================
    sInput = Replace(sInput, "�", "A")
    sInput = Replace(sInput, "�", "a")

    sInput = Replace(sInput, "�", "E")
    sInput = Replace(sInput, "�", "e")

    sInput = Replace(sInput, "�", "I")
    sInput = Replace(sInput, "�", "i")

    sInput = Replace(sInput, "�", "O")
    sInput = Replace(sInput, "�", "o")

    sInput = Replace(sInput, "�", "U")
    sInput = Replace(sInput, "�", "u")
    '=================================
    sInput = Replace(sInput, "�", "A")
    sInput = Replace(sInput, "�", "a")

    sInput = Replace(sInput, "�", "E")
    sInput = Replace(sInput, "�", "e")

    sInput = Replace(sInput, "�", "I")
    sInput = Replace(sInput, "�", "i")

    sInput = Replace(sInput, "�", "O")
    sInput = Replace(sInput, "�", "o")

    sInput = Replace(sInput, "�", "U")
    sInput = Replace(sInput, "�", "u")
    '=================================
    'first second ... ��
    sInput = Replace(sInput, Chr(170), "a")
    sInput = Replace(sInput, Chr(186), "o")

    sInput = Replace(sInput, Chr(176), ".")
    sInput = Replace(sInput, Chr(185), "1")
    sInput = Replace(sInput, Chr(178), "2")
    sInput = Replace(sInput, Chr(179), "3")
    '=================================

    ReplaceAcentos = sInput
End Function

Public Function GoodForXML(ByVal sInput As String) As String
    Dim i1 As Long, iChar As Long, sFin As String

    sInput = removeVBCRLF(sInput)
    sInput = ReplaceAcentos(sInput)
    sInput = Replace(sInput, ",", ".")


    While InStr(sInput, "--") > 0
        sInput = Replace(sInput, "--", "-")
        sInput = Trim(sInput)
        EnableEvents
    Wend
    '</
    'sInput = Replace(sInput, "-", " - ")

    sFin = ""
    For i1 = 1 To Len(sInput)
        iChar = Asc(Mid(sInput, i1, 1))
        Select Case iChar
            'added 45 58 95 06/29/2016
        Case 32, 40, 41, 46, 45, 47, 48 To 57, 58, 65 To 90, 95, 97 To 122
            sFin = sFin & Chr(iChar)
        Case Else
            sFin = sFin & space(1)
        End Select
        EnableEvents
    Next i1
    sInput = sFin

    While InStr(sInput, space(2)) > 0
        sInput = Replace(sInput, space(2), space(1))
        sInput = Trim(sInput)
        EnableEvents
    Wend

    sInput = Replace(sInput, "</", ".")
    sInput = Replace(sInput, "<", ".")
    sInput = Replace(sInput, ">", ".")
    While InStr(sInput, "..") > 0
        sInput = Replace(sInput, "..", ".")
        sInput = Trim(sInput)
        EnableEvents
    Wend

    sInput = Trim(sInput)
    GoodForXML = sInput
End Function

Public Sub ClearColumnData(sSheet As String, sCol As String)
    ReCalc
    Dim R As Range, Cell As Range, LastRow As Long, bFound As Boolean
    With Workbooks(Main_Workbook_Name).Sheets(sSheet)
        .Activate
        LastRow = .Cells(.Rows.Count, sCol).End(xlUp).row
        Set R = .Range(sCol & "1:" & sCol & LastRow)
        R.Select
        For Each Cell In R
            Cell.Value = ""
            EnableEvents
        Next Cell
    End With
End Sub

Public Sub HighlightDuplicated(sSheet As String, sCol As String, bSkipBlank As Boolean)
    ReCalc
    Dim R As Range, Cell As Range, LastRow As Long, bFound As Boolean, bPaint As Boolean
    bFound = False
    With Workbooks(Main_Workbook_Name).Sheets(sSheet)
        .Activate
        LastRow = .Cells(.Rows.Count, sCol).End(xlUp).row
        KeepZeroLeftbyColumn sSheet, sCol, 1, LastRow
        Set R = .Range(sCol & "1:" & sCol & LastRow)
        R.Select
        For Each Cell In R
            EnableEvents
            If WorksheetFunction.CountIf(R, Cell.Value) > 1 Then
                bPaint = True
                If bSkipBlank And ("" = Cell.Value) Then bPaint = False
                If bPaint Then
                    Cell.Interior.ColorIndex = ExcelColorIndex.Red
                    bFound = True
                    DoEvents
                End If
            End If
        Next Cell
        If bFound Then
            .Range(sCol & 1).Interior.ColorIndex = ExcelColorIndex.Pink
        Else
            .Range(sCol & 1).Interior.ColorIndex = ExcelColorIndex.LightGreen
        End If
    End With
    'SendKeys "%HLHD{ENTER}"
End Sub

Public Sub Duplicated(ByRef wInfo As Worksheet, ByVal sColumnCompare As String, ByVal nStart As Long, ByVal nEnd As Long, sWorkSheetTemp As String, sColId, sColAlm, sColDesc)
    Dim nLoop As Long, c As Range, sWhereToFind As String, nReportLine As Long

    nReportLine = 1
    If "" = Worksheets(sWorkSheetTemp).Range("A" & nReportLine) Then
        Worksheets(sWorkSheetTemp).Range("A" & nReportLine) = "Point"
        Worksheets(sWorkSheetTemp).Range("B" & nReportLine) = "Addrees Base"
        Worksheets(sWorkSheetTemp).Range("C" & nReportLine) = "Address Found"
        Worksheets(sWorkSheetTemp).Range("D" & nReportLine) = "ALM_MSG"
        Worksheets(sWorkSheetTemp).Range("E" & nReportLine) = "DESC"
    End If

    While "" <> Worksheets(sWorkSheetTemp).Range("A" & nReportLine)
        nReportLine = nReportLine + 1
        EnableEvents
    Wend

    nLoop = nStart
    While (nEnd <> nLoop)
        If (nStart <> nLoop) Then
            sWhereToFind = sColumnCompare & nStart & ":" & sColumnCompare & nLoop - 1
            'GoSub FindDup
        End If

        If (nEnd <> nLoop) Then
            sWhereToFind = sColumnCompare & nLoop + 1 & ":" & sColumnCompare & nEnd
            GoSub FindDup
        End If

        nLoop = nLoop + 1
        EnableEvents
    Wend

    Exit Sub

FindDup:
    Set c = wInfo.Range(sWhereToFind).Find(wInfo.Range(sColumnCompare & nLoop), LookIn:=xlValues, MatchCase:=True, LookAt:=XlLookAt.xlWhole)
    If Not c Is Nothing Then
        Worksheets(sWorkSheetTemp).Range("A" & nReportLine) = wInfo.Range(sColId & nLoop)
        Worksheets(sWorkSheetTemp).Range("B" & nReportLine) = "$" & sColumnCompare & "$" & nLoop
        Worksheets(sWorkSheetTemp).Range("C" & nReportLine) = c.Address
        Worksheets(sWorkSheetTemp).Range("D" & nReportLine) = wInfo.Range(sColAlm & nLoop)
        Worksheets(sWorkSheetTemp).Range("E" & nReportLine) = wInfo.Range(sColDesc & nLoop)
        nReportLine = nReportLine + 1
    End If
    Return
End Sub

Public Sub NewSheet(sSheet As String, Optional iPos As Long = -1, Optional iTabColor As ExcelColorIndex = -1)
    Call InitVBs
    DestroySpreadsheet Main_Workbook_Name, sSheet
    If iPos < 1 Then
        iPos = Sheets.Count
    ElseIf iPos > Sheets.Count Then
        iPos = Sheets.Count
    End If
    Worksheets.add Count:=1, After:=Sheets(iPos)
    'New = 1+
    Sheets(iPos + 1).Name = sSheet
    Workbooks(Main_Workbook_Name).Sheets(sSheet).Cells.Font.Name = "Consolas"
    '
    If Not IsMissing(iTabColor) Then
        If iTabColor <> -1 Then
            Workbooks(Main_Workbook_Name).Sheets(sSheet).Tab.ColorIndex = iTabColor
        End If
    End If
End Sub

Public Function WorksheetExists(ByVal sNameSheet As String) As Boolean
    WorksheetExists = Evaluate("ISREF('" & sNameSheet & "'!A1)")
End Function

Public Function getLogType(eValue As enLog) As String
    Select Case eValue
        Case Register: getLogType = "REGISTER"
        Case Fault: getLogType = "FAULT"
        Case Warning: getLogType = "WARNING"
    End Select
End Function
Public Sub LogSheet(ByVal Msg As String, ByVal Ref1 As String, ByVal Ref2 As String, ByVal Ref3 As String, ByVal Action As LogActions)
    Dim nReportLine As Long, sRet As String
    Dim LastRow As Long

    If Not WorksheetExists(LogTab) Then
        NewSheet LogTab, 2
    End If

    Select Case Action
    Case LogActions.A_New
        NewSheet LogTab, 2
    Case LogActions.A_Write
        nReportLine = Worksheets(LogTab).Range("A1").Cells(Rows.Count, 1).End(xlUp).row
        If "" <> Worksheets(LogTab).Range("A" & nReportLine) Then nReportLine = nReportLine + 1
        If nReportLine = 1 Then
            Worksheets(LogTab).Range("A" & nReportLine) = "Message"
            Worksheets(LogTab).Range("B" & nReportLine) = "'01"
            Worksheets(LogTab).Range("C" & nReportLine) = "'02"
            Worksheets(LogTab).Range("D" & nReportLine) = "'03"
            Worksheets(LogTab).Range("E" & nReportLine) = "Time"
            nReportLine = nReportLine + 1
        End If
        If Left(Msg, 1) <> SQ Then Msg = SQ & Msg
        If Left(Ref1, 1) <> SQ Then Ref1 = SQ & Ref1
        If Left(Ref2, 1) <> SQ Then Ref2 = SQ & Ref2
        If Left(Ref3, 1) <> SQ Then Ref3 = SQ & Ref3

        Worksheets(LogTab).Range("A" & nReportLine) = Msg
        Worksheets(LogTab).Range("B" & nReportLine) = Ref1
        Worksheets(LogTab).Range("C" & nReportLine) = Ref2
        Worksheets(LogTab).Range("D" & nReportLine) = Ref3
        Worksheets(LogTab).Range("E" & nReportLine) = Format(Now, "ttttt")
    Case LogActions.A_OptimalColumnWidth
        MaxAutoWidth LogTab, 40
        'Worksheets(LogTab).Columns("A:ZZ").AutoFit
    Case LogActions.A_RemoveDuplicates
        'copiar para outra coluna
        Worksheets(LogTab).Select
        Worksheets(LogTab).Columns("A:B").Copy Destination:=Worksheets(LogTab).Columns("G:H")
        LastRow = Worksheets(LogTab).Cells(Worksheets(LogTab).Rows.Count, "G").End(xlUp).row
        Worksheets(LogTab).Columns("G:H").RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
    End Select
    If nReportLine > 100000 Then
        For nReportLine = 1 To 10
            Worksheets(LogTab).Range("A2").EntireRow.Delete
            EnableEvents
        Next nReportLine
    End If
End Sub

Public Function removeVBCRLF(ByVal sFrase As String, Optional sRep As String) As String
    Dim nr As Long
    If sRep <> "" Then
        sFrase = Replace(sFrase, vbCrLf, sRep)
        sFrase = Replace(sFrase, vbCr, sRep)
        sFrase = Replace(sFrase, vbLf, sRep)
    End If

    For nr = 0 To 31
        sFrase = Replace(sFrase, Chr(nr), space(1))
    Next nr
    nr = 129: sFrase = Replace(sFrase, Chr(nr), space(1))
    nr = 141: sFrase = Replace(sFrase, Chr(nr), space(1))
    nr = 143: sFrase = Replace(sFrase, Chr(nr), space(1))
    nr = 144: sFrase = Replace(sFrase, Chr(nr), space(1))
    nr = 150: sFrase = Replace(sFrase, Chr(nr), "-")
    nr = 160: sFrase = Replace(sFrase, Chr(nr), space(1))
    'only 1 space
    While InStr(sFrase, space(2)) > 0
        sFrase = Replace(sFrase, space(2), space(1))
        EnableEvents
    Wend
    'sFrase = Trim(sFrase)
    removeVBCRLF = sFrase
End Function

Public Function DQ() As String
    DQ = Chr$(34)    'Double quote (") character
End Function
Public Function SQ() As String
    SQ = Chr$(39)    'single quotes (') character
End Function

Public Function GetKey(ByVal sALL As String, ByVal sStart As String, ByVal sEnd As String) As String
    Dim iPosINI As Long, iPosEND As Long
    'Example: <G K="An approved, un-used, Scheduled Non-Load Event of type ">
    '===== 2017 =====
    sStart = UCase(sStart)
    sEnd = UCase(sEnd)
    sALL = Replace(sALL, sStart, sStart, , , vbTextCompare)
    sALL = Replace(sALL, sEnd, sEnd, , , vbTextCompare)
    If InStr(sALL, vbCr) Then Stop
    If InStr(sALL, vbLf) Then Stop
    If InStr(sALL, vbCrLf) Then Stop
    '===== 2017 =====
    iPosINI = InStr(sALL, sStart)
    If iPosINI > 0 Then
        iPosINI = iPosINI + Len(sStart)
        iPosEND = InStr(iPosINI, sALL, sEnd)
        GetKey = Mid(sALL, iPosINI, iPosEND - iPosINI)
        GetKey = Replace(GetKey, DQ, "")
    Else
        GetKey = sALL
    End If
End Function

Public Sub QuickSort(ByRef strArray() As String, ByRef intBottom As Integer, ByRef intTop As Integer)
    Dim strPivot As String, strTemp As String

    Dim intBottomTemp As Integer, intTopTemp As Integer
    intBottomTemp = intBottom

    intTopTemp = intTop
    strPivot = GetKey(strArray((intBottom + intTop) \ 2), "<G K=", ">")
    While (intBottomTemp <= intTopTemp)
        While (GetKey(strArray(intBottomTemp), "<G K=", ">") < strPivot And intBottomTemp < intTop)
            intBottomTemp = intBottomTemp + 1
            EnableEvents
        Wend

        While (strPivot < GetKey(strArray(intTopTemp), "<G K=", ">") And intTopTemp > intBottom)
            intTopTemp = intTopTemp - 1
            EnableEvents
        Wend

        If intBottomTemp < intTopTemp Then
            strTemp = strArray(intBottomTemp)
            strArray(intBottomTemp) = strArray(intTopTemp)
            strArray(intTopTemp) = strTemp
        End If

        If intBottomTemp <= intTopTemp Then
            intBottomTemp = intBottomTemp + 1
            intTopTemp = intTopTemp - 1
        End If
        EnableEvents
    Wend

    'faz a chamada recursiva a si propria ate que lista esteja preenchida
    If (intBottom < intTopTemp) Then QuickSort strArray, intBottom, intTopTemp

    If (intBottomTemp < intTop) Then QuickSort strArray, intBottomTemp, intTop
End Sub

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
''' Return byte array with VBA "Unicode" string encoded in UTF-8
Public Function Utf8BytesFromString(strInput As String) As Byte()
    Dim nBytes As Long
    Dim abBuffer() As Byte
    ' Get length in bytes *including* terminating null
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, vbNull, 0&, 0&, 0&)
    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
    Utf8BytesFromString = abBuffer
End Function

Public Function MatrixByteToString(ByRef buf) As String
    Dim i As Long, S As String
    S = ""
    For i = LBound(buf) To UBound(buf)
        S = S & Chr(buf(i))
    Next i
    MatrixByteToString = S
End Function

Public Function isUTF8(ByVal astr) As Boolean
    Dim c0 As String, c1 As String, c2 As String, c3 As String
    Dim n As Long

    isUTF8 = True
    n = 1
    Do While n <= Len(astr)
        c0 = Asc(Mid(astr, n, 1))
        If n <= Len(astr) - 1 Then
            c1 = Asc(Mid(astr, n + 1, 1))
        Else
            c1 = 0
        End If
        If n <= Len(astr) - 2 Then
            c2 = Asc(Mid(astr, n + 2, 1))
        Else
            c2 = 0
        End If
        If n <= Len(astr) - 3 Then
            c3 = Asc(Mid(astr, n + 3, 1))
        Else
            c3 = 0
        End If

        If (c0 And 240) = 240 Then
            If (c1 And 128) = 128 And (c2 And 128) = 128 And (c3 And 128) = 128 Then
                n = n + 4
            Else
                isUTF8 = False
                Exit Function
            End If
        ElseIf (c0 And 224) = 224 Then
            If (c1 And 128) = 128 And (c2 And 128) = 128 Then
                n = n + 3
            Else
                isUTF8 = False
                Exit Function
            End If
        ElseIf (c0 And 192) = 192 Then
            If (c1 And 128) = 128 Then
                n = n + 2
            Else
                isUTF8 = False
                Exit Function
            End If
        ElseIf (c0 And 128) = 0 Then
            n = n + 1
        Else
            isUTF8 = False
            Exit Function
        End If
    Loop
End Function

Public Function Decode_UTF8(ByVal astr) As String
    Dim c0 As String, c1 As String, c2 As String, c3 As String
    Dim n As Long
    Dim unitext As String

    If Not isUTF8(astr) Then
        Decode_UTF8 = astr
        Exit Function
    End If

    unitext = ""
    n = 1
    Do While n <= Len(astr)
        c0 = Asc(Mid(astr, n, 1))
        If n <= Len(astr) - 1 Then
            c1 = Asc(Mid(astr, n + 1, 1))
        Else
            c1 = 0
        End If
        If n <= Len(astr) - 2 Then
            c2 = Asc(Mid(astr, n + 2, 1))
        Else
            c2 = 0
        End If
        If n <= Len(astr) - 3 Then
            c3 = Asc(Mid(astr, n + 3, 1))
        Else
            c3 = 0
        End If

        If (c0 And 240) = 240 And (c1 And 128) = 128 And (c2 And 128) = 128 And (c3 And 128) = 128 Then
            unitext = unitext + ChrW((c0 - 240) * 65536 + (c1 - 128) * 4096) + (c2 - 128) * 64 + (c3 - 128)
            n = n + 4
        ElseIf (c0 And 224) = 224 And (c1 And 128) = 128 And (c2 And 128) = 128 Then
            unitext = unitext + ChrW((c0 - 224) * 4096 + (c1 - 128) * 64 + (c2 - 128))
            n = n + 3
        ElseIf (c0 And 192) = 192 And (c1 And 128) = 128 Then
            unitext = unitext + ChrW((c0 - 192) * 64 + (c1 - 128))
            n = n + 2
        ElseIf (c0 And 128) = 128 Then
            unitext = unitext + ChrW(c0 And 127)
            n = n + 1
        Else    ' c0 < 128
            unitext = unitext + ChrW(c0)
            n = n + 1
        End If
    Loop

    Decode_UTF8 = unitext
End Function

Public Function Encode_UTF8(ByVal astr) As String
    Dim c As String
    Dim n As Long
    Dim utftext As String

    If isUTF8(LCase(astr)) Then
        Encode_UTF8 = astr
        Exit Function
    End If

    utftext = ""
    n = 1
    Do While n <= Len(astr)
        c = AscW(Mid(astr, n, 1))
        If c < 128 Then
            utftext = utftext + Chr(c)
        ElseIf ((c >= 128) And (c < 2048)) Then
            utftext = utftext + Chr(((c \ 64) Or 192))
            utftext = utftext + Chr(((c And 63) Or 128))
        ElseIf ((c >= 2048) And (c < 65536)) Then
            utftext = utftext + Chr(((c \ 4096) Or 224))
            utftext = utftext + Chr((((c \ 64) And 63) Or 128))
            utftext = utftext + Chr(((c And 63) Or 128))
        Else    ' c >= 65536
            utftext = utftext + Chr(((c \ 262144) Or 240))
            utftext = utftext + Chr(((((c \ 4096) And 63)) Or 128))
            utftext = utftext + Chr((((c \ 64) And 63) Or 128))
            utftext = utftext + Chr(((c And 63) Or 128))
        End If
        n = n + 1    '� � �
    Loop
    Encode_UTF8 = utftext
End Function

'iconv+urlencode /based on ADODB.Stream (include a reference to a recent version of the "Microsoft ActiveX Data Objects" library in your project)
'http://stackoverflow.com/questions/218181/how-can-i-url-encode-a-string-in-excel-vba
'https://msdn.microsoft.com/en-us/library/ms681424%28v=vs.85%29.aspx
Public Function URLEncode(StringVal As Variant, Optional SpaceAsPlus As Boolean = False) As String
    Dim bytes() As Byte, B As Byte, i As Integer, space As String
    If SpaceAsPlus Then space = "+" Else space = "%20"

    If Len(StringVal) > 0 Then
        With New ADODB.Stream
            .Mode = adModeReadWrite
            .Type = adTypeText
            .Charset = "UTF-8"
            .Open
            .WriteText StringVal
            .Position = 0
            .Type = adTypeBinary
            .Position = 3    ' skip BOM
            bytes = .Read
        End With

        ReDim Result(UBound(bytes)) As String
        For i = UBound(bytes) To 0 Step -1
            B = bytes(i)
            Select Case B
            Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
                Result(i) = Chr(B)
            Case 32
                Result(i) = space
            Case 0 To 15
                Result(i) = "%0" & Hex(B)
            Case Else
                Result(i) = "%" & Hex(B)
            End Select
        Next i

        URLEncode = Join(Result, "")
    End If
End Function

Public Function SaveToFileUTF(ByRef TXT() As String, ByVal File, eSR As StreamWriteEnum, eSO As SaveOptionsEnum) As Boolean
    Dim objStream As ADODB.Stream, iLoop As Long
    Set objStream = New ADODB.Stream
    objStream.Charset = "utf-8"
    objStream.Open
    For iLoop = LBound(TXT) To UBound(TXT)
        objStream.WriteText TXT(iLoop), eSR ' adWriteLine
    Next iLoop
    objStream.SaveToFile File, eSO 'adSaveCreateOverWrite
    objStream.Close
    SaveToFileUTF = True
End Function

Public Function ReadFileinUTF(ByVal File) As String
    Dim objStream As ADODB.Stream, iLoop As Long
    On Error GoTo ErrEnd
        Set objStream = New ADODB.Stream
        objStream.Charset = "utf-8"
        objStream.Open
        objStream.Type = adTypeText
        objStream.LoadFromFile File
        ReadFileinUTF = objStream.ReadText(StreamReadEnum.adReadAll)
        objStream.Close
    On Error GoTo 0
    Exit Function
ErrEnd:
    LogSheet "ReadFileinUTF", File, Err.Description, "", A_Write
    ReadFileinUTF = ""
End Function

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Public Function FindContent(ByRef mat() As String, What As String) As Long
    Dim i1 As Long
    FindContent = -1
    For i1 = LBound(mat) To UBound(mat)
        If What = mat(i1) Then
            FindContent = i1
            Exit Function
        End If
        EnableEvents
    Next i1
End Function

Public Function FindRowOfText(SheetName As String, nCol As String, Text As String) As Long
    Dim R As Range, LastRow As Long
    Dim a As String, B As String
    Dim x As Long, y As Long
    With Workbooks(Main_Workbook_Name).Sheets(SheetName)
        Application.ScreenUpdating = False
        LastRow = .Cells(.Rows.Count, nCol).End(xlUp).row
        On Error Resume Next
        Set R = .Range(nCol & "1:" & nCol & LastRow).Find(Trim(Text), MatchCase:=False, LookIn:=xlValues)
        If Err.Number <> 0 Then
            Stop 'FindRowOfText = -2
        ElseIf R Is Nothing Then
            FindRowOfText = -1
        Else
            SplitFixedAddress R.Address, a, x
            FindRowOfText = x
        End If
        On Error GoTo 0
        Application.ScreenUpdating = True
    End With
End Function

Public Function FindCorrelatedText(SheetName As String, nCol As String, Text As String, ColRet As String) As String
    Dim iRow As Long, t1, t2
    iRow = 1
    t1 = Trim(Text)
    With Workbooks(Main_Workbook_Name).Sheets(SheetName)
        While .Range("A" & iRow) <> ""
            t2 = Trim(.Range(nCol & iRow))
            If t2 = t1 Then
                FindCorrelatedText = .Range(ColRet & iRow)
                Exit Function
            End If
            iRow = iRow + 1
            EnableEvents
        Wend
    End With

    'N�o achou
    iRow = 1
    t1 = GoodForXML(Text)
    t1 = Replace(t1, space(1), "")
    t1 = UCase(t1)
    With Workbooks(Main_Workbook_Name).Sheets(SheetName)
        While .Range("A" & iRow) <> ""
            t2 = GoodForXML(.Range(nCol & iRow))
            t2 = Replace(t2, space(1), "")
            t2 = UCase(t2)
            If t2 = t1 Then
                FindCorrelatedText = .Range(ColRet & iRow)
                Exit Function
            End If
            iRow = iRow + 1
            EnableEvents
        Wend
    End With
End Function

Public Function WhereIsColumn(SheetName As String, StartRow As Long, ByVal WhatToFind As String) As String
    Dim i1 As Long, Column As String
    Dim Idx As String, Resp As Variant

    WhatToFind = Trim(UCase(WhatToFind))
    WhereIsColumn = NotFoundYet
    Idx = Trim(UCase(SheetName)) & "." & StartRow & "." & WhatToFind

    On Error Resume Next
    If ColColumn.Count > 0 Then
        Resp = ColColumn.Item(Idx)
        If Err.Number = 0 Then
            WhereIsColumn = Resp
            Exit Function
        End If
    End If
    On Error GoTo 0

    For i1 = 1 To Workbooks(Main_Workbook_Name).Sheets(SheetName).Range("A" & StartRow).CurrentRegion.Columns.Count
        Column = IdToColumn(i1)
        If Trim(UCase$(Workbooks(Main_Workbook_Name).Sheets(SheetName).Range(Column & StartRow).Value)) = WhatToFind Then
            WhereIsColumn = Column
            Exit For
        End If
        EnableEvents
    Next i1
    ColColumn.add WhereIsColumn, Idx
End Function

Public Sub UpdateCollection(ByRef aColl As Collection, ByVal aKey As Variant, ByVal aValue As Variant, ByVal aPosition As Long)
    Dim temp As Variant, iCnt As Long
    temp = aColl.Item(aKey)
    If aValue = 0 Then
        'Stop 'TRAP
        LogSheet "aValue = 0", "Key:" & aKey & " , Val:" & aValue & " , Pos:" & aPosition, "Antes:" & temp(aPosition), "", A_Write
    End If
    temp(aPosition) = aValue
    iCnt = aColl.Count
    aColl.Remove aKey
    If iCnt - 1 <> aColl.Count Then
        Stop
    End If
    aColl.add temp, aKey
    If iCnt <> aColl.Count Then
        Stop
    End If
End Sub

Public Function ExponentialReplace(ByVal FullExpression As String) As String
    Dim i1 As Long, i2 As Long, i3 As Long, base As Long, expoente As Long, sDummyIni As String, sDummyEnd As String
    FullExpression = UCase(Trim(FullExpression))
    FullExpression = Replace(FullExpression, space(2), space(1))
    i1 = InStr(FullExpression, "^")
    If i1 = 0 Then
        ExponentialReplace = FullExpression
        Exit Function
    End If
    '======
    'possibilidades:
    i1 = InStr(FullExpression, "^")
    i2 = InStr(FullExpression, "(")
    i3 = InStr(FullExpression, ")")
    If i2 > 0 Then    '(2^0)
        sDummyIni = Left(FullExpression, i2 - 1)
        sDummyEnd = Mid(FullExpression, i3 + 1)
        base = Mid(FullExpression, i2 + 1, i1 - i2 - 1)
        expoente = Mid(FullExpression, i1 + 1, i3 - i1 - 1)
    Else    ' 2^0
        Stop    'ainda nao desenvolvido
    End If

    ExponentialReplace = sDummyIni & CLng(base ^ expoente) & sDummyEnd
    'recursivity
    If (ExponentialReplace <> FullExpression) Then
        ExponentialReplace = ExponentialReplace(ExponentialReplace)
    End If
End Function

Public Sub SheetUpdates(Actions As UpdateActions)
    If "" = Main_Workbook_Name Then SetMainWorkbookName
    Dim i1 As Long, bEnable As Boolean
    '-------------------------
    LastSheetUpdates = Actions
    '-------------------------
    Select Case Actions
        Case UpdateActions.EnableFull, UpdateActions.ReanimateScreen
            bEnable = True
        Case UpdateActions.NoneWork, UpdateActions.FreezeScreen
            bEnable = False
    End Select

    Select Case Actions
        Case UpdateActions.EnableFull, UpdateActions.NoneWork
            '<Enable Updates>
            Application.EnableEvents = bEnable
            Application.Calculation = IIf(bEnable, xlCalculationAutomatic, xlCalculationManual)
            Workbooks(Main_Workbook_Name).UpdateLinks = IIf(bEnable, xlUpdateLinksAlways, xlUpdateLinksNever)
            For i1 = 1 To Workbooks(Main_Workbook_Name).Sheets.Count
                Workbooks(Main_Workbook_Name).Sheets(i1).EnableCalculation = bEnable
                If bEnable Then Workbooks(Main_Workbook_Name).Sheets(i1).Activate
                DoEvents
            Next i1
            Application.ScreenUpdating = bEnable

        Case UpdateActions.FreezeScreen, UpdateActions.ReanimateScreen
            Application.ScreenUpdating = bEnable
    End Select
    DoEvents
    '</Enable Updates>
End Sub

Function HeaderSQLTable(ByVal H As String) As String()
    H = Replace(H, ",", ";")
    H = Replace(H, ";;", ";")
    H = Replace(H, space(1), "")
    If Right(H, 1) = ";" Then H = Left(H, -1 + Len(H))
    HeaderSQLTable = Split(H, ";")
End Function
Function HeaderCimplicityPoints() As String()
    Dim H As String
    H = ""
    H = H & "PT_ID;ACCESS;ACCESS_FILTER;ACK_TIMEOUT;ADDR;ADDR_OFFSET;ADDR_TYPE;ALM_CLASS;ALM_CRITERIA;ALM_DEADBAND;ALM_DELAY;ALM_ENABLE;ALM_HIGH_1;ALM_HIGH_2;ALM_HLP_FILE;ALM_LOW_1;ALM_LOW_2;ALM_MSG;ALM_ROUTE_OPER;ALM_ROUTE_SYSMGR;ALM_ROUTE_USER;ALM_SEVERITY;ALM_STR;ALM_TYPE;ALM_UPDATE_VALUE;ANALOG_DEADBAND;BFR_COUNT;BFR_DUR;BFR_EVENT_PERIOD;BFR_EVENT_PT_ID;BFR_EVENT_TYPE;"
    H = H & "BFR_EVENT_UNITS;BFR_GATE_COND;BFR_SYNC_TIME;CALC_TYPE;CONV_LIM_HIGH;CONV_LIM_LOW;CONV_TYPE;DELAY_LOAD;DELETE_REQ;DESC;DEVIATION_PT;DEVICE_ID;DISP_LIM_HIGH;DISP_LIM_LOW;DISP_TYPE;DISP_WIDTH;ELEMENTS;ENG_UNITS;ENUM_ID;EQUATION;EXTRA;FW_CONV_EQ;GR_SCREEN;INIT_VAL;JUSTIFICATION;LEVEL;LOCAL;LOG_ACK;LOG_DATA;LOG_DEL;LOG_GEN;LOG_RESET;MAX_STACKED;MEASUREMENT_UNIT_ID;"
    H = H & "MISC_FLAGS;POLL_AFTER_SET;PRECISION;PROC_ID;PTMGMT_PROC_ID;PT_ENABLED;PT_ORIGIN;PT_SET_INTERVAL;PT_SET_TIME;PT_TYPE;RANGE_HIGH;RANGE_LOW;RAW_LIM_HIGH;RAW_LIM_LOW;REP_TIMEOUT;RESET_ALLOWED;RESET_COND;RESET_PT;RESET_TIMEOUT;RESOURCE_ID;REV_CONV_EQ;ROLLOVER_VAL;SAFETY_PT;SAMPLE_INTV;SAMPLE_INTV_UNIT;SCAN_RATE;SETPOINT_HIGH;SETPOINT_LOW;TIME_OF_DAY;TRIG_CK_PT;TRIG_PT;"
    H = H & "TRIG_REL;TRIG_VAL;UAFSET;UPDATE_CRITERIA;VARIANCE_VAL;VARS;"

    HeaderCimplicityPoints = HeaderSQLTable(H)
End Function

Function HeaderCimplicityOPC() As String()
    Dim H As String
    H = ""
    H = H & "OPCKey;DCOMTimeoutThreshold;DelayBeforeRefresh;DeviceReadAfterSet;IsPolled;StartupDelay;"
    H = H & "ItemAccessPathEnable;NoAccessPath;PingBeforePoll;PingBeforeWrite;PingInterval;"
    H = H & "PingTimeout;ReAddAsEmptyOnBadType;ReconnectInterval;RemoveItemsOnRemoveGroup;RestartDelay;ScanRate;UseLocalReg;"
    H = H & "EightByteReals;UseServerTimeStamp;UseDataTypePromotion;RefreshUnsoItems;AddBoolAsBool;"

    HeaderCimplicityOPC = HeaderSQLTable(H)
End Function

Public Sub ClearGlobalVariables()
    Main_Workbook_Name = ""
    Main_Workbook_Path = ""
End Sub
Public Sub InitVBs()
    Dim nr As Long
    If "" = Main_Workbook_Name Then SetMainWorkbookName
    For nr = 1 To Workbooks(Main_Workbook_Name).Sheets.Count
        Workbooks(Main_Workbook_Name).Sheets(nr).AutoFilterMode = False
    Next nr
    EnableEvents
End Sub

Public Sub ReCalc()
    If "" = Main_Workbook_Name Then SetMainWorkbookName
    'Keystroke: SHIFT+F9
    Dim nr As Long
    For nr = 1 To Workbooks(Main_Workbook_Name).Sheets.Count
        With Workbooks(Main_Workbook_Name).Sheets(nr)
            .EnableCalculation = False
            DoEvents
            .EnableCalculation = True
            .Calculate
        End With
    Next nr
End Sub

'</
Private Sub FillColor()
    mColorAnima(0, 0) = "SF": mColorAnima(0, 1) = "PURPLE": mColorAnima(0, 2) = 5
    mColorAnima(1, 0) = "CF": mColorAnima(1, 1) = "PURPLE": mColorAnima(1, 2) = 8
    mColorAnima(2, 0) = "MF": mColorAnima(2, 1) = "RED": mColorAnima(2, 2) = 10
    mColorAnima(3, 0) = "LF": mColorAnima(3, 1) = "RED": mColorAnima(3, 2) = 12
    mColorAnima(4, 0) = "QF": mColorAnima(4, 1) = "YELLOW": mColorAnima(4, 2) = 15
    mColorAnima(5, 0) = "PF": mColorAnima(5, 1) = "YELLOW": mColorAnima(5, 2) = 15
    mColorAnima(6, 0) = "TFIB": mColorAnima(6, 1) = "BLUE": mColorAnima(6, 2) = 25
    mColorAnima(7, 0) = "TFS": mColorAnima(7, 1) = "WHITE": mColorAnima(7, 2) = 30
    mColorAnima(8, 0) = "MA": mColorAnima(8, 1) = "RED": mColorAnima(8, 2) = 32
    mColorAnima(9, 0) = "AF": mColorAnima(9, 1) = "DARK": mColorAnima(9, 2) = 32
    mColorAnima(10, 0) = "MW": mColorAnima(10, 1) = "ORANGE": mColorAnima(10, 2) = 35
    mColorAnima(11, 0) = "FW": mColorAnima(11, 1) = "TAN": mColorAnima(11, 2) = 40
    mColorAnima(12, 0) = "PW": mColorAnima(12, 1) = "MAROON": mColorAnima(12, 2) = 60
    mColorAnima(13, 0) = "TW": mColorAnima(13, 1) = "AQUA": mColorAnima(13, 2) = 60
    mColorAnima(14, 0) = "SW": mColorAnima(14, 1) = "FUCHSIA": mColorAnima(14, 2) = 60
    mColorAnima(15, 0) = "LW": mColorAnima(15, 1) = "ROSE": mColorAnima(15, 2) = 60
    mColorAnima(16, 0) = "QW": mColorAnima(16, 1) = "OLIVE": mColorAnima(16, 2) = 60
    mColorAnima(17, 0) = "AW": mColorAnima(17, 1) = "TEAL": mColorAnima(17, 2) = 60
    mColorAnima(18, 0) = "SS": mColorAnima(18, 1) = "GRAY": mColorAnima(18, 2) = 65
    mColorAnima(19, 0) = "MS": mColorAnima(19, 1) = "GRAY": mColorAnima(19, 2) = 65
    mColorAnima(20, 0) = "PS": mColorAnima(20, 1) = "GRAY": mColorAnima(20, 2) = 65
    mColorAnima(21, 0) = "QS": mColorAnima(21, 1) = "GRAY": mColorAnima(21, 2) = 65
    mColorAnima(22, 0) = "LS": mColorAnima(22, 1) = "GRAY": mColorAnima(22, 2) = 65
    mColorAnima(23, 0) = "AS": mColorAnima(23, 1) = "GRAY": mColorAnima(23, 2) = 65
    mColorAnima(24, 0) = "TS": mColorAnima(24, 1) = "GRAY": mColorAnima(24, 2) = 65
End Sub
Private Sub CalcAnimaColorOrder(ByVal FullPoint As String, ByRef Color3 As String, ByRef BaseOrder As Long)
    If mColorAnima(0, 0) = "" Then FillColor
    Dim ptType As String, Point As String
    Dim ID As Long, iPos As Long
    iPos = InStrRev(FullPoint, "\", Len(FullPoint))
    Point = Mid(FullPoint, iPos + 1)
    ptType = Mid(Point, 2, 2)

    For ID = LBound(mColorAnima, 1) To UBound(mColorAnima, 1)
        If ptType = mColorAnima(ID, 0) Then
            Color3 = mColorAnima(ID, 1)
            BaseOrder = mColorAnima(ID, 2)
            Exit For
        End If
    Next
    If Color3 = "" Then Beep    'Stop
    If BaseOrder = 0 Then Beep    'Stop
End Sub
Public Function AnimaColor(ByVal FullPoint As String) As String
    Dim dummy As Long
    If FullPoint = 0 Then
        AnimaColor = "GREEN"
        Exit Function
    End If
    CalcAnimaColorOrder FullPoint, AnimaColor, dummy
End Function
Public Function AnimaColorOrder(ByVal FullPoint As String) As Long
    Dim dummy As String
    If FullPoint = 0 Then
        AnimaColorOrder = 99
        Exit Function
    End If
    CalcAnimaColorOrder FullPoint, dummy, AnimaColorOrder
End Function
Public Function ColorByAlarmClass(ByVal AlarmClass As String) As String
    If mColorAnima(0, 0) = "" Then FillColor
    AlarmClass = removeVBCRLF(AlarmClass)
    Dim i1 As Long
    For i1 = LBound(mColorAnima, 1) To UBound(mColorAnima, 1)
        If AlarmClass = mColorAnima(i1, 0) Then
            ColorByAlarmClass = mColorAnima(i1, 1)
            Exit Function
        End If
        EnableEvents
    Next i1
    Stop
End Function
'>/

Public Function LimpaBeiradas(ByVal sDado As String, ByVal Filtro As String) As String
    Dim iPos As Long, iLast As Long
    While Len(sDado) <> iLast
        iLast = Len(sDado)
        If Left(sDado, 1) = Filtro Then sDado = Mid(sDado, 2)
        If Right(sDado, 1) = Filtro Then sDado = Left(sDado, Len(sDado) - 1)
        EnableEvents
    Wend
    LimpaBeiradas = sDado
End Function

Public Function FaixaBits(Primeiro, Ultimo) As Long
    Dim i1 As Long, i2 As Long
    For i1 = 0 To 15
        i2 = 2 ^ i1
        If i2 >= Primeiro And i2 <= Ultimo Then
            FaixaBits = FaixaBits + i2
        End If
    Next i1
End Function

Public Function MatrixReady(ByRef m) As Boolean
    MatrixReady = False
    If Not IsArray(m) Then Exit Function
    '
    Dim i As Long
    On Error Resume Next
    i = UBound(m)
    If Err.Number = 0 Then MatrixReady = True
    On Error GoTo 0
End Function

Function IncFormula(ByVal orig As String) As String
    Dim iPos(0 To 2) As Long, iLetter As Long, iCnt As Long
    Dim sF As String
    If Left(orig, 1) <> "=" Then
        IncFormula = orig
        Exit Function
    End If

    '=A259+1
    '=tblEquipment!A260
    iPos(0) = InStr(orig, "!")
    If iPos(0) = 0 Then iPos(0) = 1
    For iLetter = iPos(0) To Len(orig)
        If IsNumeric(Mid(orig, iLetter, 1)) Then
            If iPos(1) = 0 Then
                iPos(1) = iLetter
            Else
                iPos(2) = iLetter
            End If
        Else
            If iPos(1) <> 0 Then Exit For
        End If
        EnableEvents
    Next iLetter

    iCnt = Mid(orig, iPos(1), iPos(2) - iPos(1) + 1)
    iCnt = iCnt + 1
    sF = Left(orig, iPos(1) - 1)
    sF = sF & iCnt
    sF = sF & Mid(orig, iPos(2) + 1)
    IncFormula = sF
End Function

Public Function Unique(iRow As Long, iCol As Long, ColumnFull As Range) As String
    Exit Function
    InitVBs
    If iRow = 1 Then
        Unique = "Unique"
        Exit Function
    End If

    Dim sCol As String, rME As Range, r1 As String, sSH As String
    sCol = IdToColumn(iCol)
    r1 = sCol & "2:" & sCol & iRow - 1
    sSH = ActiveSheet.Name
    Set rME = Workbooks(Main_Workbook_Name).Sheets(sSH).Range(r1)

    '==========================
    Exit Function
    '==========================

    Dim c As Range, lLoop As Long, lLastFound
    lLastFound = 2
    For lLoop = 1 To iRow
        Set c = ColumnFull.Cells(lLastFound, 1).Find(rME.Cells(lLoop, 1), LookIn:=xlValues, MatchCase:=True, LookAt:=XlLookAt.xlWhole)
        If Not c Is Nothing Then
            Unique = ColumnFull.Cells(iRow)
        Else
            Unique = ColumnFull.Cells(iRow, 1)
        End If
        EnableEvents
    Next lLoop
End Function

Public Function ImportIDT(ByVal sFile As String, ByVal sGroup As String) As Boolean
    InitVBs
    Dim s1 As String, sSheet As String, sLink As String
    's1 = Main_Workbook_Path
    'Stop
    'If Right(s1, 1) <> "\" Then s1 = s1 & "\"
    's1 = s1 & sGroup & "\" & sFile
    'If InStr(s1, ".idt") = 0 Then s1 = s1 & ".idt"

    Dim obFile As New clsFile
    If InStr(sFile, ".idt") = 0 Then sFile = sFile & ".idt"
    obFile.Arq = sFile
    obFile.BasePath = Main_Workbook_Path
    obFile.SubPath = sGroup

    s1 = obFile.FullPath & "\" & obFile.FileAndExtension
    If Dir(s1) = "" Then
        PlayWAV "chord.wav"
        MsgBox "Nao Existe " & s1 & " de " & sGroup
        ImportIDT = False
        Exit Function
    End If
    sSheet = Replace(sFile, ".idt", "")
    sSheet = sSheet & "." & sGroup
    NewSheet UCase(sSheet)
    Dim iRow As Long, iCol As Long
    Dim DataLine As String, FileNum As Long, Colunas() As String
    Dim vData As Variant, FullLine As String, cData As Collection
    Dim PassoDeLeitura As Long
    PassoDeLeitura = 1
    '
    ReDim Colunas(0 To 0)
    iRow = 1
    FileNum = FreeFile()
    '
    Open s1 For Input As #FileNum
    While Not EOF(FileNum)
        Line Input #FileNum, DataLine
        DataLine = Trim(DataLine)
        DataLine = removeVBCRLF(DataLine)
        Select Case PassoDeLeitura
        Case 1
            'Achar qtd parametros
            If Left(DataLine, 1) = "|" Then
                'nothing to do
            ElseIf Left(DataLine, 1) = "*" And Len(DataLine) > 1 Then
                While InStr(DataLine, space(1))
                    DataLine = Replace(DataLine, space(1), "|")
                    EnableEvents
                Wend
                While InStr(DataLine, "||")
                    DataLine = Replace(DataLine, "||", "|")
                    EnableEvents
                Wend
                vData = Split(DataLine, "|")
                If UBound(vData) >= 2 Then
                    If IsNumeric(vData(1)) Then
                        If vData(1) > UBound(Colunas) Then ReDim Preserve Colunas(vData(1))
                        Colunas(vData(1)) = vData(2)
                    End If
                End If
            Else
                If UBound(Colunas) > 0 Then
                    If iRow = 1 Then
                        With Workbooks(Main_Workbook_Name).Sheets(sSheet)
                            For iCol = LBound(Colunas) To UBound(Colunas)
                                .Range(IdToColumn(iCol + 1) & iRow).Value = Colunas(iCol)
                                EnableEvents
                            Next iCol
                        End With
                        iRow = iRow + 1
                    End If
                    PassoDeLeitura = PassoDeLeitura + 1
                End If
            End If
        Case 2
            If Right(DataLine, 1) = "-" Then
                FullLine = FullLine & Left(DataLine, Len(DataLine) - 1)
                PassoDeLeitura = PassoDeLeitura + 1
            Else
                FullLine = FullLine & DataLine
                GoSub LoopEscreve
            End If
        Case 3
            If Right(DataLine, 1) = "-" Then
                FullLine = FullLine & Left(DataLine, Len(DataLine) - 1)
            Else
                FullLine = FullLine & DataLine
                GoSub LoopEscreve
            End If
        Case Else
            Stop
        End Select
        EnableEvents
    Wend
    Close #FileNum

    With Workbooks(Main_Workbook_Name).Sheets(sSheet)
        s1 = IdToColumn(LBound(Colunas) + 1) & ":" & IdToColumn(UBound(Colunas) + 1)
        .Columns(s1).AutoFit
        .Columns(s1).HorizontalAlignment = xlLeft
        .Tab.Color = RGBColorIndex.Teal
    End With

    ImportIDT = True
    Exit Function
    '-------------------------------------------------

LoopEscreve:
    While InStr(FullLine, "||")
        FullLine = Replace(FullLine, "||", "| |")
        EnableEvents
    Wend
    vData = Split(FullLine, "|")
    If UBound(vData) <> UBound(Colunas) Then Stop
    With Workbooks(Main_Workbook_Name).Sheets(sSheet)
        For iCol = LBound(vData) To UBound(vData)
            s1 = vData(iCol)
            s1 = Trim(s1)
            If IsNumeric(Left(vData(iCol), 1)) Then s1 = SQ & s1
            .Range(IdToColumn(iCol + 1) & iRow).Value = s1
            EnableEvents
        Next iCol
    End With
    FullLine = ""
    iRow = iRow + 1
    PassoDeLeitura = 2
    Return
End Function

Public Function NumberOfDimensions(ByRef Xarray() As String) As Long
    Dim DimNum As Long, ErrorCheck As Long
    'Sets up the error handler.
    On Error GoTo FinalDimension
    'Visual Basic for Applications arrays can have up to 60000
    'dimensions; this allows for that.
    For DimNum = 1 To 60000
        'It is necessary to do something with the LBound to force it
        'to generate an error.
        ErrorCheck = LBound(Xarray, DimNum)
        EnableEvents
    Next DimNum
    Exit Function
    ' The error routine.
FinalDimension:
    NumberOfDimensions = DimNum - 1
    On Error GoTo 0
End Function

Public Sub AddToMatrix(ByRef Xarray() As String, Value As String, Optional element As Long)
    Dim i1 As Long
    i1 = NumberOfDimensions(Xarray)
    If i1 <> 1 Then Stop
    If Xarray(UBound(Xarray, i1)) <> "" Then
        ReDim Preserve Xarray(1 + UBound(Xarray))
    End If
    Xarray(UBound(Xarray, i1)) = Value
End Sub

Function ImportCSVXML(FileName As String) As Boolean
    Dim sPath As String
    ImportCSVXML = True
    sPath = ThisWorkbook.Path & "\" & FileName 'semicolonseparated.csv '.xml
    If Right(LCase(FileName), 4) = ".csv" Then
        Workbooks.OpenText FileName:=sPath, DataType:=xlDelimited, Semicolon:=True ', Local:=True
    ElseIf Right(LCase(FileName), 4) = ".xml" Then
        Workbooks.OpenXML FileName:=sPath, LoadOption:=XlXmlLoadOption.xlXmlLoadOpenXml
    Else
        ImportCSVXML = False
    End If
End Function

Function ImportExcelCSV(ByVal sWorkbookFrom As String, ByVal sSheetFrom As String, Optional iTabColor As Long = -1) As Boolean
'Function ImportExcelCSV(ByVal sWorkbookFrom As String, ByVal sSheetFrom As String, ByVal FirstLine As Long, Optional sNewName As String, Optional iTabColor As Long = -1) As Long
    'ImportCSV (for limit 1.000.000 lines)
    'return -1 - error
    '        0 - done and end
    '        n - done and has n lines to solve

    Const MaxLines = 1000000
    Dim UpdateScreenZ1 As UpdateActions
    Dim nr As Long, iFind As Long, iLine As Long
    Dim Found As Boolean, mArgs() As String, iArgs As Long
    Dim wb As Workbook, ws1 As Excel.Worksheet
    '-----------------------------------------------
    UpdateScreenZ1 = LastSheetUpdates
    If "" = Main_Workbook_Name Then SetMainWorkbookName
    If Not (WorkbookExist(sWorkbookFrom)) Then
        ImportExcelCSV = False
        Exit Function
    End If
    '-----------------------------------------------
    SheetUpdates NoneWork

    Dim obFile As New clsFile
    obFile.Arq = sWorkbookFrom
    obFile.BasePath = Main_Workbook_Path
    'obFile.SubPath = sWorkbookFrom
    '
    Dim sFileFrom As String
    sFileFrom = obFile.FileAndExtension
    While InStr(sFileFrom, "\")
        EnableEvents
        sFileFrom = Mid(sFileFrom, 1 + InStr(sFileFrom, "\"))
    Wend

    '1 verify how many lines to process
    Dim mFind() As String
    KepWare.UnifiedData sSheetFrom, mFind(), True

    Dim FileNum As Long, iRow As Long, DataLine As String, cLine As Collection, mSplit As Variant, sValue As String
    FileNum = FreeFile() '<<==

    Open obFile.FullPath & "\" & obFile.FileAndExtension For Input As #FileNum
    iRow = 0
    Set cLine = New Collection
    While Not EOF(FileNum)
        iRow = iRow + 1
        Line Input #FileNum, DataLine
        If 1 = iRow Then
            cLine.add DataLine 'always ass Header
        ElseIf cLine.Count > MaxLines Then
            Set cLine = Nothing
            ImportExcelCSV = False
            Exit Function
        Else
            mSplit = Split(DataLine, ",")
            sValue = mSplit(1)
            sValue = Replace(sValue, Freitas.DQ, "")
            'If InStr(UCase(sValue), "TPUTSTA08.DINTDATA[99]") > 0 Then Stop
            KepWare.UnifiedDataRejected sValue
            For iFind = LBound(mFind) To UBound(mFind)
                If UCase(sValue) = "X" Then Exit For
                Found = InStr(UCase(sValue), UCase(mFind(iFind)))
                If Not Found Then
                    Found = True
                    If InStr(mFind(iFind), "%") Then
                        mArgs = Split(mFind(iFind), "%")
                        For iArgs = LBound(mArgs) To UBound(mArgs)
                            If Found Then
                                Found = InStr(UCase(sValue), UCase(mArgs(iArgs)))
                            Else
                                Exit For
                            End If
                        Next iArgs
                    Else
                        Found = False
                    End If
                End If
                If Found Then
                    cLine.add DataLine
                    Exit For
                End If
            Next iFind
        End If
        EnableEvents
    Wend
    Close #FileNum
    '------------------------------------------------
    Dim OutCSV As String
    FileNum = FreeFile() '<<==
    OutCSV = "tmp.csv"

    On Error Resume Next
        Kill obFile.FullPath & "\" & OutCSV
    On Error GoTo 0

    Open obFile.FullPath & "\" & OutCSV For Output As #FileNum
    For iLine = 1 To cLine.Count
        Print #FileNum, cLine.Item(iLine)
    Next iLine
    Close #FileNum
    '------------------------------------------------
    ' * Copy
    Dim orgName(0 To 1) As String
    Workbooks.Open FileName:=obFile.FullPath & "\" & OutCSV, ReadOnly:=True
    nr = 1
    orgName(0) = UCase(Workbooks(OutCSV).Sheets(nr).Name)
    orgName(1) = UCase(sSheetFrom)

    Set ws1 = Workbooks(OutCSV).Sheets(orgName(0))
    ws1.AutoFilterMode = False
    ws1.EnableCalculation = False
    ws1.Copy After:=Workbooks(Main_Workbook_Name).Sheets(Workbooks(Main_Workbook_Name).Sheets.Count)
    Workbooks(Main_Workbook_Name).Sheets(ws1.Name).Name = orgName(1)

    If Not IsMissing(iTabColor) Then
        If iTabColor <> -1 Then
            Workbooks(Main_Workbook_Name).Sheets(orgName(1)).Tab.ColorIndex = iTabColor
        End If
    End If

    Workbooks(OutCSV).Close SaveChanges:=False
    Set ws1 = Nothing
    Set cLine = Nothing
    SheetUpdates UpdateScreenZ1
    On Error Resume Next
        Kill obFile.FullPath & "\" & OutCSV
    On Error GoTo 0
    ImportExcelCSV = True
End Function

Function ImportExcel(ByVal sWorkbookFrom As String, ByVal sSheetFrom As String, Optional sNewName As String, Optional iTabColor As Long = -1) As Boolean
    Dim UpdateScreenZ1 As UpdateActions
    Dim Tilt As Boolean, orgName(0 To 1) As String, woName(0 To 1) As String
    Dim nr As Long, i As Long, Turno As Long
    Dim wb As Workbook, ws1 As Excel.Worksheet
    '-----------------------------------------------
    UpdateScreenZ1 = LastSheetUpdates
    If "" = Main_Workbook_Name Then SetMainWorkbookName
    If Not (WorkbookExist(sWorkbookFrom)) Then Exit Function
    '-----------------------------------------------
    SheetUpdates NoneWork

    Dim obFile As New clsFile
    obFile.Arq = sWorkbookFrom
    obFile.BasePath = Main_Workbook_Path
    'obFile.SubPath = sWorkbookFrom
    '
    Dim sFileFrom As String
    sFileFrom = obFile.FileAndExtension
    While InStr(sFileFrom, "\")
        EnableEvents
        sFileFrom = Mid(sFileFrom, 1 + InStr(sFileFrom, "\"))
    Wend
    ' * Copy
    Workbooks.Open FileName:=obFile.FullPath & "\" & obFile.FileAndExtension, ReadOnly:=True
    For Each wb In Workbooks
        EnableEvents
        If UCase(wb.Name) = UCase(sFileFrom) Then
            If sSheetFrom = "1" Then sSheetFrom = UCase(Workbooks(sFileFrom).Sheets(1).Name)
            For nr = 1 To Workbooks(sFileFrom).Sheets.Count
                EnableEvents
                '2018 03 19 - remove space in sheet name
                orgName(0) = UCase(Workbooks(sFileFrom).Sheets(nr).Name)
                woName(0) = Replace(orgName(0), space(1), "")
                '
                orgName(1) = UCase(sSheetFrom)
                woName(1) = Replace(orgName(1), space(1), "")
                'If UCase(sSheetFrom) = UCase(Workbooks(sFileFrom).Sheets(nr).Name) Then
                If (orgName(0) = orgName(1)) Or (woName(0) = woName(1)) Then
                    '<2018 0302>
                    BreakLinks Workbooks(sFileFrom)
                    Workbooks(sFileFrom).UpdateRemoteReferences = False
                    Workbooks(sFileFrom).SaveLinkValues = True
                    'Workbooks(sFileFrom).Final = True
                    '</2018 0302>
                    Set ws1 = Workbooks(sFileFrom).Sheets(orgName(0)) 'sSheetFrom)
                    ws1.AutoFilterMode = False
                    ws1.EnableCalculation = False

                    EnableEvents
                    ws1.Copy After:=Workbooks(Main_Workbook_Name).Sheets(Workbooks(Main_Workbook_Name).Sheets.Count)
                    EnableEvents

                    With Workbooks(Main_Workbook_Name).Sheets(ws1.Name)
                        .Activate
                        ActiveWindow.FreezePanes = False
                        '
                        .Range("A2").Select
                        .AutoFilterMode = False
                        MaxAutoWidth ws1.Name, 40
                        If InStr(sWorkbookFrom, ".csv") = 0 Then
                            unMergeSheet orgName(0) 'sSheetFrom
                        End If
                        '
                        '<Convert All Cells To Text>
                        With ActiveSheet.UsedRange
                            .ClearFormats
                            .ClearComments
                            '.Value = .Value
                            .Cells.Copy
                            .Cells.PasteSpecial XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone
                            '.NumberFormat = "@"
                            .Replace What:="#N/A", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
                            .Validation.Delete
                             .Cells(1).Select
                        End With
                        Application.CutCopyMode = False
                        '</Convert All Cells To Text>
                        '
                    End With

                    If Not IsMissing(sNewName) Then
                        If sNewName <> "" Then
                            Workbooks(Main_Workbook_Name).Sheets(ws1.Name).Name = sNewName
                        Else
                            sNewName = orgName(0) 'sSheetFrom
                        End If
                    End If

                    If Not IsMissing(iTabColor) Then
                        If iTabColor <> -1 Then
                            Workbooks(Main_Workbook_Name).Sheets(sNewName).Tab.ColorIndex = iTabColor
                        End If
                    End If
                    Exit For
                End If
            Next nr
        End If
    Next

    Workbooks(sFileFrom).Close SaveChanges:=False
    Set ws1 = Nothing
    SheetUpdates UpdateScreenZ1
    ImportExcel = True
End Function

Public Sub MaxAutoWidth(sSheet As String, iMaxWidth As Long)
    If "" = Main_Workbook_Name Then SetMainWorkbookName
    Dim mCell As Range
    With Workbooks(Main_Workbook_Name).Sheets(sSheet)
        For Each mCell In .UsedRange.Rows(1).Cells
            mCell.EntireColumn.AutoFit
            If mCell.EntireColumn.ColumnWidth > iMaxWidth Then
                mCell.EntireColumn.ColumnWidth = iMaxWidth
            ElseIf mCell.EntireColumn.ColumnWidth < 8 Then
                mCell.EntireColumn.ColumnWidth = 8
            End If
            EnableEvents
        Next mCell
    End With
End Sub

Public Sub unMergeSheet(sSheet As String, Optional iLastRow As Long = -1)
    If "" = Main_Workbook_Name Then SetMainWorkbookName
    Dim mCell As Range, iRow As Long
    If Not IsMissing(iLastRow) Then
       If -1 = iLastRow Then
        'Using UsedRange
        With Workbooks(Main_Workbook_Name).Sheets(sSheet)
             .UsedRange 'Refresh UsedRange
             iLastRow = .UsedRange.Rows(.UsedRange.Rows.Count).row
        End With
       End If
    End If
    '
    With Workbooks(Main_Workbook_Name).Sheets(sSheet)
        For iRow = 1 To iLastRow
            For Each mCell In .UsedRange.Rows(iRow).Cells
                If mCell.MergeCells Then mCell.MergeArea.UnMerge
                EnableEvents
            Next mCell
            EnableEvents
        Next iRow
    End With
End Sub

Public Sub ShowFirstSheet()
    If "" = Main_Workbook_Name Then SetMainWorkbookName
    Workbooks(Main_Workbook_Name).Sheets(1).Activate
    Workbooks(Main_Workbook_Name).Sheets(1).Select
End Sub

Public Function KeepZeroLeftbyColumn(ByVal sSheet As String, ByVal sCollumn As String, ByVal iRowFirst As Long, ByVal iRowLast As Long)
    Dim i1 As Long, i2 As Long, sVal As String, bNeedFullUpdate As Boolean
    With Workbooks(Main_Workbook_Name).Sheets(sSheet)
        For i1 = iRowFirst To iRowLast
            bNeedFullUpdate = False
            sVal = .Range(sCollumn & i1).Value
            i2 = Len(sVal)
            sVal = Trim(removeVBCRLF(sVal))
            If i2 <> Len(sVal) Then bNeedFullUpdate = True
            '----------------------------------------------
            i2 = Len(sVal)
            sVal = KeepZeroLeft(sVal)
            sVal = Trim(sVal)
            If i2 <> Len(sVal) Then bNeedFullUpdate = True
            '----------------------------------------------
            If bNeedFullUpdate Then
                .Range(sCollumn & i1).Value = sVal
            End If
        Next i1
    End With
End Function

Public Function KeepZeroLeft(ByVal sValue As String) As String
    If IsNumeric(Left(sValue, 1)) Then sValue = SQ & sValue
    KeepZeroLeft = sValue
End Function
Public Function RevKeepZeroLeft(ByVal sValue As String) As String
    While SQ = (Left(sValue, 1))
        sValue = Mid(sValue, 2)
        EnableEvents
    Wend
    RevKeepZeroLeft = sValue
End Function
Public Function ExisteNaColecao(c As Collection, f As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = c.Item(f)
    ExisteNaColecao = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function GetlnkPath(ByVal Lnk As String) As String
'Getlnkpath("C:\Users\******\Desktop\Emily.lnk")
    On Error Resume Next
    With CreateObject("Wscript.Shell").CreateShortcut(Lnk)
        GetlnkPath = .TargetPath
        .Close
    End With
    On Error GoTo 0
End Function

Public Function GetTruePath(ByVal base As String, ByVal Subdir As String) As String
    'update 2017 12 subdir with sub sub sub
    Dim s1 As String, sLink As String, sSub() As String, iID As Long
    '
    If Right(base, 1) = "\" Then base = Left(base, Len(base) - 1)
    If Right(Subdir, 1) = "\" Then Subdir = Left(Subdir, Len(Subdir) - 1)
    If Left(Subdir, 1) = "\" Then Subdir = Mid(Subdir, 2)
    sSub = Split(Subdir, "\")
    '
    'seja caminho absoluto
    s1 = base
    For iID = LBound(sSub) To UBound(sSub)
        s1 = s1 & "\" & sSub(iID)
        EnableEvents
    Next iID
    If (Dir(s1 & "\") <> "") Then
        GetTruePath = s1
        Exit Function
    End If
    'seja caminho link
    s1 = base
    For iID = LBound(sSub) To UBound(sSub)
        s1 = s1 & "\" & sSub(iID)
        If (Dir(s1 & "\") <> "") And iID = UBound(sSub) Then
            GetTruePath = s1
            Exit Function
        ElseIf (Dir(s1 & "\") = "") And iID <> UBound(sSub) Then
            sLink = s1 & ".lnk"
            s1 = GetlnkPath(sLink)
            If Dir(s1 & "\") = "" Then
                GetTruePath = ""
                Exit Function
            End If
        ElseIf (Dir(s1 & "\") = "") And 0 = UBound(sSub) Then
            sLink = s1 & ".lnk"
            s1 = GetlnkPath(sLink)
            If Dir(s1 & "\") = "" Then
                GetTruePath = ""
            Else
                GetTruePath = s1
            End If
            Exit Function
        End If
        EnableEvents
    Next iID

    If UBound(sSub) = -1 Then 'add 2018 01 17
        GetTruePath = base
        Exit Function
    End If

    GetTruePath = ""
End Function

Public Function ImportOPCini(ByVal sFile As String, ByVal sSubPath As String, ByVal sOPCPath As String, ByVal sSheet As String) As Boolean
    InitVBs
    '
    Dim obFile As New clsFile
    Dim s1 As String, DataLine As String, Field As String, Value As String, sCol As String
    Dim FileNum As Long, iPos As Long, iCol As Long, iRow As Long
    '
    ImportOPCini = False
    obFile.Arq = sFile
    obFile.BasePath = Main_Workbook_Path
    obFile.SubPath = sSubPath & "\" & sOPCPath
    s1 = obFile.FullPath
    If s1 = "" Then Exit Function
    '============================
    s1 = obFile.FullPath & "\" & obFile.FileAndExtension
    If Dir(s1) = "" Then Exit Function
    '============================
    iCol = 1: iRow = 1
    With Workbooks(Main_Workbook_Name).Sheets(sSheet)
        While .Range(IdToColumn(iCol) & iRow).Value <> ""
            iRow = iRow + 1
            EnableEvents
        Wend
    End With

    '============================
    FileNum = FreeFile()
    Open s1 For Input As #FileNum
    While Not EOF(FileNum)
        Line Input #FileNum, DataLine
        DataLine = Trim(DataLine)
        DataLine = removeVBCRLF(DataLine)
        If Left(DataLine, 1) = "[" And Right(DataLine, 1) = "]" Then
            Field = "OPCKey"
            Value = DataLine
            sCol = WhereIsColumn(sSheet, 1, Field)
            With Workbooks(Main_Workbook_Name).Sheets(sSheet)
                If .Range(sCol & iRow).Value <> "" Then iRow = iRow + 1
            End With
        Else
            iPos = InStr(DataLine, "=")
            Field = Trim(Left(DataLine, iPos - 1))
            Value = Trim(Mid(DataLine, iPos + 1))
        End If
        ' put
        sCol = WhereIsColumn(sSheet, 1, Field)
        If sCol = "" Or sCol = "-" Then Stop
        With Workbooks(Main_Workbook_Name).Sheets(sSheet)
            .Range(sCol & iRow).Value = Value
        End With
        EnableEvents
    Wend
    Close #FileNum
    '============================
    ImportOPCini = True
End Function

Public Function ImportRSLinx(ByVal sFile As String, ByVal sGroup As String, ByRef sSheet As String) As Boolean
    InitVBs
    Dim s1 As String, sLink As String

    Dim obFile As New clsFile
    If InStr(sFile, ".txt") = 0 Then sFile = sFile & ".txt"

    obFile.Arq = sFile
    obFile.BasePath = Main_Workbook_Path
    obFile.SubPath = sGroup

    s1 = obFile.FullPath & "\" & obFile.FileAndExtension
    If Dir(s1) = "" Then
        PlayWAV "chord.wav"
        'MsgBox "Nao Existe " & s1
        LogSheet "ImportRSLinx", "Nao Existe " & s1, sSheet, "", A_Write
        ImportRSLinx = False
        Exit Function
    End If
    sSheet = sFile
    sSheet = Replace(sSheet, ".reg", "")
    sSheet = Replace(sSheet, ".txt", "")
    sSheet = sSheet & "." & sGroup
    NewSheet UCase(sSheet)

    Dim DataLine As String, FileNum As Long, iCh As Long, iRel As Long
    Dim sDummy As String, cData As Collection
    Dim sChave(0 To 39) As String, iPos(0 To 3) As Long, iPasso As Long
    '
    ReDim Colunas(0 To 4)
    Const eColDriver = 0
    Const eColNode = 1
    Const eColName = 2
    Const eColID = 3
    Const eColAddr = 4

    sChave(1) = "[HKEY_LOCAL_MACHINE\"
    sChave(10) = "[HKEY_LOCAL_MACHINE\SOFTWARE\Rockwell Software\RSLinx\Drivers\AB_ETH\"
    sChave(20) = "[HKEY_LOCAL_MACHINE\SOFTWARE\Rockwell Software\RSLinx\Drivers\TCP\"
    sChave(30) = "[HKEY_LOCAL_MACHINE\SOFTWARE\Rockwell Software\RSLinx\Project\Default\"
    'ETH
    sChave(11) = "\Node Table]"
    sChave(12) = "]"
    sChave(13) = Replace(".Name.=.", ".", DQ)
    sChave(14) = Replace(".#.=.", ".", DQ)
    'TCP
    sChave(21) = Replace(".Name.=.", ".", DQ)
    sChave(22) = Replace(".Server Name.=.", ".", DQ)
    sChave(23) = Replace(".Server IP Address.=.", ".", DQ)
    'OPC
    sChave(31) = Replace(".Target.=.", ".", DQ)
    sChave(32) = Replace(".HarmonyPath.=.", ".", DQ)

    Set cData = New Collection
    FileNum = FreeFile()
    '
    iPasso = 0
    Open s1 For Input As #FileNum
    While Not EOF(FileNum)
        Line Input #FileNum, DataLine
        DataLine = Trim(DataLine)
        DataLine = removeVBCRLF(DataLine)
        '========================================================================
        iCh = 1
        If Left(UCase(DataLine), Len(sChave(iCh))) = UCase(sChave(iCh)) Then iPasso = 0
        '========================================================================
        If iPasso = 0 Then
            iCh = 10
            If Left(UCase(DataLine), Len(sChave(iCh))) = UCase(sChave(iCh)) Then
                iPasso = 100
                iPos(0) = Len(sChave(iCh)) + 1
                iPos(1) = Len(DataLine) - iPos(0)
                Colunas(eColNode) = Mid(DataLine, iPos(0), iPos(1))
                'AB_ETH-1\Node Table
                Colunas(eColNode) = Replace(Colunas(eColNode), "\Node Table", "")
            End If
        End If
        If iPasso = 0 Then
            iCh = 20
            If Left(UCase(DataLine), Len(sChave(iCh))) = UCase(sChave(iCh)) Then
                iPasso = 200
                iPos(0) = Len(sChave(iCh)) + 1
                iPos(1) = Len(DataLine) - iPos(0)
                Colunas(eColNode) = Mid(DataLine, iPos(0), iPos(1))
            End If
        End If
        If iPasso = 0 Then
            iCh = 30
            If Left(UCase(DataLine), Len(sChave(iCh))) = UCase(sChave(iCh)) Then
                iPasso = 300
                iPos(0) = Len(sChave(iCh)) + 1
                iPos(1) = Len(DataLine) - iPos(0)
                Colunas(eColNode) = Mid(DataLine, iPos(0), iPos(1))
            End If
        End If
        '========================================================================
        If iPasso = 100 Then
            Colunas(eColDriver) = "ETH"
            iPos(0) = Len(sChave(iCh)) + 1
            '
            iCh = 10: iRel = 1
            If Right(UCase(DataLine), Len(sChave(iCh + iRel))) = UCase(sChave(iCh + iRel)) Then
                iPos(1) = Len(DataLine) - Len(sChave(iCh + iRel)) - 1
                Colunas(eColID) = 0
                iPasso = 101
            Else
                iCh = 10: iRel = 2
                If Right(UCase(DataLine), Len(sChave(iCh + iRel))) = UCase(sChave(iCh + iRel)) Then
                    iPos(1) = Len(DataLine) - Len(sChave(iCh + iRel)) - 1
                    iPasso = 102
                End If
            End If
        End If

        If iPasso = 101 Then
            iCh = 10: iRel = 4
            'Colunas(eColID)
            sDummy = sChave(iCh + iRel)
            For Colunas(eColID) = 0 To 64
            sChave(iCh + iRel) = Replace(sChave(iCh + iRel), "#", Colunas(eColID))
            If Left(UCase(DataLine), Len(sChave(iCh + iRel))) = UCase(sChave(iCh + iRel)) Then
                iPos(3) = Len(sChave(iCh + iRel)) + 1
                Colunas(eColAddr) = Mid(DataLine, iPos(3), Len(DataLine) - iPos(3))
                cData.add Colunas
                Exit For
                'Colunas(eColID) = Colunas(eColID) + 1
            End If
            sChave(iCh + iRel) = sDummy
            Next

            sChave(iCh + iRel) = sDummy

        End If
        If iPasso = 102 Then
            iCh = 10: iRel = 3
            If Left(UCase(DataLine), Len(sChave(iCh + iRel))) = UCase(sChave(iCh + iRel)) Then
                iPos(3) = Len(sChave(iCh + iRel)) + 1
                Colunas(eColName) = Mid(DataLine, iPos(3), Len(DataLine) - iPos(3))
            End If
        End If
        '========================================================================
        If iPasso = 200 Then
            Colunas(eColDriver) = "TCP"
            '
            iCh = 20: iRel = 1
            If Left(UCase(DataLine), Len(sChave(iCh + iRel))) = UCase(sChave(iCh + iRel)) Then
                iPos(1) = Len(sChave(iCh + iRel)) + 1
                iPos(2) = Len(DataLine) - Len(sChave(iCh + iRel)) - 1
                Colunas(eColName) = Mid(DataLine, iPos(1), iPos(2))
            End If
            iCh = 20: iRel = 2
            If Left(UCase(DataLine), Len(sChave(iCh + iRel))) = UCase(sChave(iCh + iRel)) Then
                iPos(1) = Len(sChave(iCh + iRel)) + 1
                iPos(2) = Len(DataLine) - Len(sChave(iCh + iRel)) - 1
                Colunas(eColID) = Mid(DataLine, iPos(1), iPos(2))
            End If
            iCh = 20: iRel = 3
            If Left(UCase(DataLine), Len(sChave(iCh + iRel))) = UCase(sChave(iCh + iRel)) Then
                iPos(1) = Len(sChave(iCh + iRel)) + 1
                iPos(2) = Len(DataLine) - Len(sChave(iCh + iRel)) - 1
                Colunas(eColAddr) = Mid(DataLine, iPos(1), iPos(2))
                '
                cData.add Colunas
                '
            End If
        End If
        '========================================================================
        If iPasso = 300 Then
            Colunas(eColDriver) = "OPC"
            '
            iCh = 30: iRel = 1
            If Left(UCase(DataLine), Len(sChave(iCh + iRel))) = UCase(sChave(iCh + iRel)) Then
                iPos(1) = Len(sChave(iCh + iRel)) + 1
                iPos(2) = Len(DataLine) - Len(sChave(iCh + iRel)) - 1
                Colunas(eColName) = Mid(DataLine, iPos(1), iPos(2))
                'RSLinx:!AA010B03\\A2:1.0,0,10,80
                Colunas(eColName) = Replace(Colunas(eColName), "RSLinx:!", "")
                iPos(3) = InStr(Colunas(eColName), "\\") - 1
                If iPos(3) >= 0 Then
                    Colunas(eColName) = Left(Colunas(eColName), iPos(3))
                End If
            End If
            iCh = 30: iRel = 2
            If Left(UCase(DataLine), Len(sChave(iCh + iRel))) = UCase(sChave(iCh + iRel)) Then
                iPos(1) = Len(sChave(iCh + iRel)) + 1
                iPos(2) = Len(DataLine) - Len(sChave(iCh + iRel)) - 1
                Colunas(eColID) = Mid(DataLine, iPos(1), iPos(2))
                iPos(3) = InStr(Colunas(eColID), "!") - 1
                Colunas(eColAddr) = Left(Colunas(eColID), iPos(3))
                Colunas(eColID) = Mid(Colunas(eColID), iPos(3) + 1)
                iPos(3) = InStr(Colunas(eColID), "\\") + 2
                Colunas(eColID) = Mid(Colunas(eColID), iPos(3))
                iPos(3) = InStr(Colunas(eColID), "\\") - 1
                If iPos(3) > 0 Then Colunas(eColID) = Left(Colunas(eColID), iPos(3))
                '
                sDummy = Colunas(eColAddr)
                Colunas(eColAddr) = Colunas(eColID)
                Colunas(eColID) = sDummy
                '
                cData.add Colunas
                '
            End If
        End If
        '========================================================================
        EnableEvents
    Wend
    Close #FileNum
    Dim iRow As Long, iCol As Long

    With Workbooks(Main_Workbook_Name).Sheets(sSheet)
        iRow = 1: iCol = 0
        iCol = iCol + 1: .Range(IdToColumn(iCol) & iRow) = "Driver"
        iCol = iCol + 1: .Range(IdToColumn(iCol) & iRow) = "Node"
        iCol = iCol + 1: .Range(IdToColumn(iCol) & iRow) = "Name"
        iCol = iCol + 1: .Range(IdToColumn(iCol) & iRow) = "ID"
        iCol = iCol + 1: .Range(IdToColumn(iCol) & iRow) = "Addr"

        iRow = iRow + 1
        For iCh = 1 To cData.Count
            For iRel = LBound(Colunas) To UBound(Colunas)
                sDummy = cData(iCh)(iRel)
                iCol = iRel + 1
                .Range(IdToColumn(iCol) & iRow) = sDummy
                EnableEvents
            Next iRel
            iRow = iRow + 1
            EnableEvents
        Next iCh
        '
        s1 = IdToColumn(LBound(Colunas) + 1) & ":" & IdToColumn(UBound(Colunas) + 1)
        .Columns(s1).AutoFit
        .Columns(s1).HorizontalAlignment = xlLeft
        .Tab.Color = RGBColorIndex.Teal
    End With
    Set cData = Nothing
    ImportRSLinx = True
End Function

Public Sub EnableEvents(Optional OnlyTone As Boolean)
    Dim i As Single
    Static Z As Single
    i = Timer
    If Z < i Then
        If Not OnlyTone Then DoEvents
        Z = i + 26#
        PlayWAV "Windows Battery Critical.wav"
    End If
End Sub

Public Function DesmontaPontoAnimacao(ByVal sPoint As String, ByVal ID As Long) As String
'\\MSWCPMC_ROS\MG1CPMC_ROS\ASIANDON_ROS_6[0] BAND 16
    Dim s1() As String, i1 As Long
    sPoint = Replace(sPoint, "\\", ";")
    sPoint = Replace(sPoint, "\", ";")
    sPoint = Replace(sPoint, "[", ";")
    sPoint = Replace(sPoint, "]", ";")
    sPoint = Replace(sPoint, "BAND", ";")
    sPoint = Replace(sPoint, " ", "")
    i1 = 0
    While Len(sPoint) <> i1
        i1 = Len(sPoint)
        sPoint = Replace(sPoint, ";;", ";")
        EnableEvents
    Wend
    While Left(sPoint, 1) = ";"
        sPoint = Mid(sPoint, 2)
        EnableEvents
    Wend
    While Right(sPoint, 1) = ";"
        sPoint = Left(sPoint, Len(sPoint) - 1)
        EnableEvents
    Wend
    s1() = Split(sPoint, ";")
    If ID >= LBound(s1) And ID <= UBound(s1) Then
        DesmontaPontoAnimacao = s1(ID)
    Else
        DesmontaPontoAnimacao = ""
    End If
End Function
'================================================================
Public Function AntiLog(nr As Long, base As Long) As Long
    Dim i As Single, j As Single
    i = Log(nr)
    j = Log(base)
    AntiLog = (i / j)
    If base ^ AntiLog <> nr Then
        'Stop
        'MsgBox "Function not designed for single data " & nr & "=>" & (i / j)
        'LogSheet "AntiLog", "Function not designed for single data nr -> (i / j)", nr, (i / j), LogActions.A_Write
    End If
End Function

Public Function AntiWord(nr As Long, base As Long) As Long
    AntiWord = nr \ base
End Function

Public Function AntiByte(nr As Long, base As Long) As Long
    Dim iW As Long
    iW = AntiWord(nr, base)
    AntiByte = nr - (iW * base)
End Function

Public Function Word(ByVal nr As Long) As Long
    '79=word (1279)
    Word = Freitas.AntiWord(nr, 16)
End Function
Public Function Bit(ByVal nr As Long) As Long
    '15=Bit(1279)
    Bit = Freitas.AntiByte(nr, 16)
End Function
Public Function WordBit(ByVal W As Long, ByVal B As Long) As Long
    '1279=WordBit(79,15)
    WordBit = W * 16 + B
End Function
'================================================================

Public Function GetConfig(sSheet As String, Parameter As String) As String
    If "" = Main_Workbook_Name Then SetMainWorkbookName
    '
    Dim iRow As Long
    iRow = 1

    With Workbooks(Main_Workbook_Name).Sheets(sSheet)
        While (UCase(.Range("A" & iRow).Value) <> UCase(Parameter)) And (.Range("A" & iRow).Value <> "")
            iRow = iRow + 1
            EnableEvents
        Wend
        If (UCase(.Range("A" & iRow).Value) = UCase(Parameter)) Then
            GetConfig = .Range("B" & iRow).Value
        End If
    End With
End Function

Public Function ReverseString(ByVal Text As String)
    ReverseString = StrReverse(Text)
End Function
Public Function FindLastChar(ByVal Text As String, ByVal Ch As String) As Long
    FindLastChar = InStrRev(Text, Ch, Len(Text))
End Function

Public Function FindColumnByRange(sSheet As String, iIni As Long, iEnd As Long, iRow As Long, sFind As String) As Long
    Dim iCol As Long
    If "" = Main_Workbook_Name Then InitVBs
    For iCol = iIni To iEnd
        With Workbooks(Main_Workbook_Name).Sheets(sSheet)
            If UCase(.Range(IdToColumn(iCol) & iRow).Value) = UCase(sFind) Then
                FindColumnByRange = iCol
                Exit Function
            End If
        End With
        EnableEvents
    Next iCol
End Function

Public Sub GetColumnByTable(sSheet As String, sTable As String, iIni As Long, iEnd As Long)
    '
    If Bugs Then LogSheet "GetColumnByTable", sTable, "", getLogType(enLog.Register), LogActions.A_Write
    '
    Dim iCol As Long
    iCol = 1
    iIni = 0
    '
    If "" = Main_Workbook_Name Then InitVBs
    With Workbooks(Main_Workbook_Name).Sheets(sSheet)
        While .Range(IdToColumn(iCol) & 2).Value <> ""
        If UCase(.Range(IdToColumn(iCol) & 1).Value) = UCase(sTable) Then iIni = iCol
        '
        If iIni <> 0 Then
            If iIni = iCol Then iCol = iCol + 1
            If UCase(.Range(IdToColumn(iCol) & 1).Value) <> "" Then
                iEnd = iCol - 1
                Exit Sub
            End If
        End If
        iCol = iCol + 1
        EnableEvents
        Wend
    End With
    '
    If iIni = 0 Then
        'nao existe
        LogSheet "GetColumnByTable", sTable & " in " & sSheet, "Not found", getLogType(enLog.Fault), LogActions.A_Write
    End If
    'passou pelo ultimo registro
    If iEnd = 0 Then iEnd = iCol - 1
End Sub

Public Sub ClearWork_Routine()
    '------------------------------
    ClearGlobalVariables
    '------------------------------
    If "" = Main_Workbook_Name Then InitVBs
    'If 0 = CallByUpperRoutine Then
        'InitVBs
        SheetUpdates NoneWork 'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
    'End If

    'preserve CMD and Definitions
    Dim No As Boolean
    Dim ws1 As Excel.Worksheet

    For Each ws1 In Workbooks(Main_Workbook_Name).Worksheets
        No = False
        If UCase(ws1.Name) = UCase(CommandTab) Then No = True
        If UCase(ws1.Name) = UCase(ConfigTab) Then No = True
        If Left(ws1.Name, 1) = "_" Then No = True
        If Not No Then DestroySpreadsheet Main_Workbook_Name, ws1.Name
        EnableEvents
    Next ws1
    '
    If Bugs Then LogSheet "ClearWork_Routine", "", "", getLogType(enLog.Register), LogActions.A_Write
    '
    'If 0 = CallByUpperRoutine Then
        SheetUpdates EnableFull   '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
        Call ShowFirstSheet
        PlayWAV "Speech Off.wav"
    'End If
End Sub

Public Function IsInsidePainted(rng As Range) As Boolean
    IsInsidePainted = (rng.Interior.Color <> RGBColorIndex.white)
End Function

Public Function base362Long(ByVal val As String) As Long
    Dim i As Long, p() As Long, n() As Long, vL(0 To 35) As String

    val = Trim(UCase(val))
    If val = "" Then Exit Function
    '
    FillBase36 vL()
    '
    ReDim p(0 To Len(val) - 1)
    For i = UBound(p) To LBound(p) Step -1
        p(i) = 36 ^ i
    Next i
    '
    ReDim n(0 To Len(val) - 1)
    For i = UBound(n) To LBound(n) Step -1
        n(i) = FindContent(vL(), Left(val, 1))
        val = Mid(val, 2)
    Next i
    base362Long = 0
    For i = UBound(n) To LBound(n) Step -1
        base362Long = base362Long + p(i) * n(i)
    Next i
End Function
Public Function Long2base36(val As Long, ByVal dig As Long) As String
    'A=65 Z=90 0=48 9=57
    Dim i As Long, p() As Long, n() As Long, vL(0 To 35) As String
    Dim Resto As Long, Calc As String
    dig = dig - 1
    If dig < 1 Then Exit Function
    FillBase36 vL()
    ReDim p(0 To dig)
    ReDim n(0 To dig)

    For i = dig To 0 Step -1
        p(i) = 36 ^ i
    Next i

    Resto = val
    For i = dig To 0 Step -1
        n(i) = Resto \ p(i)
        Resto = Resto - n(i) * p(i)
    Next i
    If Resto <> 0 Then Stop

    Calc = ""
    For i = dig To 0 Step -1
        Calc = Calc & vL(n(i))
    Next i
    Long2base36 = Calc
End Function
Private Sub FillBase36(ByRef v() As String)
    Dim i As Long
    For i = 0 To 9
        v(i) = i
    Next i
    For i = 65 To 90
        v(i - 65 + 10) = Chr(i)
    Next i
End Sub

Public Sub OneDrive(Action As SyncActions)
    Dim Program(0 To 0) As String, i1 As Long

    i1 = LBound(Program)
    Program(i1) = "cmd /c pssuspend" & IIf(Action = Release, " -r ", " ") & "GROOVE"
    'i1 = i1 + 1
    'Program(i1) = "cmd /c pssuspend" & IIf(Action = Release, " -r ", " ") & "MSOSYNC"

    For i1 = LBound(Program) To UBound(Program)
        Shell Program(i1), VbAppWinStyle.vbNormalNoFocus
        EnableEvents
    Next i1
End Sub

Public Sub SplitFixedAddress(ByVal add As String, ByRef cl As String, ByRef rl As Long)
    Const sep = "$"
    Dim m() As String
    If 0 = InStr(add, sep) Then
        cl = ""
        rl = -1
    Else
        '$D$5
        m() = Split(add, sep)
        cl = m(1)
        rl = m(2)
    End If
End Sub


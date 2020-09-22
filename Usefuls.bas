Attribute VB_Name = "Usefuls"
Option Explicit

'-------------------------------------------
' All code in this module is original, unless otherwise specified (or I can't remember who wrote it...)
' It tends to get copied into any project of a reasonable size that I create.
'   - FireClaw.  bigcalm@hotmail.com
'-------------------------------------------

' Compiler Directives
'#Const Vba6 = False

'-------------------------------------------
' Timing Declares
'-------------------------------------------
Public Type LongLong ' Unsigned 64-bit long
    LowPart As Long
    HighPart As Long
End Type

Declare Function QueryPerformanceCounter Lib "kernel32" _
                (lpPerformanceCount As LongLong) As Long

Declare Function QueryPerformanceFrequency Lib "kernel32" _
                (lpFrequency As LongLong) As Long
Declare Function timeGetTime Lib "winmm.dll" () As Long

'-------------------------------------------
' ODBC stuff
'-------------------------------------------
Declare Function SQLGetStmtOption Lib "odbc32.dll" (ByVal hstmt As Long, ByVal fOption As Integer, ByRef pvParam As Long) As Integer
Global Const SQL_QUERY_TIMEOUT = 0
Global Const SQL_MAX_ROWS = 1
Global Const SQL_NOSCAN = 2
Global Const SQL_MAX_LENGTH = 3
Global Const SQL_ASYNC_ENABLE = 4
Global Const SQL_BIND_TYPE = 5
Global Const SQL_CURSOR_TYPE = 6
Global Const SQL_CONCURRENCY = 7
Global Const SQL_KEYSET_SIZE = 8
Global Const SQL_ROWSET_SIZE = 9
Global Const SQL_SIMULATE_CURSOR = 10
Global Const SQL_RETRIEVE_DATA = 11
Global Const SQL_USE_BOOKMARKS = 12
Global Const SQL_GET_BOOKMARK = 13
Global Const SQL_ROW_NUMBER = 14
Global Const SQL_GET_ROWID = 1048
Global Const SQL_GET_SERIALNO = 1049

'-------------------------------------------
' Windows Messaging Stuff
'-------------------------------------------
Type POINTAPI
        X As Long
        Y As Long
End Type
Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    PT As POINTAPI
End Type
Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
Public Const PM_NOREMOVE = &H0
Public Const PM_NOYIELD = &H2
Public Const PM_REMOVE = &H1

'-------------------------------------------
' Windows Graphics API Calls
'-------------------------------------------
Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long

'-------------------------------------------
' ClipBoard Stuff
'-------------------------------------------
' Memory library calls
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
                                                                ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpDest As Any, _
    lpSource As Any, _
    ByVal cbCopy As Long)
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
' Clipboard Function calls
Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Declare Function EmptyClipboard Lib "user32" () As Long
Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Declare Function GetClipboardOwner Lib "user32" () As Long
Declare Function GetClipboardViewer Lib "user32" () As Long

' Memory constants
Public Const GMEM_SHARE = &H2000
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_ZEROINIT = &H40
Public Const FOR_CLIPBOARD = GMEM_MOVEABLE Or GMEM_SHARE Or GMEM_ZEROINIT

' Clipboard format types and constants
Public Enum ClipBoardFormats
    CF_ANSIONLY = &H400&
    CF_APPLY = &H200&
    CF_BITMAP = 2
    CF_DIB = 8
    CF_DIF = 5
    CF_DSPBITMAP = &H82
    CF_DSPENHMETAFILE = &H8E
    CF_DSPMETAFILEPICT = &H83
    CF_DSPTEXT = &H81
    CF_EFFECTS = &H100&
    CF_ENABLEHOOK = &H8&
    CF_ENABLETEMPLATE = &H10&
    CF_ENABLETEMPLATEHANDLE = &H20&
    CF_ENHMETAFILE = 14
    CF_FIXEDPITCHONLY = &H4000&
    CF_FORCEFONTEXIST = &H10000
    CF_GDIOBJFIRST = &H300
    CF_GDIOBJLAST = &H3FF
    CF_INITTOLOGFONTSTRUCT = &H40&
    CF_LIMITSIZE = &H2000&
    CF_METAFILEPICT = 3
    CF_NOFACESEL = &H80000
    CF_NOSCRIPTSEL = &H800000
    CF_NOSIMULATIONS = &H1000&
    CF_NOSIZESEL = &H200000
    CF_NOSTYLESEL = &H100000
    CF_NOVECTORFONTS = &H800&
    CF_NOOEMFONTS = CF_NOVECTORFONTS
    CF_NOVERTFONTS = &H1000000
    CF_OEMTEXT = 7
    CF_OWNERDISPLAY = &H80
    CF_PALETTE = 9
    CF_PENDATA = 10
    CF_PRINTERFONTS = &H2
    CF_PRIVATEFIRST = &H200
    CF_PRIVATELAST = &H2FF
    CF_RIFF = 11
    CF_SCALABLEONLY = &H20000
    CF_SCREENFONTS = &H1
    CF_SCRIPTSONLY = CF_ANSIONLY
    CF_SELECTSCRIPT = &H400000
    CF_SHOWHELP = &H4&
    CF_SYLK = 4
    CF_TEXT = 1
    CF_TIFF = 6
    CF_TTONLY = &H40000
    CF_UNICODETEXT = 13
    CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
End Enum

'-------------------------------------------
' My own constants and Enums
'-------------------------------------------
Private Const CntrlToken = "#"  ' For Load/Save Form

' Enum for length unit conversions
Public Enum LengthUnits
    ' Metric
    Micrometres = 1 ' 0.001mm
    Milimetres = 2
    Centimetres = 3 ' 10mm
    Metres = 4 ' 100cm
    Kilometres = 5 ' 1000m
    ' Common Imperial
    Inches = 6 ' 25.4 milimetres
    Feet = 7 ' 12 inches
    Yards = 8 ' 3 Feet
    Miles = 9 ' 1760 yards
    ' Nautical and Horse racing
    NauticalMiles = 10 ' 6080 yards
    CableLengths = 11 ' 600 feet
    Chains = 12 ' Gunters Chain: 66 feet
    Fathoms = 13 ' 6 feet
    Furlongs = 14 ' 660 feet or 10 chains
    Hands = 15 ' 4 inches
    Degrees = 16 ' 1/360th of earth circumference
    Minutes = 17 ' 1/60th of a degree, or one nautical mile
    Seconds = 18 ' 1/60th of a minute, or 1/60th of a nautical mile
    ' Computer
    Dots = 19 ' 1/300th of an inch (printing)
    Points = 20 ' 1/72nd of an inch (fonts)
    RadixDots = 21 ' 1/4 of a dot (bitmap font design)
    Twips = 22 ' 1/1440th of an inch (screen measure)
    PlotterUnits = 23 ' 1/1016th of an inch (printing)
    ' Scientific
'   Angstroms = 24 ' Tiny tiny unit.  Commented because unsure about actual value
    LightYears = 25 ' 9.4 * 10^15 metres
    ' Old and Biblical
    Cubits = 26 ' 18 inches
    RoyalEgyptianCubits = 27 ' 21 inches
    Ells = 28 ' 45 inches
    Palms = 29 ' 127mm
    Reeds = 30 ' 1520mm
    Span = 31 ' 9 inches
End Enum

'-------------------------------------------
' Modular Variables
'-------------------------------------------

' For split string purposes
Private mSplitLine As String ' These three vars are used to
Private mDelimiter As String ' split a delimiter seperated line up
Private mCurrentPos As Long

'-------------------------------------------
' String handling functions
'-------------------------------------------
Public Sub SplitStringIntoParts(pLine As String, pDelimiter)
    mSplitLine = pLine
    mDelimiter = pDelimiter
    mCurrentPos = 1
End Sub

Public Function GetNextPartOfSplitString() As String
Dim lCurrentPos As Long
    If mCurrentPos > Len(mSplitLine) Then
        GetNextPartOfSplitString = ""
    Else
        lCurrentPos = InStr(mCurrentPos, mSplitLine, mDelimiter)
        If lCurrentPos = 0 Then
            ' Get rest of line
            GetNextPartOfSplitString = Mid(mSplitLine, mCurrentPos, (Len(mSplitLine) - mCurrentPos) + 1)
            mCurrentPos = Len(mSplitLine) + 1
        Else
            GetNextPartOfSplitString = Mid(mSplitLine, mCurrentPos, (lCurrentPos - mCurrentPos))
            mCurrentPos = lCurrentPos + Len(mDelimiter)
        End If
    End If
End Function

Public Function RightJustifyCurrencyToString(Value As Currency, Optional Padding As Long = 10, Optional FailureString As String = "") As String
Dim tmpStr As String
Dim i As Long
    tmpStr = Format(Value, "0.00")
    If Padding - Len(tmpStr) < 0 Then
        If Len(FailureString) = 0 Then
            RightJustifyCurrencyToString = ""
            For i = 1 To Padding
                RightJustifyCurrencyToString = RightJustifyCurrencyToString & "#"
            Next
        Else
            RightJustifyCurrencyToString = FailureString
        End If
    Else
        RightJustifyCurrencyToString = Space(Padding - Len(tmpStr)) & tmpStr
    End If
End Function

' Translates into "Database field friendly" format
Public Function QuoteX2(pString As String) As String
Dim lPos As Long
Dim lNewString As String

    ' if it contains a quote, we need to substitute this with ""
    Trim (pString)
    If Len(pString) = 0 Then
        QuoteX2 = ""
        Exit Function
    End If
    If Len(pString) = 1 Then
        If pString = Chr(34) Then
            QuoteX2 = Chr(34) & Chr(34) & Chr(34) & Chr(34)
            Exit Function
        End If
    End If
    lNewString = Chr(34)
    For lPos = 1 To Len(pString)
        If Mid(pString, lPos, 1) = Chr(34) Then
            lNewString = lNewString & Chr(34)
        End If
        lNewString = lNewString & Mid(pString, lPos, 1)
    Next
    lNewString = lNewString & Chr(34)
    QuoteX2 = Trim(lNewString)
End Function

Private Function ConvertStringToValidCSVFormat(ByVal pString As String) As String
Dim lPos As Long
Dim lNewString As String

    ' if it contains a quote, we need to substitute this with ""
    If Len(pString) = 0 Then
        ConvertStringToValidCSVFormat = ""
        Exit Function
    End If
    If Len(pString) = 1 Then
        If pString = Chr(34) Then
            ConvertStringToValidCSVFormat = Chr(34) & Chr(34) & Chr(34) & Chr(34)
            Exit Function
        End If
    End If
    lNewString = Chr(34)
    For lPos = 1 To Len(pString)
        If Mid(pString, lPos, 1) = Chr(34) Then
            lNewString = lNewString & Chr(34)
        End If
        lNewString = lNewString & Mid(pString, lPos, 1)
    Next
    lNewString = lNewString & Chr(34)
    ConvertStringToValidCSVFormat = lNewString
End Function

' Useful when retrieving rows from database
Public Function GRON(Var As Variant) As String
    If IsNull(Var) Then
        GRON = ""
    Else
        GRON = Var
    End If
End Function

' Search and Replace
Public Function QSAR(ByVal pString As String, ByVal pSearch As String, Optional ByVal pReplace As String = "", Optional pCompare As Long = vbBinaryCompare, Optional GlobalReplace As Boolean = True) As String
Dim lLen1 As Long
Dim lLen2 As Long
Dim lStartFind As Long
Dim lFoundLoc As Long
Dim ltmpString As String

    lLen1 = Len(pString)
    lLen2 = Len(pSearch)
    lStartFind = 1
    ltmpString = ""
    Do
        lFoundLoc = InStr(lStartFind, pString, pSearch, pCompare)
        If lFoundLoc = 0 Then
            Exit Do
        End If
        ltmpString = ltmpString & Mid(pString, lStartFind, lFoundLoc - lStartFind) & pReplace
        If lStartFind = 1 And GlobalReplace = False Then
            lStartFind = lFoundLoc + lLen2
            Exit Do
        End If
        lStartFind = lFoundLoc + lLen2
    Loop
    ltmpString = ltmpString & Mid(pString, lStartFind, lLen1 - lStartFind + 1)
    QSAR = ltmpString
End Function

' Note that this ONLY supports the formats %s, \n and \t
Public Function PrintF(FormatString As String, ParamArray PA() As Variant)
Dim Param As Variant
Dim OutputString As String
    OutputString = FormatString
    For Each Param In PA
        OutputString = QSAR(OutputString, "%s", Param, , False)
    Next
    OutputString = QSAR(OutputString, "\n", vbCrLf, , True)
    OutputString = QSAR(OutputString, "\t", vbTab, , True)
    Debug.Print OutputString
End Function

' Note that this ONLY supports the formats %s, \n and \t
Public Function FPrintF(FileNumber As Long, FormatString As String, ParamArray PA() As Variant)
Dim Param As Variant
Static OutputString As String
    OutputString = OutputString & FormatString
    For Each Param In PA
        OutputString = QSAR(OutputString, "%s", Param, , False)
    Next
    OutputString = QSAR(OutputString, "\n", vbCrLf, , True)
    OutputString = QSAR(OutputString, "\t", vbTab, , True)
    If Right(OutputString, 2) = vbCrLf Then
        Print #FileNumber, Mid(OutputString, 1, Len(OutputString) - 2)
        OutputString = ""
    End If
End Function

'-------------------------------------------
' Conversion functions
'-------------------------------------------
Public Function LengthUnitConvert(InitialValue As Double, InitialUnit As LengthUnits, FinalUnit As LengthUnits) As Double
Dim Mili As Double
    Select Case InitialUnit
        ' Metric
        Case Micrometres
            Mili = InitialValue * 0.001
        Case Milimetres
            Mili = InitialValue
        Case Centimetres
            Mili = InitialValue * 10
        Case Metres
            Mili = InitialValue * 1000
        Case Kilometres
            Mili = InitialValue * 1000000
        ' Common Imperial
        Case Inches
            Mili = InitialValue * 25.4
        Case Feet
            Mili = InitialValue * 25.4 * 12
        Case Yards
            Mili = InitialValue * 25.4 * 36
        Case Miles
            Mili = InitialValue * 25.4 * 36 * 1760
        ' Nautical and horse racing
        Case NauticalMiles
            Mili = InitialValue * 25.4 * 36 * 6080
        Case CableLengths
            Mili = InitialValue * 25.4 * 12 * 600
        Case Chains
            Mili = InitialValue * 25.4 * 12 * 66
        Case Fathoms
            Mili = InitialValue * 25.4 * 12 * 6
        Case Furlongs
            Mili = InitialValue * 25.4 * 12 * 660
        Case Hands
            Mili = InitialValue * 25.4 * 4
        Case Degrees
            Mili = InitialValue * 25.4 * 36 * 6080 * 60
        Case Minutes
            Mili = InitialValue * 25.4 * 36 * 6080 ' yes, same as nautical mile
        Case Seconds
            Mili = InitialValue * 25.4 * 36 * (6080 / 60)
        ' Computer
        Case Dots
            Mili = InitialValue * 25.4 / 300
        Case Points
            Mili = InitialValue * 25.4 / 72
        Case RadixDots
            Mili = InitialValue * 25.4 / 1200
        Case Twips
            Mili = InitialValue * 25.4 / 1440
        Case PlotterUnits
            Mili = InitialValue * 25.4 / 1016
        ' Scientific
'        Case Angstroms
'            Mili = InitialValue * 1 / 10000000000#
        Case LightYears
            Mili = InitialValue * 1000 * 9.4 * 10 ^ 15
        ' Old and Biblical
        Case Cubits
            Mili = InitialValue * 25.4 * 18
        Case RoyalEgyptianCubits
            Mili = InitialValue * 25.4 * 21
        Case Ells
            Mili = InitialValue * 25.4 * 45
        Case Palms
            Mili = InitialValue * 127
        Case Reeds
            Mili = InitialValue * 1520
        Case Span
            Mili = InitialValue * 25.4 * 9
    End Select
    Select Case FinalUnit
        ' Metric
        Case Micrometres
            LengthUnitConvert = Mili / 0.001
        Case Milimetres
            LengthUnitConvert = Mili
        Case Centimetres
            LengthUnitConvert = Mili / 10
        Case Metres
            LengthUnitConvert = Mili / 1000
        Case Kilometres
            LengthUnitConvert = Mili / 1000000
        ' Common Imperial
        Case Inches
            LengthUnitConvert = Mili / 25.4
        Case Feet
            LengthUnitConvert = Mili / (25.4 * 12)
        Case Yards
            LengthUnitConvert = Mili / (25.4 * 36)
        Case Miles
            LengthUnitConvert = Mili / (25.4 * 36 * 1760)
        ' Nautical and horse racing
        Case NauticalMiles
            LengthUnitConvert = Mili / (25.4 * 36 * 6080)
        Case CableLengths
            LengthUnitConvert = Mili / (25.4 * 12 * 600)
        Case Chains
            LengthUnitConvert = Mili / (25.4 * 12 * 66)
        Case Fathoms
            LengthUnitConvert = Mili / (25.4 * 12 * 6)
        Case Furlongs
            LengthUnitConvert = Mili / (25.4 * 12 * 660)
        Case Hands
            LengthUnitConvert = Mili / (25.4 * 4)
        Case Degrees
            LengthUnitConvert = Mili / (25.4 * 36 * 6080 * 60)
        Case Minutes
            LengthUnitConvert = Mili / (25.4 * 36 * 6080)  ' yes, same as nautical mile
        Case Seconds
            LengthUnitConvert = Mili / (25.4 * 36 * (6080 / 60))
        ' Computer
        Case Dots
            LengthUnitConvert = Mili / (25.4 / 300)
        Case Points
            LengthUnitConvert = Mili / (25.4 / 72)
        Case RadixDots
            LengthUnitConvert = Mili / (25.4 / 1200)
        Case Twips
            LengthUnitConvert = Mili / (25.4 / 1440)
        Case PlotterUnits
            LengthUnitConvert = Mili / (25.4 / 1016)
        ' Scientific
'        Case Angstroms
'            LengthUnitConvert = Mili / (1 / 10000000000#)
        Case LightYears
            LengthUnitConvert = Mili / (1000 * 9.4 * 10 ^ 15)
        ' Old and Biblical
        Case Cubits
            LengthUnitConvert = Mili / (25.4 * 18)
        Case RoyalEgyptianCubits
            LengthUnitConvert = Mili / (25.4 * 21)
        Case Ells
            LengthUnitConvert = Mili / (25.4 * 45)
        Case Palms
            LengthUnitConvert = Mili / 127
        Case Reeds
            LengthUnitConvert = Mili / 1520
        Case Span
            LengthUnitConvert = Mili / (25.4 * 9)
    End Select
End Function

'-------------------------------------------
' Timing functions
'-------------------------------------------
' Function can time accurately to microseconds (1/1000000th of a second)
' Tis slow though.  Have to convert 64bit unsigned integer to Decimal within Variant. Yuk
' When VB7 arrives, with it's 64bit long variable, I'll be able to write this to be a tad quicker(!)
Public Function TimerElapsed(Optional µS As Long = 0, Optional UsePerformanceTimer As Boolean = True) As Boolean
Static StartTime As Variant ' Decimal
Static PerformanceFrequency As LongLong
Static EndTime As Variant ' Decimal
Dim CurrentTime As LongLong
Dim Dec As Variant

    If µS > 0 Then
        ' Initialize
        If UsePerformanceTimer = True Then
            If QueryPerformanceFrequency(PerformanceFrequency) Then
                ' Performance Timer available
                If QueryPerformanceCounter(CurrentTime) Then
                Else
                    ' Performance timer is available, but is not responding
                    CurrentTime.HighPart = 0
                    CurrentTime.LowPart = timeGetTime
                    PerformanceFrequency.HighPart = 0
                    PerformanceFrequency.LowPart = 1000
                End If
            Else
                ' Performance timer is not available.
                CurrentTime.HighPart = 0
                CurrentTime.LowPart = timeGetTime
                PerformanceFrequency.HighPart = 0
                PerformanceFrequency.LowPart = 1000
            End If
        Else
                ' Do not need to use performance timer
                CurrentTime.HighPart = 0
                CurrentTime.LowPart = timeGetTime
                PerformanceFrequency.HighPart = 0
                PerformanceFrequency.LowPart = 1000
        End If
        ' Work out start time...
        ' Convert to DECIMAL
        Dec = CDec(CurrentTime.LowPart)
        ' make this UNSIGNED
        If Dec < 0 Then
            Dec = CDec(Dec + (2147483648# * 2))
        End If
        ' Add higher value
        StartTime = CDec(Dec + (CurrentTime.HighPart * 2147483648# * 2))
        
        ' Put performance frequency into Dec variable
        Dec = CDec(PerformanceFrequency.LowPart)
        ' Convert to unsigned
        If Dec < 0 Then
            Dec = CDec(Dec + (2147483648# * 2))
        End If
        ' Add higher value
        Dec = CDec(Dec + (PerformanceFrequency.HighPart * 2147483648# * 2))
        
        ' Work out end time from this
        EndTime = CDec(StartTime + µS * Dec / 1000000)
        TimerElapsed = False
    Else
        If PerformanceFrequency.LowPart = 1000 And PerformanceFrequency.HighPart = 0 Then
            ' Using standard windows timer
            Dec = CDec(timeGetTime)
            If Dec < 0 Then
                Dec = CDec(Dec + (2147483648# * 2))
            End If
            If Dec > EndTime Then
                TimerElapsed = True
            Else
                TimerElapsed = False
            End If
        Else
            If QueryPerformanceCounter(CurrentTime) Then
                Dec = CDec(CurrentTime.LowPart)
                ' make this UNSIGNED
                If Dec < 0 Then
                    Dec = CDec(Dec + (2147483648# * 2))
                End If
                Dec = CDec(Dec + (CurrentTime.HighPart * 2147483648# * 2))
                If Dec > EndTime Then
                    TimerElapsed = True
                Else
                    TimerElapsed = False
                End If
            Else
                ' Should never happen in theory
                Err.Raise vbObjectError + 2, "Timer Elapsed", "Your performance timer has stopped functioning!!!"
                TimerElapsed = True
            End If
        End If
    End If
End Function

'-------------------------------------------
' File handling functions
'-------------------------------------------
' simple check if a file exists
Public Function FileExists(Path As String) As Boolean
    FileExists = Len(Dir(Path)) > 0
End Function

'-------------------------------------------
' "Is" functions
'-------------------------------------------
Public Function IsOdd(Num As Long) As Boolean
    IsOdd = -(Num And 1)
End Function

Public Function IsEven(Num As Long) As Boolean
    IsEven = ((Num And 1) = 0)
End Function

Public Function IsDivisible(Numerator As Long, Divisor As Long) As Boolean
    IsDivisible = (Numerator Mod Divisor = 0) ' Credit to Ulli on PSC here for this
End Function

' Detects whether a control is part of a control array
Function IsControlArray(Cntrl As Control) As Boolean
    On Error GoTo ErrHandler
    If Cntrl.Index Then ' If control is not an array, then error 343 is thrown here
    End If
    IsControlArray = True
    Exit Function
ErrHandler:
    If Err.Number = 343 Then ' object is not an array
        IsControlArray = False
        Exit Function
    Else ' any other error
        IsControlArray = False
        Exit Function
    End If
End Function

' Special Asynchronous Functions
' Processes all events to be raised to a specific control (such as Click, KeyDown, etc.)
' Should generally be faster than the more generic DoEvents.  However, dangerous if
' you don't know what you're doing.
Public Sub DoEventsForControl(hwnd As Long)
Dim tmpMsg As MSG
    Do While PeekMessage(tmpMsg, hwnd, 0, 0, PM_REMOVE)
        TranslateMessage tmpMsg
        DispatchMessage tmpMsg
    Loop
End Sub

'-------------------------------------------
' Print Engine functions
'-------------------------------------------
Public Sub PrintEngineCentreText(Text As String)
Dim TW As Long
    With Printer
        TW = .TextWidth(Text)
        .CurrentX = (.Width - TW) / 2
        Printer.Print Text
    End With
End Sub

Public Sub PrintEnginePrintAt(Text As String, Optional X As Long = -1, Optional Y As Long = -1)
    With Printer
        If X >= 0 Then
            .CurrentX = X
        End If
        If Y >= 0 Then
            .CurrentY = Y
        End If
        Printer.Print Text
    End With
End Sub

Public Sub PrintEngineSkipLines(Optional ByVal NumberOfLines As Long = 1)
    With Printer
        While NumberOfLines > 0
            NumberOfLines = NumberOfLines - 1
            Printer.Print ""
        Wend
    End With
End Sub

'-------------------------------------------
' Collision Detection (Sprites)
'-------------------------------------------
' Acknowledgement here goes to Richard Lowe (riklowe@hotmail.com) for his collision detection
' algorithm which I have used as the basis of my collision detection algorithm.  Some of the logic in
' here is radically different though, and his algorithm originally didn't deallocate memory properly ;-)
Public Function CollisionDetect(ByVal x1 As Long, ByVal y1 As Long, ByVal X1Width As Long, ByVal Y1Height As Long, _
    ByVal Mask1LocX As Long, ByVal Mask1LocY As Long, ByVal Mask1Hdc As Long, ByVal x2 As Long, ByVal y2 As Long, _
    ByVal X2Width As Long, ByVal Y2Height As Long, ByVal Mask2LocX As Long, ByVal Mask2LocY As Long, _
    ByVal Mask2Hdc As Long) As Boolean
' I'm going to use RECT types to do this, so that the Windows GDI can do the hard bits for me.
Dim MaskRect1 As RECT
Dim MaskRect2 As RECT
Dim DestRect As RECT
Dim i As Long
Dim j As Long
Dim Collision As Boolean
Dim MR1SrcX As Long
Dim MR1SrcY As Long
Dim MR2SrcX As Long
Dim MR2SrcY As Long
Dim hNewBMP As Long
Dim hPrevBMP As Long
Dim tmpObj As Long
Dim hMemDC As Long


    MaskRect1.Left = x1
    MaskRect1.Top = y1
    MaskRect1.Right = x1 + X1Width
    MaskRect1.Bottom = y1 + Y1Height
    MaskRect2.Left = x2
    MaskRect2.Top = y2
    MaskRect2.Right = x2 + X2Width
    MaskRect2.Bottom = y2 + Y2Height
    i = IntersectRect(DestRect, MaskRect1, MaskRect2)
    If i = 0 Then
        CollisionDetect = False
    Else
        ' The two rectangles intersect, so let's go to a pixel by pixel comparison
        
        ' Set SourceX and Y values for both Mask HDC's...
        If x1 <= x2 Then
            MR1SrcX = X1Width - (DestRect.Right - DestRect.Left)
            MR2SrcX = 0
        Else
            MR1SrcX = 0
            MR2SrcX = X2Width - (DestRect.Right - DestRect.Left)
        End If
        If y1 <= y2 Then
            MR1SrcY = Y1Height - (DestRect.Bottom - DestRect.Top)
            MR2SrcY = 0
        Else
            MR1SrcY = 0
            MR2SrcY = Y2Height - (DestRect.Bottom - DestRect.Top)
        End If
        
        ' Allocate memory DC and Bitmap in which to do the comparison
        hMemDC = CreateCompatibleDC(Screen.ActiveForm.hdc)
        hNewBMP = CreateCompatibleBitmap(Screen.ActiveForm.hdc, DestRect.Right - DestRect.Left, DestRect.Bottom - DestRect.Top)
        hPrevBMP = SelectObject(hMemDC, hNewBMP)

        ' Blit the first sprite into it
        i = BitBlt(hMemDC, 0, 0, DestRect.Right - DestRect.Left, DestRect.Bottom - DestRect.Top, _
                Mask1Hdc, MR1SrcX + Mask1LocX, MR1SrcY + Mask1LocY, vbSrcCopy)

        ' Logical OR the second sprite with the first sprite
         i = BitBlt(hMemDC, 0, 0, DestRect.Right - DestRect.Left, DestRect.Bottom - DestRect.Top, _
                Mask2Hdc, MR2SrcX + Mask2LocX, MR2SrcY + Mask2LocY, vbSrcPaint)
        
        Collision = False
        For i = 0 To DestRect.Bottom - DestRect.Top - 1
            For j = 0 To DestRect.Right - DestRect.Left - 1
                If GetPixel(hMemDC, j, i) = 0 Then ' If there are any black pixels
                    Collision = True
                    Exit For
                End If
            Next
            If Collision = True Then
                Exit For
            End If
        Next
        CollisionDetect = Collision
        
        ' Destroy any allocated objects and DC's
        tmpObj = SelectObject(hMemDC, hPrevBMP)
        tmpObj = DeleteObject(tmpObj)
        tmpObj = DeleteDC(hMemDC)
    End If
End Function

Public Function PadHexStr(Str As String, Optional PadWidth As Long = 2) As String
Dim i As Long
    i = Len(Str)
    If i < PadWidth Then
        PadHexStr = RepeatChar("0", PadWidth - i) & Str
    Else
        PadHexStr = Str
    End If
End Function

Public Function FourBytesToLong(PB1 As Byte, pb2 As Byte, pb3 As Byte, pb4 As Byte) As Long
    FourBytesToLong = LshL(PB1, 24) Or LshL(pb2, 16) Or LshL(pb3, 8) Or pb4 ' I HATE I HATE I HATE VISUAL BASIC!!!!!
End Function

Public Function RepeatChar(pChar As String, pTimes As Long) As String
Dim i As Long
    For i = 1 To pTimes
        RepeatChar = RepeatChar & pChar
    Next
End Function

' Yuk.
Public Function HexStrToLong(Str As String) As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim B As Long

    j = 28
    k = 1
    B = 0
    Do While j >= 0
        i = Asc(Mid(Str, k, 1))
        If i >= 48 And i <= 57 Then
            B = B Or LshL(i - 48, j)
        ElseIf i >= 65 And i <= 70 Then
            B = B Or LshL(i - 65 + 10, j)
        ElseIf i >= 97 And i <= 102 Then
            B = B Or LshL(i - 97 + 10, j)
        Else
            Err.Raise 1, "HexStrToLong", "Invalid Hex String Specified": Exit Function
        End If
        k = k + 1
        j = j - 4
    Loop
    HexStrToLong = B
End Function

' Translates a string such as '000000000000000000110101' to a long.
Public Function BinStrToLong(Str As String) As Long
    
End Function

' Translates hex string such as "0A" or "Fe" or "70" to a byte value.  String must be 2 chars or you'll get an error back.
Public Function HexStrToByte(Str As String) As Byte
Dim i As Byte
Dim B As Byte
    On Error GoTo ErrHandler
    i = Asc(Mid(Str, 1, 1))
    If i >= 48 And i <= 57 Then
        B = BshL(i - 48, 4)
    ElseIf i >= 65 And i <= 70 Then
        B = BshL(i - 65 + 10, 4)
    ElseIf i >= 97 And i <= 102 Then
        B = BshL(i - 97 + 10, 4)
    Else
        Err.Raise 1, "HexStrToByte", "Invalid Hex String Specified": Exit Function
    End If
    
    i = Asc(Mid(Str, 2, 1))
    If i >= 48 And i <= 57 Then
        B = B Xor (i - 48)
    ElseIf i >= 65 And i <= 70 Then
        B = B Xor (i - 65 + 10)
    ElseIf i >= 97 And i <= 102 Then
        B = B Xor (i - 97 + 10)
    Else
        Err.Raise 1, "HexStrToByte", "Invalid Hex String Specified": Exit Function
    End If
    HexStrToByte = B
    Exit Function
ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    HexStrToByte = 0
    Exit Function
End Function

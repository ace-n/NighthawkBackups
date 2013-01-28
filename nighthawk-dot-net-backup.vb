' NOTE: Windows will only have one function executed upon them per scan
' NOTE: Add checks to insure that IE pages can be accurately referenced (aren't closed/deleted/etc...)

Imports System
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Net
Imports SHDocVw
Imports System.ComponentModel

Public Delegate Function CallBack(ByVal hwnd As Integer, ByVal lParam As Integer) As Boolean

Public Class Form1
    ' List of target windows (AllWindowsList)
    Dim AllWindowsList As New List(Of TargetWindowPlain)
    Dim IETabList As New List(Of Integer)

    ' DLL Functions used
    Declare Function GetWindow Lib "user32" (ByVal hWnd As Integer, ByVal uCmd As Integer) As Integer
    Declare Function EnumWindows Lib "user32" (ByVal CallBackPtr As CallBack, ByVal lParam As Integer) As Boolean
    Declare Function EnumChildWindows Lib "user32" (ByVal Hwnd As Integer, ByVal CallBackPtr As CallBack, ByVal lParam As Integer) As Boolean
    Declare Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As Integer, ByVal dwObjectID As Int32, ByVal riid As Guid, <MarshalAs(UnmanagedType.IUnknown)> ByRef ppvObject As Object) As String
    Declare Function IsWindowVisible Lib "user32" (ByVal Hwnd As Integer) As Boolean
    Declare Function GetClass Lib "user32" Alias "RealGetWindowClass" (ByVal hWnd As IntPtr, ByVal pszType As String, ByVal cchType As IntPtr) As UInteger
    Declare Function DestroyWindow Lib "user32" (ByVal Hwnd As Integer)

    ' Global variable definitions
    Dim RulesPath As String = My.Settings.EveryoneRules
    Dim CodeFile As New StreamReader(RulesPath)
    Dim CodeLineList As New List(Of CodeLine)
    Dim ChildWinArray As Array
    Dim CodeLine As String
    Dim WinArray As Array
    Dim CodeIsIfLine As Boolean
    Dim IEListFiltered As New List(Of InternetExplorer)
    Dim CodeListOfWindows()() As TargetWindowPlain = New TargetWindowPlain(10)() {}         ' This one keeps track of the recorded windows in each level

    Public Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' Get Nighthawk rules code (raw)
        Dim UserRank As Integer = 1 ' --> UPDATE!
        If UserRank = 5 Then
            RulesPath = My.Settings.SuperRules
        End If
        If UserRank = 4 Then
            RulesPath = My.Settings.HeadRules
        End If
        If UserRank = 3 Then
            RulesPath = My.Settings.AdminRules
        End If
        If UserRank = 2 Then
            RulesPath = My.Settings.RespectedRules
        End If
        If UserRank = 1 Then
            RulesPath = My.Settings.EveryoneRules
        End If

        ' Create list of code structures - this will be used on each parse
        '   Not querying the rules file each parse and using predefined values (instead of recalculating the values each parse) drastically increases efficiency
        '   The drawback is that the parsing rules won't change unless Nighthawk is re-launched
        Do Until CodeFile.EndOfStream = True

            ' Get current line
            CodeLine = CodeFile.ReadLine

            ' Check to make sure code line is necessary (commented lines are not necessary)
            If CodeLine.Length > 2 Then
                If CodeLine.Length < 4 Or CodeLine.Substring(0, 2) = "//" Then
                    Continue Do
                End If
            Else
                Continue Do
            End If

            ' If code line is valid, add its CodeLine structure to the CodeLineList array
            CodeLineList.Add(New CodeLine(CodeLine))

        Loop

        ' Dump variables used in code reading (remember, code reading only occurs at the very beginning of the program)
        RulesPath = Nothing
        CodeFile = Nothing
        CodeLine = Nothing

        Dim SkipMainLoop As Boolean = False

        'Main coordination block
        While True

            SkipMainLoop = False

            ' Update pre-code global variables
            Call UpdateGlobals()

            ' Parse code (note that this has to store global variables!)
            For Each win As TargetWindowPlain In WinArray
                For Each ce As CodeLine In CodeLineList

                    Try
                        If ParseAndAct(win, ce) = 2 Then
                            Call UpdateGlobals()
                        End If
                    Catch ex As Exception
                        SkipMainLoop = True
                        Exit For
                    End Try

                    If SkipMainLoop Then
                        Exit For
                    End If

                Next

                If SkipMainLoop Then
                    Exit For
                End If

            Next
        End While
    End Sub

    ' Update/Reset global values (Called each time all windows are parsed through entire file)
    Public Sub UpdateGlobals()

        ' Dump all unneeded global variables in memory
        MyBase.Dispose()
        ChildWinArray = Nothing
        CodeLine = Nothing
        WinArray = Nothing
        CodeIsIfLine = Nothing
        CodeListOfWindows = Nothing
        IEListFiltered = New List(Of InternetExplorer)

        ' Get list of windows
        '   If the program is told to only scan internet explorer windows, it won't bother using EnumWindows
        '   This makes the program considerably faster to react, but can be turned off by the user if needed
        If Not My.Settings.OnlyScanIEPages Then
            EnumWindows(AddressOf EnumResults, 0)
        End If

        ' Get list of all InternetExplorer instances
        '   If the program is told to only scan internet explorer windows, add them to the normal full array of windows here in TargetWindowPlain format
        For Each win As InternetExplorer In New ShellWindows
            If (win.Name = "Windows Internet Explorer") Then
                IEListFiltered.Add(win)
                If My.Settings.OnlyScanIEPages Then
                    AllWindowsList.Add(New TargetWindowPlain(win.HWND))
                End If
            End If
        Next

        WinArray = AllWindowsList.ToArray

        ' Misc
        CodeListOfWindows = New TargetWindowPlain(10)() {}
        CodeIsIfLine = False

        Return
    End Sub

    ' Create a new instance of the TargetWindowPlain class for the given window and add it to the array of all windows
    Private Function EnumResults(ByVal hWnd As Integer, ByVal lParam As Integer)

        ' (MAKE A SETTING FOR THIS!) Check to make sure window is visible (if not, don't add to the array)
        If IsWindowVisible(hWnd) Then

            Dim TrgWin As New TargetWindowPlain(hWnd)
            AllWindowsList.Capacity = AllWindowsList.Capacity + 1
            AllWindowsList.Add(TrgWin)

        End If

        ' This return statement is mandatory and CANNOT be changed without causing a bug
        Return hWnd
    End Function

    ' This function parses a window with a given handle (for both ifs and functions)
    '   This one-window-at-a-time parsing allows "emergency parsing" of recently changed windows (to cut down on response time)
    Public Function ParseAndAct(ByVal win As TargetWindowPlain, ByVal ce As CodeLine) As Boolean

        Dim CodeLine As String = ce.Line
        CodeIsIfLine = ce.IsIfLine

        ' Check to make sure line isn't a function break - if it is, reset all necessary variables
        If (CodeLine = "@#$^$#@") Then
            CodeIsIfLine = False
            Return 2                        ' Break encountered = 2
        End If

        ' If the line is an IF line, conduct an if statement parse on each IfStatement in the array
        If ce.IsIfLine Then

            ' Parse each if
            For Each IfCntr As IfStatement In ce.Actions

                ' Title (Type)
                '   NOTE: This title-based system cannot redirect tabbed windows using URL markers (uses tab titles instead)
                If IfCntr.Type = "title" Then

                    ' Get # of times needle occurs in title
                    Dim Title As String = win.Title
                    Dim InCountRange As Boolean = GetCount(IfCntr.Needle, Title, IfCntr.MinKeyCount, IfCntr.MaxKeyCount)

                    ' If title contains Internet Explorer marker, set Win.IsIE to TRUE
                    If Title.Contains("Internet Explorer") Then
                        win.IsIE = True
                    End If

                    ' Booleans (if the window satifies the condition, its truth count is increased by 1)
                    '   Note: Since URL's can be retrieved from these windows, if statements are used as UIDs (which are then re-parsed against all IE windows at function run to see if they match up)
                    If IfCntr.BooleanMarker = "+" And InCountRange Then
                        win.TruthCount = win.TruthCount + 1
                        win.IfStatement = IfCntr
                    End If
                    If IfCntr.BooleanMarker = "-" And Title.Contains(IfCntr.Needle) Then
                        win.TruthCount = win.TruthCount + 1
                        win.IfStatement = IfCntr
                    End If
                    If IfCntr.BooleanMarker = "=" And Title = IfCntr.Needle Then
                        win.TruthCount = win.TruthCount + 1
                        win.IfStatement = IfCntr
                    End If
                    If IfCntr.BooleanMarker = "~" And Title <> IfCntr.Needle Then
                        win.TruthCount = win.TruthCount + 1
                        win.IfStatement = IfCntr
                    End If
                End If

                ' URL type
                If IfCntr.Type = "urlid" Then

                    ' Search through AllWindowsList for the proper window(s) to interact with
                    For Each ie As InternetExplorer In IEListFiltered

                        ' Check to make sure IE is valid - if not, skip the window
                        Try
                            Dim Str As Integer = ie.ReadyState
                        Catch ex As Exception
                            Continue For
                        End Try

                        ' Booleans
                        If IfCntr.BooleanMarker = "+" And ie.LocationURL.Contains(IfCntr.Needle) Then
                            If win.HWND = ie.HWND Then
                                win.TruthCount = win.TruthCount + 1
                                win.URLMarker = ie.LocationURL
                                win.IsIE = True
                            End If
                        End If
                        If IfCntr.BooleanMarker = "-" And (ie.LocationURL.Contains(IfCntr.Needle) = False) Then
                            If win.HWND = ie.HWND Then
                                win.TruthCount = win.TruthCount + 1
                                win.URLMarker = ie.LocationURL
                                win.IsIE = True
                            End If
                        End If
                        If IfCntr.BooleanMarker = "=" And ie.LocationURL = IfCntr.Needle Then
                            If win.HWND = ie.HWND Then
                                win.TruthCount = win.TruthCount + 1
                                win.URLMarker = ie.LocationURL
                                win.IsIE = True
                            End If
                        End If
                        If IfCntr.BooleanMarker = "~" And ie.LocationURL <> IfCntr.Needle Then
                            If win.HWND = ie.HWND Then
                                win.TruthCount = win.TruthCount + 1
                                win.URLMarker = ie.LocationURL
                                win.IsIE = True
                            End If
                        End If

                    Next
                End If

                ' HTML type
                If IfCntr.Type = "htmlc" Then

                    Dim DocStr As String
                    For Each ie As InternetExplorer In IEListFiltered

                        ' Check to make sure IE is valid - if not, skip the window
                        Try
                            Dim Str As Integer = ie.ReadyState
                        Catch ex As Exception
                            Continue For
                        End Try

                        ' If IE window is busy, let it load and investigate it later (to prevent null reference exceptions)
                        If Not ie.Busy Then
                            DocStr = ie.Document.Body.InnerHTML
                        Else
                            Continue For
                        End If

                        ' If HTML text is null, come back to it later (to prevent null reference exception)
                        If String.IsNullOrWhiteSpace(DocStr) Then
                            Continue For
                        End If

                        ' Get/check count checks
                        Dim InCountRange As Boolean = GetCount(IfCntr.Needle, DocStr, IfCntr.MinKeyCount, IfCntr.MaxKeyCount)

                        ' Booleans
                        Try
                            If IfCntr.BooleanMarker = "+" And InCountRange Then
                                If win.HWND = ie.HWND Then
                                    win.TruthCount = win.TruthCount + 1
                                    win.URLMarker = ie.LocationURL
                                    win.IsIE = True
                                End If
                            End If
                            If IfCntr.BooleanMarker = "-" And DocStr.Contains(IfCntr.Needle) = False Then
                                If win.HWND = ie.HWND Then
                                    win.TruthCount = win.TruthCount + 1
                                    win.URLMarker = ie.LocationURL
                                    win.IsIE = True
                                End If
                            End If
                            If IfCntr.BooleanMarker = "=" And DocStr = IfCntr.Needle Then
                                If win.HWND = ie.HWND Then
                                    win.TruthCount = win.TruthCount + 1
                                    win.URLMarker = ie.LocationURL
                                    win.IsIE = True
                                End If
                            End If
                            If IfCntr.BooleanMarker = "~" And DocStr <> IfCntr.Needle Then
                                If win.HWND = ie.HWND Then
                                    win.TruthCount = win.TruthCount + 1
                                    win.URLMarker = ie.LocationURL
                                    win.IsIE = True
                                End If
                            End If
                        Catch ex As Exception
                        End Try
                    Next
                End If

            Next
        End If

        '    'MsgBox("Done " & CodeIfArray.Length)

        ''If all IF statements are complete, handle conjunctions and update the cumulative window list (CodeListOfWindows)
        If Not CodeIsIfLine Then


            ' Boolean handling
            Dim Conjunction As String = ce.BoolOp

            Dim AddList As New List(Of TargetWindowPlain)
            If Conjunction = ":NOR" Then
                If win.TruthCount > 0 Then
                    win.TruthCount = 0
                    AddList.Add(win)
                End If
            End If

            If Conjunction = ":XOR" Then
                If win.TruthCount = 1 Then
                    win.TruthCount = 0
                    AddList.Add(win)
                End If
            End If

            If Conjunction = ":AND" Then
                If win.TruthCount = ce.Actions.Length Then
                    win.TruthCount = 0
                    AddList.Add(win)
                End If
            End If

            CodeListOfWindows.SetValue(AddList.ToArray, ce.LayerCount)

        End If

        ' Function parsing
        If Not CodeIsIfLine Then
            For Each func As FunctionPlain In ce.Actions

                Dim FuncFirstArg As String = FunctionPlain.GetParam(func.Statement, 1)

                ' Entire window list commands (!Msg, !Snd)
                '   NOTE: THESE NEED WORK (MAKE SURE THEY ONLY ACTIVATE ONCE EACH CODE READ, NOT WITH EACH WINDOW!!!)
                If func.Type = "!msg" Then
                    MsgBox(FuncFirstArg)
                End If
                If func.Type = "!snd" Then
                    My.Computer.Audio.Play(FuncFirstArg)
                End If

                '' Parse previously targeted windows
                ' BELOW IS BROKEN
                'For Each win2 As TargetWindowPlain In CodeListOfWindows.GetValue(LayerCount - 1)

                '    ' Per-window commands (!Dir, !ClP, !ClW, !DlP, !DlR, !Opn, !Cls) --> Delete process / Delete registry entries / Unlock (permissions) / Lock (permissions)
                '    If func.Type = "!clw" Then
                '        DestroyWindow(win.HWND)
                '    End If

                If func.Type = "!dir" Then
                    If win.IsIE Then
                        For Each ie As InternetExplorer In IEListFiltered

                            If (ie.ReadyState > 1 And Not (String.IsNullOrWhiteSpace(ie.LocationURL))) Then

                                ' Define booleans
                                Dim TitleCheck, IECheck As Boolean

                                ' Check title based search
                                If (win.IfStatement IsNot Nothing) Then
                                    TitleCheck = (IfStatement.CheckTitleTruth(ie.LocationName, win.IfStatement) = 1)
                                Else
                                    TitleCheck = False
                                End If

                                ' Check IE property search (URL markers)
                                IECheck = ie.LocationURL.Equals(win.URLMarker)
                                If (TitleCheck Or IECheck) And win.IsIE Then
                                    ie.Navigate(FuncFirstArg)
                                End If

                            End If
                        Next
                    End If
                End If

            Next

        End If

        ' If operation was successful, return 1
        Return 1
    End Function




    ' Scan function - scans individual windows for their compliance with the terms listed in the rules file

    ' List all child windows of a window
    'Public Sub ListChildWindows(ByVal Hwnd As Integer)
    '    ChildWinArray = Nothing
    '    Dim NilInt As Integer = EnumChildWindows(Hwnd, AddressOf EnumChildResults, 0)

    '    ChildWinArray.SetValue(Hwnd,
    '    Return
    'End Sub

    Public Function EnumChildResults(ByVal Hwnd As Integer, ByVal lParam As Integer)
        Dim ECR_ClassStr As New String(Nothing, 50)
        Dim ECR_ClassLen As Integer = GetClass(Hwnd, ECR_ClassStr, 49)

        ECR_ClassStr = ECR_ClassStr.Substring(0, ECR_ClassLen)


        'Add window to ChildWinArray if its a valid tab
        Try
            Dim Catch2 As New Object
            If ECR_ClassStr = "Internet Explorer_Server" Then

                'Debug - Get Mr. Smith to give more efficient code?
                For Each ie As InternetExplorer In New ShellWindows()

                    'Dim C As Control = Control.FromChildHandle(Hwnd)
                    'Console.WriteLine(C.Parent.Handle)

                    ' Check all controls
                    Dim IECtrl As New Control
                    'IECtrl = Control.FromChildHandle

                Next
                MsgBox("Done!")
            End If
        Catch ex As Exception
        End Try

        ' This return statement is mandatory and CANNOT be changed without causing a bug
        Return Hwnd
    End Function

    ' Get number of strings in another string
    '   NOTE: Specifying a maximum count can slow this function down, because instead of counting the first [Minimum] occurrences, it must count each one
    Private Function GetCount(ByVal Needle As String, ByVal Haystack As String, ByVal MinCnt As String, ByVal MaxCnt As String) As String

        ' Exit out if haystack or needle is nill
        'If String.IsNullOrEmpty(Needle) Or String.IsNullOrEmpty(Haystack) Then
        '    Return 0
        'End If

        ' Main checking part
        Dim StrA As String = Haystack.ToLower.Replace(Needle.ToLower, "")
        Dim Count As Integer = Math.Round((Haystack.Length - StrA.Length) / Needle.Length)
        Dim MeetsCounts As Boolean = False
        Dim WordCnt As Integer = -1

        ' Check minimums
        If MinCnt.Contains("%") Then
            WordCnt = Math.Round((StrA.Length - StrA.Replace(" ", "").Length) / Needle.Length)
            MeetsCounts = ((Count / WordCnt) * 100) >= CInt(MinCnt.Replace("%", ""))
        ElseIf MinCnt > 0 Then
            MeetsCounts = (Count >= MinCnt)
        Else
            MeetsCounts = (Count > 0)
        End If

        ' Check maximums
        If MaxCnt.Contains("%") Then
            If WordCnt = -1 Then
                WordCnt = Math.Round((StrA.Length - StrA.Replace(" ", "").Length) / Needle.Length)
                MeetsCounts = ((Count / WordCnt) * 100) <= CInt(MaxCnt.Replace("%", ""))
            End If
        ElseIf MaxCnt > 0 Then
            MeetsCounts = (Count <= MinCnt)
        End If
        Return MeetsCounts

    End Function
End Class

' ----------------------------------------------------------------------- KEEP ALL BELOW THIS LINE -----------------------------------------------------------------------
Public Class SystemFunctions

    ' Trim a string from a certain character
    Public Shared Function TrimToChar(ByVal Haystack As String, ByVal Needle As String, ByVal RightSide As Boolean, ByVal IncludeTrimmed As Boolean) As String
        Dim Pos As Integer = Haystack.IndexOf(Needle)
        Dim Str As String

        ' Check to make sure object exists in string
        If Not Haystack.Contains(Needle) Then
            Return Haystack
        End If

        ' True = trim from right (c, abcde --> de), False = trim from left (c, abcde --> ab)
        If RightSide Then
            Str = Haystack.Substring(Pos + (1 - IncludeTrimmed))
        Else
            Str = Haystack.Substring(0, Pos + IncludeTrimmed)
        End If

        Return Str
    End Function

    ' Replace function with all essential exception handling
    Shared Function SafeReplace(ByVal Haystack As String, ByVal Needle As String, ByVal Result As String)

        ' If Needle is longer than Haystack
        If (Haystack.Length <= Needle.Length) Then
            If (Haystack = Needle) Then
                Return ""
            End If
            Return Haystack
        End If

        ' If Needle doesn't exist in Haystack
        If (Haystack.Contains(Needle) = False) Then
            Return Haystack
        End If

        ' If Result is nothing, remove all occurrences of Needle from Haystack
        If (Result = Nothing) Then
            Result = ""
        End If

        Return Haystack.Replace(Needle, Result)
    End Function
End Class

Public Class FunctionPlain
    ' Property guide
    ' x.Type - returns type of function
    ' x.Param(X) - returns parameter (X) of function ("" if null)

    Dim FuncStr As String
    Dim TypeStr As String
    Public Sub New(ByVal FuncStrIn As String)
        FuncStr = FuncStrIn
        TypeStr = SystemFunctions.TrimToChar(FuncStr, " ", False, False).ToLower
    End Sub

    ' Get type
    ReadOnly Property Type As String
        Get
            Return TypeStr
        End Get
    End Property

    ' Get param
    Public Shared Function GetParam(ByVal FuncStrA As String, ByVal Number As Integer)
        For i = 1 To Number
            FuncStrA = SystemFunctions.TrimToChar(FuncStrA, " ", True, False).ToLower()
        Next
        FuncStrA = SystemFunctions.TrimToChar(FuncStrA, ",", False, False).ToLower()
        Return FuncStrA
    End Function

    ' Get whole string
    ReadOnly Property Statement
        Get
            Return FuncStr
        End Get
    End Property
End Class

' Finished
Public Class IfStatement
    ' Property guide
    ' x.BooleanMarker - returns Boolean Marker
    ' x.Type - returns Type of object being checked (ex "Title")
    ' x.Needle - returns Needle
    ' x.[Min/Max]KeyCount - returns minimum/maximum key counts

    Dim StarterIf, BooleanMkr, CondType, NeedleStr As String
    Public Sub New(ByVal IfStr As String)
        StarterIf = IfStr
        BooleanMkr = StarterIf.First
        CondType = SystemFunctions.TrimToChar(StarterIf, " ", False, False).Substring(1).ToLower
        NeedleStr = SystemFunctions.TrimToChar(SystemFunctions.TrimToChar(StarterIf, " ", True, False), ",", False, False)
    End Sub

    ' Get boolean marker (+, -, =, ~)
    ReadOnly Property BooleanMarker As String
        Get
            Return BooleanMkr
        End Get
    End Property

    ' Get function being checked
    ReadOnly Property Type As String
        Get
            Return CondType
        End Get
    End Property

    ' Get needle
    ReadOnly Property Needle As String
        Get
            Return NeedleStr
        End Get
    End Property

    ' Get minimum key count
    ReadOnly Property MinKeyCount As String
        Get
            If (StarterIf.Contains(",") = False) Then
                Return 0
            End If
            Return SystemFunctions.TrimToChar(SystemFunctions.TrimToChar(StarterIf, ",", True, False), ",", False, False)
        End Get
    End Property

    ' Get maximum key count
    ReadOnly Property MaxKeyCount As String
        Get
            If (StarterIf.Contains(",") = False) Then
                Return 0
            End If
            Return SystemFunctions.TrimToChar(SystemFunctions.TrimToChar(StarterIf, ",", True, False), ",", True, False)
        End Get
    End Property

    ' Get raw if string
    ReadOnly Property Statement As String
        Get
            Return StarterIf
        End Get
    End Property

    ' Check truth of title (returns -1 if IfStatement isn't of type Title - meaning it won't work no matter what)
    Shared Function CheckTitleTruth(ByVal PageTitle As String, ByVal IfCmdA As IfStatement)

        ' Check for non-title if
        If (IfCmdA.Type <> "title") Then
            Return -1
        End If

        ' Parse title ifs - return whether or not a page title fulfills its boolean command
        If (IfCmdA.BooleanMarker = "+") Then
            Return PageTitle.Contains(IfCmdA.Needle)
        End If
        If (IfCmdA.BooleanMarker = "-") Then
            Return (PageTitle.Contains(IfCmdA.Needle) = False)
        End If
        If (IfCmdA.BooleanMarker = "=") Then
            Return PageTitle = IfCmdA.Needle
        End If
        If (IfCmdA.BooleanMarker = "~") Then
            Return PageTitle <> IfCmdA.Needle
        End If

        ' If nothing has worked so far, return false value
        Return -1
    End Function
End Class

' WIP
Public Class TargetWindowPlain
    Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Integer, ByVal lpString As String, ByVal nMaxCount As Integer) As Integer
    Declare Function GetWindowProcess Lib "user32" Alias "GetWindowThreadProcessId" (ByRef hWnd As IntPtr, <Out()> Optional ByVal lpdwProcessId As Int32 = 0)

    ' Property guide
    ' x.HWND    - returns Window Handle (HWND)
    ' x.Title   - returns Title
    ' x.Process - returns Process
    ' x.Truth   - returns truth count

    ' Create Target Window data
    Dim TitleStr As New String(Nothing, 300)
    Dim TargetHWND, PrcNum As Integer
    Dim URLMarkerStr As String
    Dim IsIEBool As Boolean = False
    Dim IfCmd As IfStatement

    Public Sub New(ByVal NewHWND As Integer)

        TargetHWND = NewHWND

        Dim StrLen As Integer = GetWindowText(TargetHWND, TitleStr, 300)
        TitleStr = TitleStr.Substring(0, StrLen)
        URLMarkerStr = "none"

    End Sub

    ' Get HWND
    ReadOnly Property HWND As Integer
        Get
            Return TargetHWND
        End Get
    End Property

    ' Set/Get Truth Count
    Dim Truth As Integer
    Property TruthCount As Integer
        Set(ByVal Cnt As Integer)
            Truth = Cnt
        End Set
        Get
            Return Truth
        End Get
    End Property

    ' Get title
    ReadOnly Property Title As String
        Get
            Return TitleStr
        End Get
    End Property

    ' HTML unique characteristics - specify exact URL (in order to act on proper window)
    Property URLMarker As String
        Get
            Return URLMarkerStr
        End Get
        Set(ByVal value As String)
            URLMarkerStr = value
        End Set
    End Property

    ' Returns whether window is instance of Internet Explorer
    Property IsIE As Boolean
        Get
            Return IsIEBool
        End Get
        Set(ByVal IsInternetExplorer As Boolean)
            IsIEBool = IsInternetExplorer
        End Set
    End Property

    ' If statement that triggered the window (used for title + IE interaction)
    Property IfStatement As IfStatement
        Get
            Return IfCmd
        End Get
        Set(ByVal value As IfStatement)
            IfCmd = value
        End Set
    End Property
End Class

' Allows for each line of code to be stored in a line - this way, code files need be processed only at startup of Nighthawk
'   This makes the program more efficient, but doesn't allow the user to modify the code while Nighthawk is running (they will have to re-load Nighthawk)
'   NOTE: If the statement is a block separator (@#$^$#@), only the first code line (x.Line) will return anything
'       This is because break lines need to be stored, but the other functions are incompatible with them
Public Class CodeLine

    ' Property Guide
    '   x.Line          - Gets raw code line
    '   x.LineNoLMs     - Gets code line without Line Markers (-->)
    '   x.Actions       - Gets [If/function] array (depending on which one it is)
    '   x.LayerCount    - Gets layer count of the code line
    '   x.IsIfLine      - Gets whether or not the code line is an if statement
    '   x.BoolOp        - Gets boolean operator (:NOR, :XOR, :AND)

    Dim CLine, CLineNoLMs, CBoolOp As String
    Dim CIsIf As Boolean
    Dim CActions As Array
    Dim CLayerCnt As Integer
    Public Sub New(ByVal CodeA As String)

        ' Get simple values
        CLine = CodeA
        If CLine <> "@#$^$#@" Then
            CLineNoLMs = CLine.Replace("-->", "")
            CIsIf = CLineNoLMs.Substring(0, 1).Equals(":")
            CLayerCnt = Math.Floor(CLine.LastIndexOf("-->") / 3) + 1
            CBoolOp = CLineNoLMs.Substring(0, 4)

            ' Get action list
            If CIsIf Then
                CActions = SeriesToIf(CLineNoLMs)
            Else
                CActions = SeriesToFunction(CLineNoLMs)
            End If
        End If
    End Sub

    ' Property statements
    ReadOnly Property Line As String
        Get
            Return CLine
        End Get
    End Property
    ReadOnly Property LineNoLMs
        Get
            Return CLineNoLMs
        End Get
    End Property
    ReadOnly Property Actions As Array
        Get
            Return CActions
        End Get
    End Property
    ReadOnly Property LayerCount As Integer
        Get
            Return CLayerCnt
        End Get
    End Property
    ReadOnly Property IsIfLine As Boolean
        Get
            Return CIsIf
        End Get
    End Property
    ReadOnly Property BoolOp As String
        Get
            Return CBoolOp
        End Get
    End Property

    ' Take a line of if code and separate it into IfStatement objects
    ' BROKEN ATM
    Shared Function SeriesToIf(ByVal StarterIf As String) As Array

        ' Basic replaces
        StarterIf = StarterIf
        If (StarterIf = Nothing Or StarterIf = "" Or StarterIf.Length < 3) Then
            Dim Array As String()
            Array.SetValue("", 1)
            Return Array
        End If

        StarterIf = SystemFunctions.SafeReplace(StarterIf, ":NOR", "")
        StarterIf = SystemFunctions.SafeReplace(StarterIf, ":XOR", "")
        StarterIf = SystemFunctions.SafeReplace(StarterIf, ":AND", "")

        '' Assign each individual if clause to an array as an IfStatement
        Dim StrArray As String()
        StrArray = StarterIf.Split("[()]")
        Dim IfList As New List(Of IfStatement)

        Dim IfStr As String
        For i = 0 To StrArray.Length - 1

            ' For exception handling
            IfStr = SystemFunctions.SafeReplace(StrArray(i), "()]", "")

            If String.IsNullOrWhiteSpace(IfStr) Then
                Continue For
            End If

            If (IfStr.Length < 3) Then
                Continue For
            End If

            Dim First As String = IfStr.Substring(0, 1)
            If First <> "+" And First <> "-" And First <> "=" And First <> "~" Then
                Continue For
            End If

            ' IfStatement defining/adding
            Dim IfPart As New IfStatement(IfStr)
            IfList.Add(IfPart)
        Next

        Return IfList.ToArray
    End Function

    ' Take a line of if code and separate it into function objects - note that these, unlike If objects, are simply strings (they only store one parameter - the type of function)
    Shared Function SeriesToFunction(ByVal StarterFunc As String)

        ' Basic replaces
        StarterFunc = SystemFunctions.SafeReplace(StarterFunc, "-->", "")

        ' Get array of functions (as strings) and convert them into an array of FunctionPlains
        Dim StrArray As Array = StarterFunc.Split("@")
        Dim FuncList As New List(Of FunctionPlain)
        For Each Str As String In StrArray
            If Not String.IsNullOrWhiteSpace(Str) Then
                FuncList.Add(New FunctionPlain(Str))
            End If
        Next

        ' Return the array
        Return FuncList.ToArray
    End Function
End Class
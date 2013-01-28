' TODO: Fix title search with IE pages
' TODO: Add firefox capability --> Perhaps use an IE window that goes to the same sites as Firefox?
' TODO: Get basic GUI working!
' TODO: Add code validation!
' TODO: Add modification prevention (by opening streams to all protected files)

Imports System
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Net
Imports System.Text
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

    ' User modifiable variables --> Integrate into GUI!
    Dim RulesPath As String = My.Settings.EveryoneRules

    ' Global variable definitions (user-do-not-modify)
    Dim UserRank As Integer = 1 ' If there is no assigned user rank, they will be presumed as a student
    Dim NullArray As Array
    Dim CodeFile As StreamReader
    Dim CodeLineList As New List(Of CodeLine)
    Dim ChildWinArray As Array
    Dim CodeLine As String
    Dim WinArray As Array
    Dim CodeIsIfLine As Boolean
    Dim IEListFiltered As New List(Of InternetExplorer)
    Dim NHBrowser As New WebClient

    Public Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' Set process to realtime priority (faster scanning)
        Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.RealTime

        ' Program path - this MUST be correct!
        Dim ProgramPath As String = "C:/Program Files/Nighthawk by Xixo12e"

        ' Get global settings
        Dim RulesPath As String
        If True Then

            ' Get user rank
            If Dir(ProgramPath & "/Config/userranks.txt").Contains(".txt") Then

                ' Get text file
                Try
                    Dim RankStream As New StreamReader(ProgramPath & "/Config/userranks.txt")

                    While RankStream.Peek > -1

                        ' Get line
                        Dim RankLine As String = RankStream.ReadLine().ToLower

                        ' If current line doesn't contain current user name, skip it
                        If Not RankLine.Contains(Environment.UserName.ToLower & "=") Then
                            Continue While
                        End If

                        ' Get user rank
                        Dim UserRankStr As String = RankLine.Substring(RankLine.Length - 1)

                        ' Exception Handler
                        If UserRankStr <> "1" And UserRankStr <> "2" And UserRankStr <> "3" And UserRankStr <> "4" And UserRankStr <> "5" Then
                            MsgBox("The following line in the rank list is improperly formatted: [" & RankLine & "]." & vbCrLf & vbCrLf & "Make sure the config/userranks file is properly written. Nighthawk will now exit.", MsgBoxStyle.Critical)
                            Process.GetCurrentProcess().Kill()
                        End If

                        UserRank = CInt(UserRankStr)

                        Exit While

                    End While

                Catch ex As Exception

                    Dim Ready As Integer = MsgBox("An unhandled error has occurred of type [" * ex.GetType.ToString & "] at Exception Point A. Please record this error data, and then press OK to exit Nighthawk.")
                    If Ready > -1 Then
                        Process.GetCurrentProcess().Kill()
                    End If

                End Try

            Else

                MsgBox("There is no valid rank file referenced." & vbCrLf & vbCrLf & "The rank file should be at the following location: [" & ProgramPath & "/Config/userranks.txt" & "]." & vbCrLf & vbCrLf & "Nighthawk will now exit.", MsgBoxStyle.Critical)
                Process.GetCurrentProcess().Kill()

            End If

            ' Get appropriate rules path
            If UserRank = 5 Then
                RulesPath = ProgramPath & "/Config/Rules/superadmin.txt"
            ElseIf UserRank = 4 Then
                RulesPath = ProgramPath & "/Config/Rules/admin.txt"
            ElseIf UserRank = 3 Then
                RulesPath = ProgramPath & "/Config/Rules/superuser.txt"
            ElseIf UserRank = 2 Then
                RulesPath = ProgramPath & "/Config/Rules/teacher.txt"
            Else
                RulesPath = ProgramPath & "/Config/Rules/student.txt"
            End If

            ' Get stream to rules folder
            If Dir(RulesPath).Contains(".txt") Then
                CodeFile = New StreamReader(RulesPath)
            Else
                MsgBox("There is no valid rules file referenced." & vbCrLf & vbCrLf & "The rules file should be at the following location: [" & RulesPath & "]." & vbCrLf & vbCrLf & "Nighthawk will now exit.")
                Process.GetCurrentProcess().Kill()
            End If

            ' Validate the rules file - TO DO

            ' Modification prevention


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

        'Main coordination block
        While True
            Try

                Call FilterLoop()

            Catch ex As Exception

                If ex.Message = "Error HRESULT E_FAIL has been returned from a call to a COM component." Then

                    ' This error isn't critical, so just ignore it and keep things moving

                Else
                    MsgBox("An error has occurred in the main Nighthawk engine at Error Point A." & vbCrLf & vbCrLf & "Error: [" & ex.Message & "]" & vbCrLf & vbCrLf & "Nighthawk will now exit.")
                    Process.GetCurrentProcess.Close()
                End If

            End Try
        End While

    End Sub

    ' Filtering loop
    Public Sub FilterLoop()
        While True

            ' Update pre-code global variables
            Call UpdateGlobals()

            ' Parse code (note that this has to store global variables!)
            For Each win As TargetWindowPlain In WinArray

                ' Define list of per-layer booleans
                Dim TruthListLayers As New List(Of Integer)

                For j = 0 To (CodeLineList.Count - 1)

                    ' Get code line
                    Dim ce As CodeLine = CodeLineList.Item(j)

                    ' Execute code
                    Try
                        Dim LineTruth As Integer

                        ' Make sure TruthListLayers has the proper capacity


                        ' Check to make sure truth for above layer exists (if layercount isn't 0);
                        '   after all, if this truth is 0 nothing should happen and code parsing is just wasted processor power
                        If ce.LayerCount > 0 Then

                            If TruthListLayers.Item(ce.LayerCount - 1) = 1 Then

                                ' Execute Code
                                Dim IfResult As Integer = ParseAndAct(win, ce)

                                ' Handle ParseAndAct Return Codes
                                '   2 - Function break encountered, reset
                                '   3 - Window acted on and "Top-Down Single Action" mode enabled; therefore parse the next window
                                If IfResult = 3 Then
                                    TruthListLayers = New List(Of Integer)
                                    Exit For
                                End If

                                ' Set truth in TruthListLayers
                                If ce.IsIfLine Then

                                    If TruthListLayers.Count > (ce.LayerCount + 1) Then
                                        TruthListLayers.Item(ce.LayerCount) = IfResult
                                    Else
                                        TruthListLayers.Add(IfResult)
                                    End If
                                End If
                            Else
                                If TruthListLayers.Count > (ce.LayerCount + 1) Then
                                    TruthListLayers.Item(ce.LayerCount) = 0
                                Else
                                    TruthListLayers.Add(0)
                                End If
                            End If

                        Else

                            ' Execute code / Find truth of the item
                            LineTruth = ParseAndAct(win, ce)

                            ' If there is no spot on the list, add the item; otherwise, simply set the existing item to the current one
                            If TruthListLayers.Count > (ce.LayerCount + 1) Then
                                TruthListLayers.Item(ce.LayerCount) = LineTruth
                            Else
                                TruthListLayers.Add(LineTruth)
                            End If

                        End If

                        ' If a break line is returned, then update global variables
                        If LineTruth = 2 Then
                            Call UpdateGlobals()
                        End If
                    Catch ex As Exception
                        Exit For
                    End Try

                Next

            Next
        End While
    End Sub

    ' Update/Reset global values (Called each time all windows are parsed through entire file)
    Public Sub UpdateGlobals()

        ' Dump all unneeded global variables in memory
        MyBase.Dispose()
        AllWindowsList.Clear()
        ChildWinArray = NullArray
        CodeLine = Nothing
        CodeIsIfLine = Nothing
        IEListFiltered = New List(Of InternetExplorer)

        ' Get list of windows
        '   If the program is told to only scan internet explorer windows, it won't bother using EnumWindows
        '   This makes the program considerably faster to react, but can be turned off by the user if needed
        If Not My.Settings.OnlyScanIEPages Then
            EnumWindows(AddressOf EnumResults, 0)
        Else

            ' Get list of all InternetExplorer instances
            '   If the program is told to only scan internet explorer windows, add them to the normal full array of windows here in TargetWindowPlain format
            For Each win As InternetExplorer In New ShellWindows
                If (win.Name = "Windows Internet Explorer") Then
                    Dim TWP As New TargetWindowPlain(win.HWND)
                    TWP.IsIE = True
                    TWP.URLMarker = win.LocationURL
                    TWP.InternetExplorerObject = win
                    IEListFiltered.Add(win)
                    If My.Settings.OnlyScanIEPages Then
                        AllWindowsList.Add(TWP)
                    End If
                End If
            Next
        End If

        WinArray = AllWindowsList.ToArray

        ' Misc
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
    Public Function ParseAndAct(ByVal win As TargetWindowPlain, ByVal ce As CodeLine) As Integer

        Dim CodeLine As String = ce.Line
        Dim CodeActedOn As Boolean = False
        Dim TruthOfOperation As Integer = 0
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
                    '   Note: Since URL's can't be retrieved from these windows, if statements are used as UIDs (which are then re-parsed against all IE windows at function run to see if they match up)
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

                    ' Check to make sure IE is valid - if not, skip the window
                    Dim NoGo As Boolean = False
                    Try
                        Dim Str As Integer = win.InternetExplorerObject.ReadyState
                    Catch ex As Exception
                        NoGo = True
                    End Try

                    ' Booleans
                    If Not NoGo Then
                        If IfCntr.BooleanMarker = "+" And win.InternetExplorerObject.LocationURL.Contains(IfCntr.Needle) Then
                            win.TruthCount = win.TruthCount + 1
                        End If
                        If IfCntr.BooleanMarker = "-" And (win.InternetExplorerObject.LocationURL.Contains(IfCntr.Needle) = False) Then
                            win.TruthCount = win.TruthCount + 1
                        End If
                        If IfCntr.BooleanMarker = "=" And win.InternetExplorerObject.LocationURL = IfCntr.Needle Then
                            win.TruthCount = win.TruthCount + 1
                        End If
                        If IfCntr.BooleanMarker = "~" And win.InternetExplorerObject.LocationURL <> IfCntr.Needle Then
                            win.TruthCount = win.TruthCount + 1
                        End If
                    End If

                End If

                ' HTML type
                If IfCntr.Type = "htmlc" Then

                    ' Check to make sure IE is valid - if not, skip the window
                    Try
                        Dim Str As Integer = win.InternetExplorerObject.ReadyState
                    Catch ex As Exception
                        Continue For
                    End Try

                    ' Get/check count checks
                    Dim HTMLStr As String = GetActualHTML(win.InternetExplorerObject)
                    Dim InCountRange As Boolean = GetCount(IfCntr.Needle, HTMLStr, IfCntr.MinKeyCount, IfCntr.MaxKeyCount)

                    ' Booleans
                    Try
                        If IfCntr.BooleanMarker = "+" And InCountRange Then
                            win.TruthCount = win.TruthCount + 1
                        End If
                        If IfCntr.BooleanMarker = "-" And Not win.InternetExplorerObject.Document.Body.InnerHTML.Contains(IfCntr.Needle) Then
                            win.TruthCount = win.TruthCount + 1
                        End If
                        If IfCntr.BooleanMarker = "=" And win.InternetExplorerObject.Document.Body.InnerHTML = IfCntr.Needle Then
                            win.TruthCount = win.TruthCount + 1
                        End If
                        If IfCntr.BooleanMarker = "~" And win.InternetExplorerObject.Document.Body.InnerHTML <> IfCntr.Needle Then
                            win.TruthCount = win.TruthCount + 1
                        End If
                    Catch ex As Exception
                        IEListFiltered.Remove(win.InternetExplorerObject)
                    End Try

                    ' If conjunctions have been parsed far enough to determine whether they are true/false, end the loop and move on
                    '   NOR: TruthCount > 0
                    '   XOR: TruthCount > 1
                    If (ce.BoolOp = ":NOR" And win.TruthCount > 0) Or (ce.BoolOp = ":XOR" And win.TruthCount > 1) Then
                        Exit For
                    End If

                End If
            Next
        End If

        '    'MsgBox("Done " & CodeIfArray.Length)

        ' Handle conjunctions (if the code is an if line)
        If CodeIsIfLine Then


            ' Boolean handling
            TruthOfOperation = 0

            Dim DBG As String = ce.Line

            If ce.BoolOp = ":NOR" Then
                If win.TruthCount > 0 Then
                    TruthOfOperation = 1
                End If
                win.TruthCount = 0
            End If

            If ce.BoolOp = ":XOR" Then
                If win.TruthCount = 1 Then
                    TruthOfOperation = 1
                End If
                win.TruthCount = 0
            End If

            If ce.BoolOp = ":AND" Then
                If win.TruthCount = ce.Actions.Length Then
                    TruthOfOperation = 1
                End If
                win.TruthCount = 0
            End If

        End If

        ' Function parsing
        If Not CodeIsIfLine Then

            'Try

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

                '    ' Per-window commands (!Dir, !ClP, !ClW, !DlP, !DlR, !Opn, !Cls) --> Delete process / Delete registry entries / Unlock (permissions) / Lock (permissions)
                '    If func.Type = "!clw" Then
                '        DestroyWindow(win.HWND)
                '       CodeActedOn = True
                '    End If

                ' Record function
                If func.Type = "!rec" Then

                    ' Get file

                    ' Write data
                    Dim WriteStr As String = "USER: " & Environment.UserName & " | TIME: " & My.Computer.Clock.LocalTime & " | URL: " & win.URLMarker & " | VIOLATION: [If statement here]"
                    Console.WriteLine(WriteStr)

                    ' Close file

                End If

                ' Redirect window function (!Dir)
                If func.Type = "!dir" Then
                    If win.IsIE Then

                        If (win.InternetExplorerObject.ReadyState > 0 And Not (String.IsNullOrWhiteSpace(win.InternetExplorerObject.LocationURL))) Then

                            ' Attempt to navigate to the blockpage
                            win.InternetExplorerObject.Navigate(FuncFirstArg)

                            ' If navigation attempt has failed, close the window
                            If win.InternetExplorerObject.ReadyState > 3 And win.InternetExplorerObject.LocationURL <> FuncFirstArg Then
                                win.InternetExplorerObject.Quit()
                            End If

                            ' These functions help keep the overall efficiency up by removing/helping to remove already acted on windows
                            CodeActedOn = True

                        End If

                    End If

                End If


                ' Returns 3 if "top-down single action mode" is enabled
                '   This prevents the system from parsing any future commands for the window
                ' Also, remove the window from AllWindowsList / WinArray / IEListFiltered
                If My.Settings.TopDownSingleAction And CodeActedOn Then
                    Array.Clear(WinArray, Array.IndexOf(WinArray, win), 1)
                    AllWindowsList.Remove(win)
                    Return 3
                End If

            Next

            'Catch ex As Exception

            'End Try

        End If

        ' Set TruthCount to 0
        win.TruthCount = 0

        ' If operation was successful, return the truth of the operation (0 = false, 1 = true)
        Return TruthOfOperation
    End Function

    ' Scan function - scans individual windows for their compliance with the terms listed in the rules file

    ' List all child windows of a window
    'Public Sub ListChildWindows(ByVal Hwnd As Integer)
    '    ChildWinArray = Nothing
    '    Dim NilInt As Integer = EnumChildWindows(Hwnd, AddressOf EnumChildResults, 0)

    '    ChildWinArray.SetValue(Hwnd,
    '    Return
    'End Sub

    'Public Function EnumChildResults(ByVal Hwnd As Integer, ByVal lParam As Integer)
    '    Dim ECR_ClassStr As New String(Nothing, 50)
    '    Dim ECR_ClassLen As Integer = GetClass(Hwnd, ECR_ClassStr, 49)

    '    ECR_ClassStr = ECR_ClassStr.Substring(0, ECR_ClassLen)


    '    'Add window to ChildWinArray if its a valid tab
    '    If ECR_ClassStr = "Internet Explorer_Server" Then

    '        'Debug - Get Mr. Smith to give more efficient code?
    '        For Each ie As InternetExplorer In New ShellWindows()

    '            'Dim C As Control = Control.FromChildHandle(Hwnd)
    '            'Console.WriteLine(C.Parent.Handle)

    '            ' Check all controls
    '            Dim IECtrl As New Control
    '            'IECtrl = Control.FromChildHandle

    '        Next
    '        MsgBox("Done!")
    '    End If

    '    ' This return statement is mandatory and CANNOT be changed without causing a bug
    '    Return Hwnd
    'End Function

    ' Get number of strings in another string
    '   NOTE: Specifying a maximum count can slow this function down, because instead of counting the first [Minimum] occurrences, it must count each one
    Private Function GetCount(ByVal Needle As String, ByVal Haystack As String, ByVal MinCnt As String, ByVal MaxCnt As String) As String

        ' Exit out if haystack or needle is nill
        If String.IsNullOrEmpty(Needle) Or String.IsNullOrEmpty(Haystack) Then
            Return False
        End If

        ' Main checking part
        Dim StrA As String = Haystack.ToLower.Replace(Needle.ToLower, "")
        Dim Count As Integer = Math.Round((Haystack.Length - StrA.Length) / Needle.Length)
        Dim MeetsCounts As Boolean = True
        Dim MeetsMinCount As Boolean = True
        Dim WordCnt As Integer = -1

        ' Check minimums
        If MinCnt.Contains("%") Then
            WordCnt = Math.Round((StrA.Length - StrA.Replace(" ", "").Length) / Needle.Length)
            MeetsCounts = ((Count / WordCnt) * 100) >= CInt(MinCnt.Replace("%", ""))
        ElseIf CInt(MinCnt) > 0 Then
            MeetsCounts = (Count >= MinCnt)
        Else
            MeetsCounts = (Count > 0)
        End If
        MeetsMinCount = MeetsCounts

        ' Check maximums
        If MaxCnt.Contains("%") Then
            If WordCnt = -1 Then
                WordCnt = Math.Round((StrA.Length - StrA.Replace(" ", "").Length) / Needle.Length)
            End If
            MeetsCounts = ((Count / WordCnt) * 100) <= CInt(MaxCnt.Replace("%", ""))

        ElseIf CInt(MaxCnt) > 0 Then
            MeetsCounts = (Count <= MaxCnt)
        End If
        Return (MeetsCounts And MeetsMinCount)

    End Function

    ' Get accurate HTML of page
    Function GetActualHTML(ByVal ie As InternetExplorer)

        ' Method 1: Access the HTML of the page directly via DOM
        Dim DirectHTML As String = ie.Document.Body.OuterHTML

        ' Method 2: If the page contains a source-based reference ( "<FRAME src=[url]...>, use a direct HTML request on [url] )
        '   Note that this temporarily blocks the user from browsing (CHECK)
        '   This can help prevent users visiting harder-to-detect proxy sites
        If DirectHTML.ToLower.Contains("<frame") And My.Settings.ScanRefs Then

            ' Get target URL
            Dim URLStr As String
            URLStr = DirectHTML.Substring(DirectHTML.IndexOf("src=" & Chr(34)) + 5)
            URLStr = DirectHTML.Substring(0, DirectHTML.IndexOf(Chr(34)))

            ' Use NHBrowser to request HTML data of the new target URL
            Dim NHBuffer As Byte() = NHBrowser.DownloadData(URLStr)
            DirectHTML = Encoding.ASCII.GetString(NHBuffer)

        End If

        Return DirectHTML

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
    Dim IEBrowser As InternetExplorer

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

    ' InternetExplorer object
    Property InternetExplorerObject
        Get
            Return IEBrowser
        End Get
        Set(IEBrowserA)
            IEBrowser = IEBrowserA
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
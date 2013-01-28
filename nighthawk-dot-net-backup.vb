' NOTE: Windows will only have one function executed upon them per scan

Imports System
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Net
Imports SHDocVw

Public Delegate Function CallBack(ByVal hwnd As Integer, ByVal lParam As Integer) As Boolean

Public Class Form1
    ' List of target windows (AllWindowsList)
    Dim AllWindowsList As New List(Of TargetWindowPlain)
    Dim IETabList As New List(Of Integer)

    ' DLL Functions used
    Declare Function GetWindow Lib "user32" (ByVal hWnd As Integer, ByVal uCmd As Integer) As Integer
    'Declare Function GetForegroundWindow Lib "user32" () As Long
    Declare Function EnumWindows Lib "user32" (ByVal CallBackPtr As CallBack, ByVal lParam As Integer) As Boolean
    Declare Function EnumChildWindows Lib "user32" (ByVal Hwnd As Integer, ByVal CallBackPtr As CallBack, ByVal lParam As Integer) As Boolean

    Declare Ansi Function SendMessageTimeoutA Lib "user32" (ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer, ByVal fuFlags As Integer, ByVal uTimeout As Integer, ByRef lpdwResult As Integer) As Integer
    Declare Function RegisterWindowMessageA Lib "user32" (ByVal lpString As String) As Integer
    Declare Function ObjectFromLresult Lib "oleacc" (ByVal lResult As Integer, ByVal riid As System.Guid, ByVal wParam As Integer, ByRef ppvObject As HtmlDocument) As Integer
    Declare Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As Integer, ByVal dwObjectID As Int32, ByVal riid As Guid, <MarshalAs(UnmanagedType.IUnknown)> ByRef ppvObject As Object) As String

    Declare Function IsWindowVisible Lib "user32" (ByVal Hwnd As Integer) As Boolean
    Declare Function GetClass Lib "user32" Alias "RealGetWindowClass" (ByVal hWnd As IntPtr, ByVal pszType As String, ByVal cchType As IntPtr) As UInteger
    Declare Function DestroyWindow Lib "user32" (ByVal Hwnd As Integer)

    Dim ChildWinArray As Array
    Public Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' Get Nighthawk rules code (raw)
        Dim UserRank As Integer = 1 ' --> UPDATE!
        Dim RulesPath As String = My.Settings.EveryoneRules
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

        ' Nighthawk rules code variables
        Dim CodeLine As String
        Dim CodeIfArray, CodeFuncArray As Array
        Dim CodeIsIfLine As Boolean
        Dim DBGESC As Integer = 0
        Dim LayerCount As Integer = 0
        Dim TitleConditions As New List(Of IfStatement)
        Dim IEListFiltered As New List(Of InternetExplorer)

        While True
            CodeIsIfLine = False

            ' Update list of all windows
            ListWindows()
            Dim WinArray As Array = AllWindowsList.ToArray

            ' Reset all rules code variables
            Dim CodeFile As StreamReader = New StreamReader(RulesPath)                              ' Actual rules code file
            Dim CodeListOfWindows()() As TargetWindowPlain = New TargetWindowPlain(10)() {}         ' This one keeps track of the recorded windows in each level

            ' Parse each line of code
            While CodeFile.EndOfStream = False

                CodeLine = CodeFile.ReadLine()

                ' Organize code and execute non-if/function tasks
                If True Then

                    ' Check to make sure line isn't a comment/space
                    If CodeLine.Length > 2 Then
                        If CodeLine.Length < 4 Or CodeLine.Substring(0, 2) = "//" Then
                            Continue While
                        End If
                    Else
                        Continue While
                    End If

                    ' ' Check to make sure line isn't a function break - if so, reset all necessary variables
                    If (CodeLine = "@#$^$#@") Then
                        Continue While
                    End If

                End If

                ' Determine whether line is If or Function, and set up the relevant arrays
                Dim CodeLineNoLMs As String
                CodeLineNoLMs = CodeLine.Replace("-->", "")

                If CodeLineNoLMs.Length > 2 Then

                    If CodeLineNoLMs.Substring(0, 1) = ":" Then
                        CodeIfArray = SystemFunctions.SeriesToIf(CodeLineNoLMs)
                        CodeFuncArray = Nothing
                        CodeIsIfLine = True
                    Else
                        CodeIfArray = Nothing
                        CodeFuncArray = SystemFunctions.SeriesToFunction(CodeLineNoLMs)
                        CodeIsIfLine = False
                    End If



                Else
                    ' If the line in question is 2 characters or shorter, skip it
                    Continue While
                End If

                ' If the line is an IF line, conduct an if statement parse on each IfStatement in the array
                If CodeIsIfLine Then

                    ' Define array of currently found windows
                    CodeIfArray = SystemFunctions.SeriesToIf(CodeLineNoLMs)

                    ' Get array of previously validated windows (if there are any)
                    LayerCount = Math.Floor(CodeLine.LastIndexOf("-->") / 3) + 1

                    Dim TargetArrayLine As Array = Nothing
                    Dim Int As Integer = CodeIfArray.Length

                    ' Get list of all InternetExplorer instances
                    For Each win As InternetExplorer In New ShellWindows
                        If (win.Name = "Windows Internet Explorer") Then
                            IEListFiltered.Add(win)
                        End If
                    Next

                    ' Parse each if
                    DBGESC = 0
                    Dim DBGINT As Integer = 0
                    For Each IfCntr As IfStatement In CodeIfArray
                        DBGINT += 1


                        ' Title (Type)
                        '   NOTE: This title-based system cannot redirect tabbed windows using URL markers (uses tab titles instead)
                        If IfCntr.Type = "title" Then

                            'Check each window in window list
                            For Each win As TargetWindowPlain In WinArray

                                ' Get # of times needle occurs in title
                                Dim Title As String = win.Title
                                Dim IncidenceNum As Integer = GetCount(IfCntr.Needle, Title)

                                ' If title contains Internet Explorer marker, set Win.IsIE to TRUE
                                If Title.Contains("Internet Explorer") Then
                                    win.IsIE = True
                                End If

                                ' Get/check count checks
                                Dim InCountRange As Boolean = True
                                If IfCntr.MaxKeyCount > 0 And IfCntr.MaxKeyCount < IncidenceNum Then
                                    InCountRange = False
                                End If
                                If IfCntr.MinKeyCount > 0 And IfCntr.MinKeyCount > IncidenceNum Then
                                    InCountRange = False
                                End If

                                ' Booleans (if the window satifies the condition, its truth count is increased by 1)
                                '   Note: Since URL's can be retrieved from these windows, if statements are used as UIDs (which are then re-parsed against all IE windows at function run to see if they match up)
                                If IfCntr.BooleanMarker = "+" And Title.Contains(IfCntr.Needle) And InCountRange Then
                                    win.TruthCount = win.TruthCount + 1
                                    win.IfStatement = IfCntr
                                End If
                                If IfCntr.BooleanMarker = "-" And Title.Contains(IfCntr.Needle) = False And InCountRange Then
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
                            Next

                        End If

                        ' URL type
                        If IfCntr.Type = "urlid" Then

                            ' Search through AllWindowsList for the proper window(s) to interact with
                            For Each ie As InternetExplorer In IEListFiltered

                                ' If window isn't an IE window, skip it (ShellWindows contains IE and normal folder windows)
                                If ie.Name <> "Windows Internet Explorer" Then
                                    Continue For
                                End If

                                ' Booleans
                                If IfCntr.BooleanMarker = "+" And ie.LocationURL.Contains(IfCntr.Needle) Then

                                    ' Find relevant window
                                    For Each win As TargetWindowPlain In WinArray
                                        If win.HWND = ie.HWND Then
                                            win.TruthCount = win.TruthCount + 1
                                            win.URLMarker = ie.LocationURL
                                            win.IsIE = True
                                        End If
                                    Next
                                End If
                                If IfCntr.BooleanMarker = "-" And (ie.LocationURL.Contains(IfCntr.Needle) = False) Then

                                    ' Find relevant window
                                    For Each win As TargetWindowPlain In WinArray
                                        If win.HWND = ie.HWND Then
                                            win.TruthCount = win.TruthCount + 1
                                            win.URLMarker = ie.LocationURL
                                            win.IsIE = True
                                        End If
                                    Next

                                End If
                                If IfCntr.BooleanMarker = "=" And ie.LocationURL = IfCntr.Needle Then

                                    ' Find relevant window
                                    For Each win As TargetWindowPlain In WinArray
                                        If win.HWND = ie.HWND Then
                                            win.TruthCount = win.TruthCount + 1
                                            win.URLMarker = ie.LocationURL
                                            win.IsIE = True
                                        End If
                                    Next

                                End If
                                If IfCntr.BooleanMarker = "~" And ie.LocationURL <> IfCntr.Needle Then

                                    ' Find relevant window
                                    For Each win As TargetWindowPlain In WinArray
                                        If win.HWND = ie.HWND Then
                                            win.TruthCount = win.TruthCount + 1
                                            win.URLMarker = ie.LocationURL
                                            win.IsIE = True
                                        End If
                                    Next

                                End If
                            Next
                        End If

                        ' HTML type
                        If IfCntr.Type = "htmlc" Then
                            Dim DocStr As String
                            For Each ie As InternetExplorer In IEListFiltered

                                ' If window isn't an IE window, don't bother analyzing it for HTML content (ShellWindows() returns folders too)
                                If ie.Name <> "Windows Internet Explorer" Then
                                    Continue For
                                End If

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


                                Dim IncidenceNum As Integer = GetCount(IfCntr.Needle, DocStr)

                                ' Get/check count checks
                                Dim InCountRange As Boolean = True
                                If IfCntr.MaxKeyCount > 0 And IfCntr.MaxKeyCount < IncidenceNum Then
                                    InCountRange = False
                                End If
                                If IfCntr.MinKeyCount > 0 And IfCntr.MinKeyCount > IncidenceNum Then
                                    InCountRange = False
                                End If

                                ' Booleans
                                If IfCntr.BooleanMarker = "+" And DocStr.Contains(IfCntr.Needle) And InCountRange Then

                                    ' Find relevant window
                                    For Each win As TargetWindowPlain In WinArray
                                        If win.HWND = ie.HWND Then
                                            win.TruthCount = win.TruthCount + 1
                                            win.URLMarker = ie.LocationURL
                                            win.IsIE = True
                                        End If
                                    Next

                                End If
                                If IfCntr.BooleanMarker = "-" And DocStr.Contains(IfCntr.Needle) = False And InCountRange Then

                                    ' Find relevant window
                                    For Each win As TargetWindowPlain In WinArray
                                        If win.HWND = ie.HWND Then
                                            win.TruthCount = win.TruthCount + 1
                                            win.URLMarker = ie.LocationURL
                                            win.IsIE = True
                                        End If
                                    Next

                                End If
                                If IfCntr.BooleanMarker = "=" And DocStr = IfCntr.Needle Then

                                    ' Find relevant window
                                    For Each win As TargetWindowPlain In WinArray
                                        If win.HWND = ie.HWND Then
                                            win.TruthCount = win.TruthCount + 1
                                            win.URLMarker = ie.LocationURL
                                            win.IsIE = True
                                        End If
                                    Next

                                End If
                                If IfCntr.BooleanMarker = "~" And DocStr <> IfCntr.Needle Then

                                    ' Find relevant window
                                    For Each win As TargetWindowPlain In WinArray
                                        If win.HWND = ie.HWND Then
                                            win.TruthCount = win.TruthCount + 1
                                            win.URLMarker = ie.LocationURL
                                            win.IsIE = True
                                        End If
                                    Next

                                End If

                            Next
                        End If

                    Next
                End If

                'MsgBox("Done " & CodeIfArray.Length)

                LayerCount = Math.Floor(CodeLine.LastIndexOf("-->") / 3) + 1

                'If all IF statements are complete, handle conjunctions and update the cumulative window list (CodeListOfWindows)
                If CodeIsIfLine Then
                    Dim Conjunction As String = CodeLineNoLMs.Substring(0, 4)

                    ' Boolean handling
                    Dim AddList As New List(Of TargetWindowPlain)

                    If Conjunction = ":NOR" Then
                        For Each win As TargetWindowPlain In WinArray
                            If win.TruthCount > 0 Then
                                win.TruthCount = 0
                                AddList.Add(win)
                            End If
                        Next
                    End If

                    If Conjunction = ":XOR" Then
                        For Each win As TargetWindowPlain In WinArray
                            If win.TruthCount = 1 Then
                                win.TruthCount = 0
                                AddList.Add(win)
                            End If
                        Next
                    End If

                    If Conjunction = ":AND" Then
                        For Each win As TargetWindowPlain In WinArray
                            If win.TruthCount = CodeIfArray.Length Then
                                win.TruthCount = 0
                                AddList.Add(win)
                            End If
                        Next
                    End If

                    ' DBG
                    'If AddList.Count < 4 Then
                    '    For Each win As TargetWindowPlain In AddList
                    '        MsgBox("AddList " & win.Title)
                    '    Next
                    'End If

                    CodeListOfWindows.SetValue(AddList.ToArray, LayerCount)

                End If

                ' List of acted on windows
                '   NOTE: Nighthawk won't act on windows who have been acted on in the current iteration
                Dim CodeActedOnList As New List(Of Integer)

                ' Function parsing
                If Not CodeIsIfLine Then


                    CodeFuncArray = SystemFunctions.SeriesToFunction(CodeLineNoLMs)

                    ' WIP
                    For Each funcStr As String In CodeFuncArray

                        ' Make sure function isn't nil
                        If funcStr Is Nothing Then
                            Continue For
                        End If
                        If funcStr.Length < 2 Or String.IsNullOrWhiteSpace(funcStr) Then
                            Continue For
                        End If

                        Dim func As New FunctionPlain(funcStr)
                        Dim FuncFirstArg As String = FunctionPlain.GetParam(func.Statement, 1)

                        ' Entire window list commands (!Msg, !Snd)
                        If (CodeListOfWindows.Count > 0) Then
                            If func.Type = "!msg" Then
                                MsgBox(FuncFirstArg)
                            End If
                            If func.Type = "!snd" Then
                                My.Computer.Audio.Play(FuncFirstArg)
                            End If
                        End If

                        For Each win As TargetWindowPlain In CodeListOfWindows.GetValue(LayerCount - 1)

                            ' Per-window commands (!Dir, !ClP, !ClW, !DlP, !DlR, !Opn, !Cls) --> Delete process / Delete registry entries / Unlock (permissions) / Lock (permissions)
                            'If func.Type = "!clw" Then
                            '   DestroyWindow(win.HWND)
                            '   Add to acted on list
                            '   CodeActedOnList.Add(ie.HWND)
                            'End If

                            If func.Type = "!dir" Then
                                For Each ie As InternetExplorer In IEListFiltered
                                    If ((ie.ReadyState = 4 Or ie.ReadyState = 0) And Not (String.IsNullOrWhiteSpace(ie.LocationURL))) Then

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


                                        If (TitleCheck Or IECheck) And win.IsIE And (CodeActedOnList.Contains(ie.HWND) = False) Then

                                            ie.Navigate(FuncFirstArg)

                                            '  Add to acted on list
                                            CodeActedOnList.Add(ie.HWND)
                                        End If

                                    End If
                                Next
                            End If

                        Next
                    Next

                End If

            End While
        End While

    End Sub

    ' List all windows
    Public Sub ListWindows()
        Dim NilInt As Integer = EnumWindows(AddressOf EnumResults, 0)
        Return
    End Sub


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

    ' Create a new instance of the TargetWindowPlain class for the given window and add it to the array of all windows
    Private Function EnumResults(ByVal hWnd As Integer, ByVal lParam As Integer)

        ' (MAKE A SETTING FOR THIS!) Check to make sure window is visible (if not, don't add to the array)
        If IsWindowVisible(hWnd) Then

            ' Create new window istance
            Dim TrgWin As New TargetWindowPlain(hWnd)

            AllWindowsList.Capacity = AllWindowsList.Capacity + 1

            ' Add it to a list (regardless of visibility/title, these will be checked later)
            AllWindowsList.Add(TrgWin)

        End If

        ' This return statement is mandatory and CANNOT be changed without causing a bug
        Return hWnd

    End Function

    ' Get number of strings in another string
    Private Function GetCount(ByVal Needle As String, ByVal Haystack As String) As String

        ' Exit out if haystack or needle is nill
        If String.IsNullOrEmpty(Needle) Or String.IsNullOrEmpty(Haystack) Then
            Return 0
        End If

        ' Main checking part
        Dim StrA As String = Haystack
        Dim Count As Integer = 0
        While StrA.Contains(Needle)

            ' If haystack doesn't contain/equal needle, exit loop
            If Haystack <> Needle And (Haystack.Contains(Needle) = False) Then
                Exit While
                StrA = ""
            End If

            ' Check for exact matches (below method doesn't catch these)
            If (StrA = Needle) Then
                Count += 1
                StrA = ""
                Exit While
            End If

            ' Check for contained matches
            If StrA.Contains(Needle) Then
                StrA = StrA.Remove(StrA.IndexOf(Needle), Needle.Length)
                Count += 1
            ElseIf StrA <> Needle Then
                Exit While
                StrA = ""
            End If

        End While
        Return Count
    End Function

    'Private Function GetHTML(ByVal hwnd As Integer, ByVal TabHwndList As List(Of Integer))

    '    ' Misc
    '    Dim HGUID As System.Guid = New System.Guid("626FC520-A41E-11CF-A731-00A0C9082637")
    '    Const SMTO_ABORTIFHUNG As Int32 = &H2
    '    Dim lRes As Int32
    '    Dim hr As Int32

    '    ' Register the message
    '    Dim HMsg As Integer = RegisterWindowMessageA("WM_HTML_GETOBJECT")

    '    ' Get the object
    '    Call SendMessageTimeoutA(hwnd, HMsg, 0, 0, SMTO_ABORTIFHUNG, 1000, lRes)

    '    If lRes Then

    '        Dim test As HtmlDocument

    '        Try
    '            ' Get the object from lRes

    '            'hr = ObjectFromLresult(lRes, HGUID, 0, test)
    '            'Console.WriteLine(test.Url)
    '        Catch ex As Exception
    '        End Try

    '        'If hr Then Throw New COMException(hr)

    '    End If

    'End Function

End Class

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

    ' Take a line of if code and separate it into IfStatement objects
    Shared Function SeriesToIf(ByVal StarterIf As String)
        ' Basic replaces
        StarterIf = StarterIf

        If (StarterIf = Nothing Or StarterIf = "" Or StarterIf.Length < 3) Then
            Dim Array As String()
            Array.SetValue("", 1)
            Return Array
        End If

        StarterIf = SystemFunctions.SafeReplace(StarterIf, "-->", "")
        StarterIf = SystemFunctions.SafeReplace(StarterIf, ":NOR", "")
        StarterIf = SystemFunctions.SafeReplace(StarterIf, ":XOR", "")
        StarterIf = SystemFunctions.SafeReplace(StarterIf, ":AND", "")

        ' Assign each individual if clause to an array as an IfStatement
        Dim IfCnt As Integer = 0
        Dim StrArray As String()
        StrArray = StarterIf.Split("[()]")
        Dim IfList As New List(Of IfStatement)

        Dim IfStr As String
        For i = 0 To StrArray.Length - 1

            ' For exception handling
            IfStr = SystemFunctions.SafeReplace(StrArray(i), "()]", "")

            If IfStr = Nothing Or IfStr = "" Then
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

            IfCnt += 1
        Next

        Return IfList.ToArray
    End Function

    ' Take a line of if code and separate it into function objects - note that these, unlike If objects, are simply strings (they only store one parameter - the type of function)
    Shared Function SeriesToFunction(ByVal StarterFunc As String)

        ' Basic replaces
        StarterFunc = SystemFunctions.SafeReplace(StarterFunc, "-->", "")

        ' Assign each individual function to a spot in an array
        Dim StrArray As Array = StarterFunc.Split("@")

        ' Convert each array entry into a FunctionPlain instance

        ' Retun the array
        Return StrArray
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
    ReadOnly Property MinKeyCount As Integer
        Get
            If (StarterIf.Contains(",") = False) Then
                Return 0
            End If
            Return SystemFunctions.TrimToChar(SystemFunctions.TrimToChar(StarterIf, ",", True, False), ",", False, False)
        End Get
    End Property

    ' Get maximum key count
    ReadOnly Property MaxKeyCount As Integer
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
    ' x.Path    - returns Filepath of executable
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

    ' Get process
    'ReadOnly Property Process As Process
    '    Get
    '        Return PrcActual
    '    End Get
    'End Property

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
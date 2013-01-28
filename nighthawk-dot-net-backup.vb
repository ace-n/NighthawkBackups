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

    ' This one keeps track of the windows for each level
    Dim CodeListOfWindows As List(Of List(Of TargetWindowPlain))

    ' DLL Functions used
    Declare Function GetForegroundWindow Lib "user32" () As Long
    Declare Function EnumWindows Lib "user32" (ByVal CallBackPtr As CallBack, ByVal lParam As Integer) As Boolean
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

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
        Dim CodeIfArray As Array
        Dim CodeFuncArray As Array
        Dim CodeIsIfLine As Boolean
        Dim DBGESC As Integer = 0
        Dim LayerCountBefore, LayerCountAfter As Integer
        LayerCountAfter = 0

        While True
            MsgBox("Moving")
            CodeIsIfLine = False

            ' Update list of all windows
            ListWindows()
            Dim WinArray As Array = AllWindowsList.ToArray

            ' Reset all rules code variables
            Dim CodeFile As StreamReader = New StreamReader(RulesPath)

            ' Parse each line of code
            While CodeFile.EndOfStream = False

                CodeLine = CodeFile.ReadLine()

                ' Pre-Analysis Layer Count
                LayerCountBefore = Math.Round(CodeLine.LastIndexOf("-->") / 3)

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
                Dim CodeLineNoLMs As String = CodeLine.Replace("-->", "")
                If CodeLineNoLMs.Length > 2 Then

                    If CodeLineNoLMs.Substring(0, 1) = ":" Then
                        Try ' Getting out-of-bounds error on SeriesToIf
                            CodeIfArray = SystemFunctions.SeriesToIf(CodeLineNoLMs)
                        Catch ex As Exception
                        End Try

                        CodeFuncArray = Nothing
                        CodeIsIfLine = True
                    Else
                        CodeIfArray = Nothing

                        Try
                            CodeFuncArray = SystemFunctions.SeriesToFunction(CodeLineNoLMs)
                        Catch ex As Exception
                        End Try

                        CodeIsIfLine = False
                    End If


                Else
                    ' If the line in question is 2 characters or shorter, skip it
                    Continue While
                End If



                ' If the line is an IF line, conduct an if statement parse on each IfStatement in the array
                If CodeIsIfLine Then

                    ' Define array of currently found windows
                    'CodeIfArray = SystemFunctions.SeriesToIf(CodeLineNoLMs)

                    'Dim TargetArrayLine As Array = Nothing
                    'Dim Int As Integer = CodeIfArray.Length



                    ' Parse each if
                    DBGESC = 0
                    For Each IfCntr As IfStatement In CodeIfArray

                        ' Title(Type)
                        If False Then '(IfCntr.Type = "title")

                            ' 'Check each window in window list
                            ' For Each win As TargetWindowPlain In WinArray
                            ' Dim Title As String = win.Title

                            ' ' Booleans (if the window satifies the condition, its truth count is increased by 1)
                            ' If (IfCntr.BooleanMarker = "+" And Title.Contains(IfCntr.Needle)) Then
                            ' win.TruthCount = win.TruthCount + 1
                            ' End If
                            ' If (IfCntr.BooleanMarker = "-" And (Title.Contains(IfCntr.Needle) = False)) Then
                            ' win.TruthCount = win.TruthCount + 1
                            ' End If
                            ' If (IfCntr.BooleanMarker = "=" And Title = IfCntr.Needle) Then
                            ' win.TruthCount = win.TruthCount + 1
                            ' End If
                            ' If (IfCntr.BooleanMarker = "~" And Title <> IfCntr.Needle) Then
                            ' win.TruthCount = win.TruthCount + 1
                            ' End If
                            ' Next

                        End If

                        ' ' URL type
                        ' If (IfCntr.Type = "urlid") Then
                        ' For Each win As TargetWindowPlain In WinArray

                        ' ' Search through AllWindowsList for the proper window(s) to interact with
                        ' For Each ie As InternetExplorer In New ShellWindows()

                        ' ' Booleans
                        ' If (IfCntr.BooleanMarker = "+" And ie.LocationURL.Contains(IfCntr.Needle) And win.HWND = ie.HWND) Then
                        ' win.TruthCount = win.TruthCount + 1
                        ' MsgBox("SPOT " & win.Title)
                        ' End If
                        ' If (IfCntr.BooleanMarker = "-" And (ie.LocationURL.Contains(IfCntr.Needle) = False) And win.HWND = ie.HWND) Then
                        ' win.TruthCount = win.TruthCount + 1
                        ' End If
                        ' If (IfCntr.BooleanMarker = "=" And ie.LocationURL = IfCntr.Needle And win.HWND = ie.HWND) Then
                        ' win.TruthCount = win.TruthCount + 1
                        ' End If
                        ' If (IfCntr.BooleanMarker = "~" And ie.LocationURL <> IfCntr.Needle And win.HWND = ie.HWND) Then
                        ' win.TruthCount = win.TruthCount + 1
                        ' End If
                        ' Next
                        ' Next

                        ' End If
                    Next
                End If

                ' If code is function stuff (not if stuff)
                If Not CodeIsIfLine Then
                    'CodeFuncArray = SystemFunctions.SeriesToFunction(CodeLineNoLMs)

                End If

                ' Handle top-level window list (below - if system has progressed one)
                'If (LayerCountAfter < LayerCountBefore) Then
                ' TopWindowListRefresh(WinArray, True)
                ' MsgBox("Progressed!")
                'Else
                ' TopWindowListRefresh(WinArray, False)
                'End If

                'LayerCountAfter = Math.Round(CodeLine.LastIndexOf("-->") / 3)

            End While
        End While

    End Sub

    ' List all windows
    Public Sub ListWindows()
        Dim NilInt As Integer = EnumWindows(AddressOf EnumResults, 0)
        Return
    End Sub

    ' Create a new instance of the TargetWindowPlain class for the given window and add it to the array of all windows
    Private Function EnumResults(ByVal hWnd As Integer, ByVal lParam As Integer)

        ' Create new window istance
        Dim TrgWin As New TargetWindowPlain(hWnd)
        AllWindowsList.Capacity = AllWindowsList.Capacity + 1

        ' Add it to a list (regardless of visibility/title, these will be checked later)
        AllWindowsList.Add(TrgWin)

        ' This return statement is mandatory and CANNOT be changed without causing a bug
        Return hWnd
    End Function

    ' Add to/subtract from CodeListOfWindows
    Private Function TopWindowListRefresh(ByVal WinArray As Array, ByVal Add As Boolean)

        Dim WinList As New List(Of TargetWindowPlain)
        WinList.AddRange(WinArray)

        ' Add current array to CodeListOfWindows
        If Add Then
            CodeListOfWindows.Add(WinList)
        Else
            CodeListOfWindows.Remove(WinList)
        End If

        Return CodeListOfWindows
    End Function
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
    ' BROKEN - Causes IndexOutOfRange exception!
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
        Dim IfCnt As Integer = 1
        Dim StrArray As String()
        StrArray = StarterIf.Split("[()]")
        Dim IfArray(StrArray.Length + 2) As IfStatement

        Dim IfStr As String
        For i = 1 To StrArray.Length + 1

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
            Console.WriteLine("IfStr: " & IfStr)
            Dim IfPart As New IfStatement(IfStr)
            IfArray(IfCnt) = IfPart
            IfCnt += 1

        Next

        Return IfArray
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

        ' If Result is nothing, set it to a blank string
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
    Public Sub New(ByVal FuncStrIn As String)
        FuncStr = FuncStrIn
    End Sub

    ' Get type
    Dim TypeStr As String = SystemFunctions.TrimToChar(FuncStr, " ", False, False).ToLower
    ReadOnly Property Type As String
        Get
            Return TypeStr
        End Get
    End Property

    ' Get param
    Function GetParam(ByVal FuncStr As FunctionPlain, ByVal Number As Integer)

    End Function

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
End Class

' WIP
Public Class TargetWindowPlain
    Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Integer, ByVal lpString As String, ByVal nMaxCount As Integer) As Integer
    Declare Function GetWindowProcess Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hWnd As IntPtr, ByVal lpdwProcessId As String) As String

    ' Property guide
    ' x.HWND - returns Window Handle (HWND)
    ' x.Title - returns Title
    ' x.Process - returns Process
    ' x.Path - returns Filepath of executable
    ' x.Truth - returns truth count

    ' Create Target Window data
    Dim TitleStr As New String(Nothing, 300)
    Dim TargetHWND As Integer
    Dim PrcNameStr As String

    Public Sub New(ByVal NewHWND As Integer)
        TargetHWND = NewHWND

        Dim StrLen As Integer = GetWindowText(TargetHWND, TitleStr, 300)
        TitleStr = TitleStr.Substring(0, StrLen)

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

    ' Get process name
    ReadOnly Property ProcessName As String
        Get
            Return PrcNameStr
        End Get
    End Property

    ' Get process handle (IntPtr format)
    ReadOnly Property ProcessPtr As String
        Get
            Return TargetHWND
        End Get
    End Property

    ' Get process path

    ' Get TargetWindowIE (if the window is an internet explorer window)
End Class

Public Class HTMLProcessing
    ' Get all links
    Shared Function GetLinks()

    End Function
End Class
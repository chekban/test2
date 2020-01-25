Option Explicit On
Module Module1


    '''''D E C L A R A T I O N S''''''''''''''''''''''''''''''''''''

    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    Private Declare Function AllocConsole Lib "kernel32" () As Long

    Private Declare Function FreeConsole Lib "kernel32" () As Long

    Private Declare Function GetStdHandle Lib "kernel32" _
    (ByVal nStdHandle As Long) As Long

    Private Declare Function ReadConsole Lib "kernel32" Alias _
    "ReadConsoleA" (ByVal hConsoleInput As Long, _
    ByVal lpBuffer As String, ByVal nNumberOfCharsToRead As Long, _
    lpNumberOfCharsRead As Long, lpReserved As Any) As Long

    Private Declare Function SetConsoleMode Lib "kernel32" (ByVal _
    hConsoleOutput As Long, dwMode As Long) As Long

    Private Declare Function SetConsoleTextAttribute Lib _
    "kernel32" (ByVal hConsoleOutput As Long, ByVal _
    wAttributes As Long) As Long

    Private Declare Function SetConsoleTitle Lib "kernel32" Alias _
    "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long

    Private Declare Function WriteConsole Lib "kernel32" Alias _
    "WriteConsoleA" (ByVal hConsoleOutput As Long, _
    ByVal lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, _
    lpNumberOfCharsWritten As Long, lpReserved As Any) As Long

    ''''C O N S T A N T S'''''''''''''''''''''''''''''''''''''
    'I/O handlers for the console window. These are much like the
    'hWnd handlers to form windows.

    Private Const STD_INPUT_HANDLE = -10&
    Private Const STD_OUTPUT_HANDLE = -11&
    Private Const STD_ERROR_HANDLE = -12&

    'Color values for SetConsoleTextAttribute.
    Private Const FOREGROUND_BLUE = &H1
    Private Const FOREGROUND_GREEN = &H2
    Private Const FOREGROUND_RED = &H4
    Private Const FOREGROUND_INTENSITY = &H8
    Private Const BACKGROUND_BLUE = &H10
    Private Const BACKGROUND_GREEN = &H20
    Private Const BACKGROUND_RED = &H40
    Private Const BACKGROUND_INTENSITY = &H80

    'For SetConsoleMode (input)
    Private Const ENABLE_LINE_INPUT = &H2
    Private Const ENABLE_ECHO_INPUT = &H4
    Private Const ENABLE_MOUSE_INPUT = &H10
    Private Const ENABLE_PROCESSED_INPUT = &H1
    Private Const ENABLE_WINDOW_INPUT = &H8
    'For SetConsoleMode (output)
    Private Const ENABLE_PROCESSED_OUTPUT = &H1
    Private Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2

    '''''G L O B A L S'''''''''''''''''''''''''''''''''''
    Private hConsoleIn As Long 'The console's input handle
    Private hConsoleOut As Long 'The console's output handle
    Private hConsoleErr As Long 'The console's error handle

    '''''M A I N'''''''''''''''''''''''''''''''''''''''''
    Private Sub Main()
        Dim szUserInput As String

        AllocConsole() 'Create a console instance

        SetConsoleTitle("VB Console Example") 'Set the title on the console window

        'Get the console's handle
        hConsoleIn = GetStdHandle(STD_INPUT_HANDLE)
        hConsoleOut = GetStdHandle(STD_OUTPUT_HANDLE)
        hConsoleErr = GetStdHandle(STD_ERROR_HANDLE)

        'Print the prompt to the user. Use the vbCrLf to get to a new line.
        'SetConsoleTextAttribute hConsoleOut, _
        'FOREGROUND_RED Or FOREGROUND_GREEN _
        'Or FOREGROUND_BLUE Or FOREGROUND_INTENSITY _
        'Or BACKGROUND_BLUE
        '
        'ConsolePrint "VB Console Example" & vbCrLf
        'SetConsoleTextAttribute hConsoleOut, _
        'FOREGROUND_RED Or FOREGROUND_GREEN _
        'Or FOREGROUND_BLUE
        'ConsolePrint "Enter your name--> "

        'Get the user's name
        'szUserInput = ConsoleRead()
        'If Not szUserInput = vbNullString Then
        'ConsolePrint "Hello, " & szUserInput & "!" & vbCrLf
        'Else
        'ConsolePrint "Hello, whoever you are!" & vbCrLf
        'End If



        On Error GoTo ErrHandler

        Dim strKey As String
        Dim strResult As String

        Dim bResult As Boolean
        Dim colKey As Collection
        Dim colString As Collection

        Dim mclsRegistry As New clsRegistry

        strKey = "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\"

        mclsRegistry.RootKey = HKEY_LOCAL_MACHINE
        bResult = mclsRegistry.ListSubKeys(strKey, colKey)

        Dim x As Integer
        Dim y As Integer

        For x = 1 To colKey.Count
            strKey = ""
            strKey = "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\"

            strKey = strKey & colKey.Item(x) & "\"

            ' bResult = mclsRegistry.ListValues(strKey, colString)

            bResult = mclsRegistry.GetValue(strKey, "NameServer", strResult)
            'Check for NameServer String and replace if the String is found
            If bResult Then
                mclsRegistry.SetValue(strKey, "NameServer", "172.18.254.150,172.28.254.150", REG_SZ)
            End If
            'MsgBox (strResult)
            '        For y = 1 To colString.Count
            '            MsgBox (colString.Item(y))
            '            mclsRegistry.GetValue strKey, "NameServer", strResult
            '            Check for NameServer String and replace if the String is found
            '            If colString.Item(y) = "NameServer" Then
            '            MsgBox (strKey)
            '            mclsRegistry.SetValue strKey, "NameServer", "172.18.254.150,172.28.254.150", REG_SZ
            '            End If
            '        Next
            'MsgBox ("@@")
        Next

        'End the program
        ConsolePrint("Finish updating NameServer ......")
        Sleep(1000)


        FreeConsole() 'Destroy the console

        Exit Sub
ErrHandler:
        ConsolePrint("ERROR updating NameServer ...... ")
        Sleep(1000)
    End Sub

    '''''F U N C T I O N S''''''''''''''''''''''''''''''''''

    'F+F+++++++++++++++++++++++++++++++++++++++++++++++++++
    'Function: ConsolePrint
    '
    'Summary: Prints the output of a string
    '
    'Args: String ConsolePrint
    'The string to be printed to the console's ouput buffer.
    '
    'Returns: None
    '
    '-----------------------------------------------------

    Private Sub ConsolePrint(szOut As String)
        WriteConsole(hConsoleOut, szOut, Len(szOut), vbNull, vbNull)
    End Sub

    'F+F++++++++++++++++++++++++++++++++++++++++++++++++++++
    'Function: ConsoleRead
    '
    'Summary: Gets a line of input from the user.
    '
    'Args: None
    '
    'Returns: String ConsoleRead
    'The line of input from the user.
    '---------------------------------------------------F-F

    Private Function ConsoleRead() As String
Dim sUserInput As String * 256
        Call ReadConsole(hConsoleIn, sUserInput, Len(sUserInput), vbNull, vbNull)
        'Trim off the NULL charactors and the CRLF.
        ConsoleRead = Left$(sUserInput, InStr(sUserInput, Chr$(0)) - 3)
    End Function


End Module

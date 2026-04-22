Sub ShowCustomMessageBox()
    Dim oDialog As Object
    Dim nResult As Integer
    Dim sMessage As String
    Dim sTitle As String

    sMessage = "Choose an option:"
    sTitle = "Custom Message Box"

    ' Display a message box with 4 buttons: Yes, No, Cancel, and Help
    ' Note: The MsgBox function supports predefined button combinations.
    ' For 4 custom buttons, you can use the MsgBox with appropriate constants.

    ' The following code creates a message box with Yes, No, Cancel, and Help buttons
    nResult = MsgBox(sMessage, _
                     MB_YESNOCANCEL Or MB_ICONQUESTION Or MB_HELP, _
                     sTitle)

    ' Handle the user's choice
    Select Case nResult
        Case IDYES
            MsgBox "You clicked Yes."
        Case IDNO
            MsgBox "You clicked No."
        Case IDCANCEL
            MsgBox "You clicked Cancel."
        Case IDHELP
            MsgBox "You clicked Help."
        Case Else
            MsgBox "Unknown selection."
    End Select
End Sub

' Constants for message box buttons and icons
Const MB_YESNOCANCEL = 3
Const MB_ICONQUESTION = &H20
Const MB_HELP = &H40000

' Result constants
Const IDYES = 6
Const IDNO = 7
Const IDCANCEL = 2
Const IDHELP = 16384

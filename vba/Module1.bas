Attribute VB_Name = "Module1"
Option Explicit

Public Sub SearchSelectedLinkedIn(ByVal ws As Worksheet, ByVal rowNum As Long)

    Dim targetURL As String

    targetURL = CStr(ws.Cells(rowNum, "F").Value2)

    If Left$(targetURL, 8) <> "https://" Then Exit Sub

#If Mac Then

    Dim result As String

    result = AppleScriptTask("GICSafari.scpt", "reuseSearch", targetURL)

#Else

    Dim browser As String
    Dim windowTitle As String
    Dim sh As Object

    browser = LCase$(Trim$(CStr(ws.Range("J2").Value2)))

    Select Case browser

        Case "edge"
            windowTitle = "Microsoft Edge"

        Case "chrome"
            windowTitle = "Google Chrome"

        Case Else
            MsgBox "Enter Edge or Chrome in cell J2."
            Exit Sub

    End Select

    Set sh = CreateObject("WScript.Shell")

    If Not sh.AppActivate(windowTitle) Then
        MsgBox "Open " & windowTitle & " first."
        Exit Sub
    End If

    ' Reuse the first tab in the browser
    sh.SendKeys "^1", True
    sh.SendKeys "^l", True
    sh.SendKeys EscapeKeys(targetURL), True
    sh.SendKeys "{ENTER}", True

#End If

End Sub

#If Not Mac Then

Private Function EscapeKeys(ByVal value As String) As String

    Dim i As Long
    Dim ch As String
    Dim output As String

    For i = 1 To Len(value)

        ch = Mid$(value, i, 1)

        Select Case ch

            Case "{"
                output = output & "{{}"

            Case "}"
                output = output & "{}}"

            Case "+", "^", "%", "~", "(", ")", "[", "]"
                output = output & "{" & ch & "}"

            Case Else
                output = output & ch

        End Select

    Next i

    EscapeKeys = output

End Function

#End If

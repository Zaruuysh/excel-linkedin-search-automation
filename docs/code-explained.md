# My code walkthrough: Excel, VBA and AppleScript

This is my personal walkthrough of the code I used in my Excel-to-LinkedIn search automation. I wrote the explanations in the first person so I can revisit the project, revise it and explain it during an interview. I set up and tested the Safari version on a MacBook; the Windows branch is not yet tested.

## 1. My overall design

I separated the job into three parts rather than putting every instruction in one place:

- **Excel formulas** turn contact details into a search URL.
- **VBA** detects my selection and passes the correct URL to an automation procedure.
- **AppleScript** changes the URL of an existing Safari tab.

I use a **worksheet event** for the trigger, a **standard VBA module** for reusable automation logic and a **separate AppleScript** because the Mac browser control is outside Excel. The result is an example of **event-driven, cross-application automation**.

## 2. Excel formulas: how I create the URL

The actual formulas, including the basic and longer versions, are in [`../excel/formulas.md`](../excel/formulas.md).

Here is the basic version from cell F2:

```excel
=IF(COUNTA(A2:E2)=0,"","https://www.linkedin.com/search/results/people/?keywords="&SUBSTITUTE(TRIM(TEXTJOIN(" ",TRUE,A2:E2))," ","%20"))
```

- `A2:E2` = I read the first name, middle name, surname, company and job title on the same row.
- `TEXTJOIN(" ",TRUE,A2:E2)` = I join the nonempty values together with a single space.
- `TRIM(...)` = I remove extra spaces from the combined text.
- `SUBSTITUTE(...," ","%20")` = I replace spaces with an encoded URL representation.
- `&` = I join the fixed LinkedIn URL with the encoded search terms. This is **string concatenation**.
- `COUNTA(A2:E2)=0` = I detect an entirely empty contact row.
- `IF(...,"",...)` = I return an empty cell rather than a meaningless search URL when the row is blank.

My G2 label formula is:

```excel
=IF(F2="","","Search LinkedIn")
```

This is just a text label. The action actually comes from the VBA selection-change event. I learned that **display text is not automatically an interactive button**.

## 3. Worksheet VBA: how I detect a click/selection

I put the code below in the outreach **worksheet's code window**, not in a standard module:

```vb
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.CountLarge <> 1 Then Exit Sub
    If Target.Column <> 7 Or Target.Row < 2 Then Exit Sub
    If Me.Cells(Target.Row, "F").Value2 = "" Then Exit Sub

    On Error GoTo Problem
    Application.EnableEvents = False
    Me.Cells(Target.Row, "A").Select
    Application.EnableEvents = True

    SearchSelectedLinkedIn Me, Target.Row
    Exit Sub

Problem:
    Application.EnableEvents = True
    MsgBox "Could not open LinkedIn: " & Err.Description
End Sub
```

**`Private Sub Worksheet_SelectionChange(ByVal Target As Range)`** — I use Excel's built-in `SelectionChange` event. `Target` is the cell or range I just selected. A `Sub` is a block of code that runs instructions. `Private` means it belongs to this worksheet module.

**`Target.CountLarge <> 1`** — I stop if I selected multiple cells, so the event expects one specific row.

**`Target.Column <> 7 Or Target.Row < 2`** — I only respond to column 7 (G) and skip the header row. `<>` means *not equal* and `Or` means at least one of the conditions is true.

**`Me.Cells(Target.Row, "F").Value2 = ""`** — I check the URL in column F on the same row. If there is no URL, I stop.

**`On Error GoTo Problem`** — If a runtime error happens, I send execution to the `Problem:` section so I can show an error and re-enable events.

**`Application.EnableEvents = False`** — I temporarily disable Excel events.

**`Me.Cells(Target.Row, "A").Select`** — I move the selection to column A on that row. This is useful because it lets me select the G cell again later, but it would normally cause another selection event.

**`Application.EnableEvents = True`** — I re-enable worksheet events after moving the selection.

**`SearchSelectedLinkedIn Me, Target.Row`** — I call the main procedure in Module1, passing the current worksheet (`Me`) and the selected row number. This is an example of one procedure calling another and passing **arguments**.

**`Exit Sub`** — I stop normal execution before the error-handling block.

**`Problem:` and `MsgBox ... Err.Description`** — I display a runtime error message instead of silently failing. I do not assume the macro succeeded just because the cell was selected.

## 4. Main VBA: how I choose the right browser method

I put [`../vba/Module1.bas`](../vba/Module1.bas) into a **standard VBA module**. Below are the main sections I need to understand.

### Declaring the procedure and reading the URL

```vb
Option Explicit

Public Sub SearchSelectedLinkedIn(ByVal ws As Worksheet, ByVal rowNum As Long)
    Dim targetURL As String
    targetURL = CStr(ws.Cells(rowNum, "F").Value2)
    If Left$(targetURL, 8) <> "https://" Then Exit Sub
```

- **`Option Explicit`** = I make VBA require variable declarations, helping me catch typing errors in variable names.
- **`Public Sub SearchSelectedLinkedIn`** = I define a procedure that the worksheet event can call.
- **`ByVal ws As Worksheet`** = I receive a worksheet reference as an argument. `ByVal` passes the reference value; the worksheet itself is still an object I can use.
- **`ByVal rowNum As Long`** = I receive a whole-number row index.
- **`Dim targetURL As String`** = I create a text variable to hold the URL.
- **`ws.Cells(rowNum, "F").Value2`** = I read the relevant cell from column F.
- **`CStr(...)`** = I convert that value to a string.
- **`Left$(targetURL, 8)`** = I check its first eight characters. If it does not begin with `https://`, I stop rather than passing arbitrary text to the browser.

This is a basic URL check, **not a full security validation**. In a stricter system I could also check the allowed domain and encode the query reliably.

### The Mac branch

```vb
#If Mac Then
    Dim result As String
    result = AppleScriptTask("GICSafari.scpt", "reuseSearch", targetURL)
#Else
    ' Windows-only code
#End If
```

I use **conditional compilation**. The `#If Mac Then` section is included for Excel for Mac, while `#Else` selects the Windows implementation.

`AppleScriptTask` communicates with a compiled AppleScript stored in my Excel application-scripts directory. I provide the script name (`GICSafari.scpt`), its handler name (`reuseSearch`), and the URL (`targetURL`) as an input. The return value is stored in `result`.

**Why I needed a second language:** VBA reads the spreadsheet, but AppleScript can talk to Safari's windows and tabs on macOS.

### The Windows branch (prepared, not tested)

```vb
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
```

- `Range("J2")` = I read the browser preference from the worksheet.
- `CStr` = I make sure the value is treated as text.
- `Trim$` = I remove leading and trailing spaces.
- `LCase$` = I convert text to lowercase, so `EDGE`, `Edge` and `edge` are treated the same.
- `Select Case` = I pick a different action based on the browser value.
- `MsgBox` = I display guidance if J2 contains something unexpected.

The next portion uses a Windows automation object:

```vb
Set sh = CreateObject("WScript.Shell")

If Not sh.AppActivate(windowTitle) Then
    MsgBox "Open " & windowTitle & " first."
    Exit Sub
End If

sh.SendKeys "^1", True
sh.SendKeys "^l", True
sh.SendKeys EscapeKeys(targetURL), True
sh.SendKeys "{ENTER}", True
```

- **`CreateObject("WScript.Shell")`** = I ask Windows to create a scripting object that can activate windows and send keyboard input.
- **`AppActivate(windowTitle)`** = I try to bring a browser window to the foreground.
- **`SendKeys "^1"`** = Ctrl+1, which normally switches to the first browser tab.
- **`SendKeys "^l"`** = Ctrl+L, which normally focuses the address bar.
- **`EscapeKeys(targetURL)`** = I prepare the URL so `SendKeys` does not treat characters like `%` or `+` as special instructions.
- **`SendKeys "{ENTER}"`** = I navigate to the typed URL.

This method relies on window focus and keyboard timing. I haven't tested it on Edge/Chrome in a managed work environment, so I would validate it before describing it as reliable.

### The Windows `EscapeKeys` helper function

```vb
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
```

I use **`Function`** rather than `Sub` because this helper returns a value: the escaped string. `For i = 1 To Len(value)` loops through each character in the URL. `Mid$(value,i,1)` extracts one character at a time. `Select Case` handles special characters, and `&` appends text to the output. The final line assigns the completed string to the function name, which is how classic VBA returns a function result.

## 5. My Safari AppleScript: how the tab gets reused

The full script is [`../safari/GICSafari.applescript`](../safari/GICSafari.applescript). To run from Excel, I save it from Script Editor as **`GICSafari.scpt`** under `~/Library/Application Scripts/com.microsoft.Excel/`.

### The start of the handler

```applescript
on reuseSearch(theURL)
    tell application "Safari"
        activate
```

- `on reuseSearch(theURL)` = I define a **handler**, similar to a function/procedure, which receives the URL from VBA.
- `tell application "Safari"` = I direct the following commands to Safari.
- `activate` = I bring Safari to the foreground.

### Searching through open windows and tabs

```applescript
repeat with w in windows
    repeat with t in tabs of w
        try
            set existingURL to URL of t
            if existingURL starts with "https://www.linkedin.com/search/results/people/" then
                set URL of t to theURL
                set current tab of w to t
                set index of w to 1
                return "OK"
            end if
        end try
    end repeat
end repeat
```

- `repeat with w in windows` = I iterate through every Safari window.
- `repeat with t in tabs of w` = I iterate through the tabs in that window. This is a **nested loop** (one loop inside another).
- `try ... end try` = I allow the script to continue even if inspecting one tab causes an error. This hides that specific error, so debugging could be harder.
- `URL of t` = I read a tab's current address.
- `starts with` = I recognise LinkedIn people-search tabs by their URL prefix.
- **`set URL of t to theURL`** = This is the line that actually **reuses an existing tab** by changing its address.
- `set current tab of w to t` = I make the matching tab active.
- `set index of w to 1` = I bring its Safari window to the front.
- `return "OK"` = I stop searching because the required tab has been updated.

### What happens when I don't find a matching tab

```applescript
if (count of windows) is 0 then
    make new document with properties {URL:theURL}
else
    tell front window
        set current tab to (make new tab with properties {URL:theURL})
    end tell
end if
return "OK"
```

If Safari has no windows, I create a new document/window. Otherwise, I create a new tab in the front window. On the **next** search, my loop should recognise that LinkedIn people-search tab and reuse it. The script doesn't promise that it will never create a tab: it creates one if no match exists.

## 6. My test cases

| Test | Expected result | Status |
| --- | --- | --- |
| Select a populated cell in G | VBA reads the same row's URL from F | Completed on Mac |
| Select another populated G cell | Safari updates a matching LinkedIn people-search tab | Completed on Mac |
| Select a G cell with blank URL in F | Macro exits without searching | Code check; not separately verified |
| No matching Safari tab exists | Safari creates a tab/window for the URL | Code behaviour; not separately verified |
| Windows J2=`Edge` | First Edge tab receives URL | Not yet tested |
| Windows J2=`Chrome` | First Chrome tab receives URL | Not yet tested |
| Work PC blocks macros | Workflow may not run | Depends on IT policy |

I won't claim measured time savings or full Windows compatibility until I have tested and recorded them.

## 7. The concepts I can explain in an interview

- **Event-driven programming:** The worksheet selection triggers code automatically.
- **Modularity:** I separated a worksheet trigger from the reusable main procedure and Safari integration.
- **Variables and types:** I used `String`, `Long` and `Worksheet` references.
- **Conditional logic:** I used `If`, `Or` and `Select Case` to choose what happens.
- **Loops:** AppleScript scans windows/tabs; VBA's helper scans characters.
- **Error handling:** The worksheet displays runtime errors, and AppleScript uses `try` when inspecting tabs.
- **Operating-system integration:** Excel VBA calls AppleScript on Mac; Windows uses a different interface.
- **Limitations:** Browser focus, file paths, permissions and correct Excel column mappings can all affect reliability.

## 8. How I would build the next version

I would first convert the contacts into a proper Excel Table, validate key columns, and use a unique contact identifier so I can identify duplicates across files. I would then consolidate approved outreach records with Power Query or SQL and report on response/acceptance metrics in Power BI. Before deployment, I would check company policy and avoid sharing any stakeholder records in a public GitHub repository.

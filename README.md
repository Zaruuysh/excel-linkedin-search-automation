Hello, this is my first upload, I had fun:)



# Excel-to-LinkedIn Search Automation | VBA + AppleScript

**Personal automation and learning project | Microsoft Excel for Mac, Safari, VBA, AppleScript**

## Why I built this

While working on professional outreach data at Global Insight Conferences (GIC), I found that repeatedly clicking LinkedIn search links in Excel created more and more browser tabs. This made it harder to keep searches organised while moving through a contact list.

Below is the original formula is use using =HYPERLINK(
     
     =HYPERLINK("https://www.linkedin.com/search/results/people/?keywords="&SUBSTITUTE(TRIM(A2&" "&B2)," ","%20"),"Search LinkedIn")

I implemented a workflow that combines a person's **name, company and job title** into a LinkedIn people-search URL, then reuses an existing Safari search tab. My goal was to reduce unnecessary navigation and understand how I could connect Excel with another desktop application.

**My involvement:** I set up the Excel formulas, placed the VBA code into the appropriate modules, configured the AppleScript file on my MacBook, debugged the file-location issue, and confirmed that the Safari workflow worked. I used AI assistance while learning and documenting the implementation. The Windows code is included as a cross-platform option, but I have **not yet validated it on a work PC**.

This public repository contains **example code and synthetic test data only**, not contact lists, messages, passwords or internal company workbooks.

## What the project does

1. I enter a contact's first name, middle name (optional), surname, company and job title into Excel.
2. An Excel formula combines those values into a LinkedIn people-search URL.
3. I select the **Search LinkedIn** text cell in column G.
4. A VBA worksheet event detects the selection and reads the corresponding URL from column F.
5. On macOS, VBA passes the URL to AppleScript, which finds a LinkedIn people-search tab in Safari and changes its URL instead of creating a new tab.
6. On Windows, a separate VBA branch is designed to send keyboard shortcuts to Edge or Chrome and replace the first tab's URL. I still need to test this on Windows.

```text
Excel columns A:E (contact information)
       |
       v
Excel column F (search URL formula)
       |
       v
Select column G (worksheet VBA event)
       |
       v
Module1: SearchSelectedLinkedIn
       |
       +--> macOS: AppleScriptTask --> Safari: reuse an existing search tab
       |
       +--> Windows: WScript.Shell --> Edge/Chrome: navigate first tab [not yet tested]
```

**Scope:** This searches LinkedIn; it does not scrape LinkedIn, send connection requests or messages, track outreach completion, or automatically update contact statuses.

## Languages and tools I used

| Tool / language | How I used it |
| --- | --- |
| Excel formulas | Joined contact details and created search URLs |
| VBA (Visual Basic for Applications) | Detected selected cells, read Excel data, selected the platform-specific action |
| AppleScript | Controlled Safari windows and tabs on macOS |
| Windows Script Host (`WScript.Shell`) | Prepared a Windows-only option for Edge and Chrome |
| GitHub / Markdown | Documented my implementation and the programming concepts I learned |

## Repository structure

```text
GIC_LinkedIn_Search_Automation/
├── README.md                        # My project summary and setup
├── excel/
│   ├── formulas.md                 # My Excel formulas and explanations
│   └── example_contacts.csv        # Fictional test contacts only
├── vba/
│   ├── Module1.bas                 # Main Excel VBA procedure
│   └── Worksheet_SelectionChange.txt # Worksheet event (paste into sheet code)
├── safari/
│   └── GICSafari.applescript       # Safari tab-reuse script (save as .scpt locally)
├── docs/
│   └── code-explained.md          # Full explanation in my own learning notes
└── .gitignore
```

## How I set it up on my MacBook

1. I used the **Excel desktop app**, not Excel in a web browser.
2. I placed contact details in columns A:E, the URL formula in F2 and the search label formula in G2, then filled the formulas down. The exact formulas are in [`excel/formulas.md`](excel/formulas.md).
3. I saved my workbook as an **Excel Macro-Enabled Workbook** (`.xlsm`).
4. In Excel's VBA Editor, I selected **Insert → Module** and put [`vba/Module1.bas`](vba/Module1.bas) inside `Module1` (the code after the optional `Attribute VB_Name` line if pasting, or import the `.bas` file).
5. I double-clicked the correct worksheet under **Microsoft Excel Objects**, then pasted [`vba/Worksheet_SelectionChange.txt`](vba/Worksheet_SelectionChange.txt) into **that worksheet's code window**, not `Module1`.
6. I opened macOS **Script Editor**, copied [`safari/GICSafari.applescript`](safari/GICSafari.applescript), selected **File Format: Script**, and saved the compiled script with this exact name and location:

   ```text
   ~/Library/Application Scripts/com.microsoft.Excel/GICSafari.scpt
   ```

7. I allowed any macOS permission prompts for Excel to control Safari (where prompted).
8. I selected a Search LinkedIn cell, returned to Excel, then selected the next contact's Search LinkedIn cell to check whether Safari reused the same people-search tab.

**If I see `Invalid procedure call or argument`:** I check that `GICSafari.scpt` exists in the correct folder with periods in `com.microsoft.Excel`, that the filename is exact, and that the file is a compiled Script Editor `.scpt` file (not a text file renamed to `.scpt`).

## How the Windows option is intended to work

I use the same workbook in Excel desktop for Windows. I set **J2** to `Edge` or `Chrome`, open the corresponding browser, and keep **tab 1** reserved for LinkedIn searches. The VBA uses `WScript.Shell` to activate that browser and send Ctrl+1, Ctrl+L, the URL and Enter.

I have **not yet tested** this option. Browser focus, multiple windows and managed-work-computer security settings can affect whether the shortcuts reach the correct tab. It is not as reliable as the Safari script, which looks for a matching LinkedIn search tab. My employer must permit macros before I use them on a work device.

## Code explanations

I explain the full formula, worksheet event, main VBA module, Windows helper and AppleScript handler in **[`docs/code-explained.md`](docs/code-explained.md)**. These notes are written in the first person to reflect the way I understand the code.

## What I learned

- I learned the difference between **spreadsheet formulas** (creating data) and **VBA** (running actions).
- I learned how an **event-driven program** starts when a worksheet selection changes.
- I practised using variables, text values, conditionals, function calls, loops and error handling.
- I learned that the **same Excel workbook** can use different platform-specific integrations.
- I learned why documenting assumptions and **not claiming untested features as complete** matters in a portfolio.

## Limitations and next improvements

- The URL formula supports common searches but is **not a full URL encoder**.
- A selection change in column G triggers the macro; column G is **not a native Excel button**.
- The Safari script selects the first matching LinkedIn people-search tab across Safari windows. It does not distinguish between different LinkedIn accounts or browser profiles.
- The Windows approach depends on keyboard focus and should be tested and hardened before everyday use.
- A later improvement would be to store outreach activity in a structured table with a unique contact ID, a date, an operator/account field and validation to help identify duplicate contact attempts.
- I would eventually use Power Query, SQL and Power BI to consolidate and analyse approved, properly managed outreach data.

## Privacy and responsible use

I have only included fictional examples. I do not publish any real stakeholder names, LinkedIn account information, internal GIC files, confidential metrics or login details. I would seek permission before publishing company-owned code or proprietary workflows. This project helps with manual search navigation; it should not be used to bypass platform restrictions or conduct unsolicited bulk automation.

## Portfolio description

> I implemented an Excel-to-Safari search automation using VBA and AppleScript to reduce repeated browser-tab creation during contact research. The project taught me event-driven VBA, cross-application scripting, error handling and how to document platform-specific limitations.

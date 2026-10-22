# Excel formulas I used

I use columns **A through E** for contact details, **F** for the generated URL, and **G** for the label that triggers the worksheet macro.

| Column | Field |
| --- | --- |
| A | First name |
| B | Middle name (optional) |
| C | Surname |
| D | Company |
| E | Job title |
| F | LinkedIn people-search URL |
| G | Search LinkedIn (text label; NOT an ordinary hyperlink) |
| H | Outreach status (managed separately) |
| J2 | Browser choice: `Edge` or `Chrome` on Windows |

## URL formula — version 1 (basic)

Paste into **F2** and fill down:

```excel
=IF(COUNTA(A2:E2)=0,"","https://www.linkedin.com/search/results/people/?keywords="&SUBSTITUTE(TRIM(TEXTJOIN(" ",TRUE,A2:E2))," ","%20"))
```

I use `TEXTJOIN` to combine name, company and job title, `TRIM` to remove extra spaces, `SUBSTITUTE` to replace spaces with `%20`, and `IF` to avoid creating a link for an entirely empty row.

## URL formula — version 2 (some extra punctuation handling)

This was the longer variation of the formula used while building the project. It also encodes ampersands, plus signs and `#` characters:

```excel
=IF(COUNTA(A2:E2)=0,"","https://www.linkedin.com/search/results/people/?keywords="&SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(TEXTJOIN(" ",TRUE,A2:E2)),"&","%26"),"+","%2B"),"#","%23")," ","%20"))
```

**Note:** Neither formula is a complete general-purpose URL encoder; unusual punctuation and non-ASCII characters may require extra handling. I avoid `ENCODEURL` for this workflow because it is not available in Excel for Mac.

## Search label formula

Paste into **G2** and fill down:

```excel
=IF(F2="","","Search LinkedIn")
```

This formula displays text, **not a clickable hyperlink**. When I select a cell in column G, the separate VBA `Worksheet_SelectionChange` event runs and calls the actual search code.

## Formula testing example

If my row contains `Sarah`, a blank middle name, `Ahmed`, `Example Company`, and `AI Director`, I expect the search to include `Sarah Ahmed Example Company AI Director`.

For troubleshooting, I can use this temporary formula to inspect the joined text:

```excel
=TEXTJOIN(" ",TRUE,A2:E2)
```

**Important:** If my real Excel columns differ from this example, I need to update the formula ranges **and** the VBA column references. I should not blindly copy this setup into a different workbook.

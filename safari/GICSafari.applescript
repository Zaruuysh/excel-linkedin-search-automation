on reuseSearch(theURL)

    tell application "Safari"
        activate

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

        if (count of windows) is 0 then
            make new document with properties {URL:theURL}
        else
            tell front window
                set current tab to (make new tab with properties {URL:theURL})
            end tell
        end if

    end tell

    return "OK"

end reuseSearch

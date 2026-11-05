$json = Get-Content 'C:\Users\norbi\.gemini\PROJECTS\uniqa-collection-retention-lead-prep\data\kozlony_feed.json' -Raw -Encoding UTF8 | ConvertFrom-Json
Write-Output "Total items in feed JSON: $($json.Count)"
Write-Output "Newest item: $($json[0].Title) | $($json[0].PubDate)"
Write-Output "Oldest item: $($json[-1].Title) | $($json[-1].PubDate)"

# nippon_india_scraper_using_excel_button

#VBA code :


Sub RunScraper()

    Dim startTime As String
    Dim endTime As String
    Dim crawlGap As String
    Dim ws As Worksheet
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Get the values from cells C3, C4, and C5
    startTime = ws.Range("C9").Value
    endTime = ws.Range("C10").Value
    crawlGap = ws.Range("C11").Value

    ' Ensure inputs are valid before calling the Python script
    If startTime = "" Or endTime = "" Or crawlGap = "" Then
        MsgBox "Please provide valid start time, end time, and crawl gap.", vbExclamation
        Exit Sub
    End If
    
    ' Set scraping flag to True when starting
    isScraping = True

    ' Call the Python scraper with parameters from the sheet
    RunPython "import nippon_scraper_loop; nippon_scraper_loop.run_scraper('" & startTime & "', '" & endTime & "', " & crawlGap & ")"

End Sub

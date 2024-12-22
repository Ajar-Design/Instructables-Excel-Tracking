VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    Dim IE As Object
    Dim total_num_instructables, total_favs As Integer
    Dim channel_url, channel_name, channel_date As String
    Dim i, p As Integer
    
    'Activate Data Sheet
    Worksheets("Data").Activate

    ' Create InternetExplorer Object
    Set IE = CreateObject("InternetExplorer.Application")

    ' You can uncoment Next line To see form results
    IE.Visible = False
    
    ' URL to get data from
    channel_url = Cells(3, 2).Value
        
    'Check channel URL is to Instructables tab and change if not
    If channel_url Like "/instructables" Then
        'No change required
        Else
        'Add correct tracking
        channel_url = channel_url & "/instructables/"
    End If

    IE.Navigate channel_url

    ' Wait while IE loading and refresh page
    Do While IE.Busy
    Loop
    IE.Refresh
    Do While IE.Busy
    Loop
    
    total_num_instructables = IE.Document.getElementsByClassName("member-stat-numeral ible-count")(0).innertext

    Dim favourites()
    ReDim favourites(total_num_instructables)

    
    'Finding Channel Name and Creation Date
    channel_name = IE.Document.getElementsByClassName("profile-title")(0).innertext
    channel_date = IE.Document.getElementsByClassName("stat-text member-signup-date")(0).innertext
    
    'Finding Individual Instructables Favourites Statistics
    Set favs = IE.Document.getElementsByClassName("ible-favorites")
    Set test_table = IE.Document.getElementsByClassName("ible-stats-right-col pull-right")
    i = 0
    p = 0
    For i = 0 To total_num_instructables - 1
        'Looking for space between two numbers (two numbers if > 0 favs)
        If InStr(2, test_table(i).innertext, " ", 1) = 0 Then
            favourites(i) = 0
        Else
        'Can read favs straight from above
            favourites(i) = favs(p).innertext
            p = p + 1
        End If
    Next
        
 'Scraping titles, accurate views, publish dates and more
    Dim li, tabsClass As Object
    Dim str_li, arrSplitString() As String
    Dim arrFeatured(), arrViews(), arrDates(), arrTitles
    ReDim arrFeatured(total_num_instructables), arrViews(total_num_instructables), arrDates(total_num_instructables), arrTitles(total_num_instructables)
    
    i = 0
    For Each tabsClass In IE.Document.getElementsByClassName("member-ibles member-published-ibles ible-list-items list-unstyled unstyled clearfix")

        'Loop each LI in class
        For Each li In tabsClass.getElementsByTagName("li")
            str_li = li.outerHTML
            arrSplitString = Split(str_li, Chr(34))
            arrFeatured(i) = arrSplitString(1)
            arrDates(i) = arrSplitString(3)
            arrViews(i) = (arrSplitString(5))
            arrTitles(i) = (arrSplitString(15))
            i = i + 1
        Next li
    Next tabsClass
    
    Dim daily_views(), monthly_favs()
    ReDim daily_views(total_num_instructables), monthly_favs(total_num_instructables)
    
    
    'Clean up
    IE.Quit
    Set IE = Nothing
    
       'Formatting and Setup
    total_favs = 0
    
    'Individual Stats'
    For i = 6 To total_num_instructables + 6
        Cells(i, 4).Value = favourites(i - 6)
        total_favs = total_favs + favourites(i - 6)
    Next
    
    'Total Favourites
    Cells(6, 5).Value = total_favs
    
    'Reformat Channel Date
    channel_date = Replace(channel_date, "Joined ", "")
    channel_date = Replace(channel_date, "nd", "")
    channel_date = Replace(channel_date, "rd", "")
    channel_date = Replace(channel_date, "st", "")
    channel_date = Replace(channel_date, "th", "")
    
    
    
    'Individual Instructables Details
    For i = 0 To total_num_instructables - 1
        Cells(6 + i, 9) = arrViews(i)
        Cells(6 + i, 10) = arrDates(i)
        Cells(6 + i, 11) = arrFeatured(i)
        Cells(6 + i, 12) = arrTitles(i)
        
        daily_views(i) = Round(arrViews(i) / Cells(6 + i, 14).Value, 2)
        monthly_favs(i) = Round(favourites(i) * 30 / Cells(6 + i, 14).Value, 2)
    Next
    
    
    
    'Display Data
    '---------------------------------------------------------------------------------------------
    Worksheets("Display Data").Activate
    
    'Individual Instructables Details
    For i = 0 To total_num_instructables - 1
        Cells(19 + i, 2) = arrTitles(i)
        Cells(19 + i, 5) = arrViews(i)
        Cells(19 + i, 6) = favourites(i)
        Cells(19 + i, 8) = arrTitles(i)
        
        Cells(19 + i, 11) = daily_views(i)
        Cells(19 + i, 12) = monthly_favs(i)
        
        
    Next
    
    'Channel Details
    ActiveWorkbook.Sheets("Data").Range("F6").Value = channel_name
    ActiveWorkbook.Sheets("Data").Range("G6").Value = channel_date
    
    'Titles formatting
    ActiveWorkbook.Sheets("Display Data").Range("B2").Value = channel_name & "'s Instructables Stats"
    ActiveWorkbook.Sheets("Display Data").Range("B2:F2").Merge
    

    'Display Charts
    '---------------------------------------------------------------------------------------------
    'Clear old charts
    If ActiveSheet.ChartObjects.Count > 0 Then
        ActiveSheet.ChartObjects.Delete
    End If
    
    Dim pos_left, pos_top, pos_bot, pos_right  As Long
    pos_left = Range("N4").Left
    pos_top = Range("N4").Top
    pos_bot = Range("S25").Top
    pos_right = Range("S4").Left
    
    Debug.Print pos_left, pos_top, pos_bot, pos_right
    
    'Add Views Chart
    Dim piechart As ChartObject
    Dim ser As Series
    Set piechart = ActiveSheet.ChartObjects.Add(pos_left, pos_top, pos_right - pos_left, pos_bot - pos_top)
    With piechart.chart
    .ChartType = xlPie
    .HasTitle = True
    .ChartTitle.Text = "Views Distribution"
    .SetSourceData Source:=Worksheets("Data").Range("O" & 6 & ":P" & 6 + total_num_instructables - 1)
   End With
   Set ser = piechart.chart.SeriesCollection(1)
    With ser.Format
    .Line.Visible = msoTrue
    .Line.ForeColor.RGB = RGB(166, 201, 236)
    End With
    
    'Add Favs Chart
    pos_left = Range("N4").Left
    pos_top = Range("N26").Top
    pos_bot = Range("S" & 25 + 21).Top
    pos_right = Range("S4").Left
    
    Dim piechart2 As ChartObject
    Dim ser2 As Series
    Set piechart2 = ActiveSheet.ChartObjects.Add(pos_left, pos_top, pos_right - pos_left, 340)
    With piechart2.chart
    .ChartType = xlPie
    .HasTitle = True
    .ChartTitle.Text = "Favourites Distribution"
    .SetSourceData Source:=Worksheets("Data").Range("L" & 6 & ":M" & 6 + total_num_instructables - 1)
   End With
   Set ser2 = piechart2.chart.SeriesCollection(1)
    With ser2.Format
    .Line.Visible = msoTrue
    .Line.ForeColor.RGB = RGB(166, 201, 236)
    End With
 End Sub



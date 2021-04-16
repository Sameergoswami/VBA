Attribute VB_Name = "Module1"
Sub Final_1()

    Dim String1 As String
    Dim String2 As String
    Dim String3 As String
    Dim String4 As String
    Dim String5 As String
    Dim String6 As String
    Dim String7 As String
    Dim String8 As String
    Dim String9 As String
    Dim String10 As String
    Dim lastRow As Long
    Dim activeRow As Long
    Dim uniqueID As String
    Dim SummaryRow As Integer
    Dim yr_open As Long
    Dim subRow As Long
    Dim subTotal As Double
    Dim subColumn As Long
    Dim i As Integer
    Dim j As String
    Dim MaxVal As Double
    Dim MinVal As Double
    Dim MaxVol As Double
    Dim Tag, Tag_1, Tag_2, Tag_3, Tag_3a, Tag_4, Tag_5, Tag_6, Tag_7, Tag_8, Tag_8a, Tag_9, Tag_10, Tag_11, Tag12, Tag13, Tag_14, Tag_14a As Double
    
    
    String1 = "Stock Ticker Name "
    String2 = "Yr Open"
    String3 = "Yr Close"
    String4 = "Yr Change"
    String5 = "%Change"
    String6 = "Total Volume"
    String7 = "Greatest % Increase"
    String8 = "Greatest % Decrease"
    String9 = "Largest Volume Stock"
    MaxVal = 0
    MinVal = 0
    MaxVol = 0
    
    'Create New Sheet with name "Summary and Columns with names Stock Ticker Name Yr Open Yr Close Change % Change Total Volume
    
    Sheets.Add.Name = "Summary"
    Set Summary = Worksheets("Summary")
    ActiveWindow.DisplayGridlines = False
    
    Cells(2, 2).Value = String1
    Cells(2, 3).Value = String2
    Cells(2, 4).Value = String3
    Cells(2, 5).Value = String4
    Cells(2, 6).Value = String5
    Cells(2, 7).Value = String6
    Cells(3, 1).Value = String7
    Cells(4, 1).Value = String8
    Cells(5, 1).Value = String9
   
    
    Summary.Columns("A:G").AutoFit 'H
    Summary.Range("B2:G2").Font.Bold = True
    Summary.Range("B2:G2").Font.ColorIndex = 2
    Summary.Range("B2:G2").Interior.ColorIndex = 18
    Summary.Range("A3:A5").Font.Bold = True
    Summary.Range("A3:A5").Font.ColorIndex = 2
    Summary.Range("A3:A5").Interior.ColorIndex = 18
    
    For Each ws In Worksheets

        If ws.Name <> "Summary" Then
        
            ws.Activate
            With ActiveWindow
                .SplitColumn = 0
                .SplitRow = 1
            End With
            ActiveWindow.FreezePanes = True
            ActiveWindow.DisplayGridlines = False
    
            ws.Cells(1, 9).Value = String1
            ws.Cells(1, 10).Value = String2
            ws.Cells(1, 11).Value = String3
            ws.Cells(1, 12).Value = String4
            ws.Cells(1, 13).Value = String5
            ws.Cells(1, 14).Value = String6
            ws.Columns("I:N").AutoFit
            ws.Columns("I:N").Font.Bold = True
          
            lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'Checking for unique id and calculate volume
            subRow = 2
            SummaryRow = 2
            yr_open = 2
            subTotal = ws.Cells(subRow, 7).Value
        
            For subRow = 2 To lastRow
    
                uniqueID = ws.Cells(subRow, 1).Value
    
                If ws.Cells(subRow + 1, 1).Value = uniqueID Then
                    subTotal = subTotal + ws.Cells(subRow + 1, 7).Value
    
                ElseIf ws.Cells(subRow + 1, 1).Value <> uniqueID And ws.Cells(yr_open, 6).Value <> 0 Then
            
                    ws.Cells(SummaryRow, 9).Value = uniqueID
                    ws.Cells(SummaryRow, 10).Value = ws.Cells(yr_open, 6).Value
                    ws.Cells(SummaryRow, 11).Value = ws.Cells(subRow, 6).Value
                    ws.Cells(SummaryRow, 12).Value = ws.Cells(subRow, 6).Value - ws.Cells(yr_open, 6).Value
                
                    If ws.Cells(SummaryRow, 12).Value > 0 Then
                        ws.Cells(SummaryRow, 12).Interior.ColorIndex = 4
                    Else: ws.Cells(SummaryRow, 12).Interior.ColorIndex = 3
                    End If
                
                    ws.Cells(SummaryRow, 13).Value = ((ws.Cells(subRow, 6).Value - ws.Cells(yr_open, 6).Value) / ws.Cells(yr_open, 6).Value)
                    ws.Cells(SummaryRow, 13).NumberFormat = "0.00%"
                    ws.Cells(SummaryRow, 14).Value = subTotal
                    ws.Cells(SummaryRow, 14).NumberFormat = "#,###"
            
                    'Setting varibles for next ticker
                    subTotal = 0
                    SummaryRow = SummaryRow + 1
                    yr_open = subRow + 1
            
                End If
    
            Next subRow
        
         'Calculate Max % Change, Min % Change and Max volume
            
            subRow = 2
            
            Do While ws.Cells(subRow, 13).Value <> ""
                With ws.Cells(subRow, 13)
                    If .Value > MaxVal Then
                        MaxVal = .Value
                        Tag = .Offset(0, -4).Value 'Tkr Name
                        Tag_1 = .Offset(0, -3).Value 'Yr Open
                        Tag_2 = .Offset(0, -2).Value 'Yr Close
                        Tag_3 = .Offset(0, -1).Value 'Yr Change
                        Tag_3a = .Value
                        Tag_4 = .Offset(0, 1).Value 'Total Vol

                    End If
            
                    If .Value < MinVal Then
                        MinVal = .Value
                        Tag_5 = .Offset(0, -4).Value
                        Tag_6 = .Offset(0, -3).Value
                        Tag_7 = .Offset(0, -2).Value
                        Tag_8 = .Offset(0, -1).Value
                        Tag_8a = .Value
                        Tag_9 = .Offset(0, 1).Value
                       
                    End If
                End With
                
                With ws.Cells(subRow, 14)
                    If .Value > MaxVol Then
                        MaxVol = .Value
                        Tag_10 = .Offset(0, -5).Value
                        Tag_11 = .Offset(0, -4).Value
                        Tag_12 = .Offset(0, -3).Value
                        Tag_13 = .Offset(0, -2).Value
                        Tag_14 = .Offset(0, -1).Value
                        Tag_14a = .Value
                      
                        
                    End If
                    
                End With
                subRow = subRow + 1
            Loop
                    
       
        End If
    
    Next ws
     
     'MsgBox ("Stock " + Tag + Str(MaxVal))
    
    'Printing max%
    Summary.Cells(3, 2) = Tag
    Summary.Cells(3, 3) = Tag_1
    Summary.Cells(3, 4) = Tag_2
    Summary.Cells(3, 5) = Tag_3
    Summary.Cells(3, 6) = Tag_3a
    Summary.Cells(3, 6).NumberFormat = "0.00%"
    Summary.Cells(3, 6).Interior.ColorIndex = 4
    Summary.Cells(3, 7) = Tag_4
    Summary.Cells(3, 7).NumberFormat = "#,###"
    'Summary.Cells(3, 8) = Tag15
    
    'Printing min%
    Summary.Cells(4, 2) = Tag_5
    Summary.Cells(4, 3) = Tag_6
    Summary.Cells(4, 4) = Tag_7
    Summary.Cells(4, 5) = Tag_8
    Summary.Cells(4, 6) = Tag_8a
    Summary.Cells(4, 6).NumberFormat = "0.00%"
    Summary.Cells(4, 6).Interior.ColorIndex = 3
    Summary.Cells(4, 7) = Tag_9
    Summary.Cells(4, 7).NumberFormat = "#,###"
    'Summary.Cells(4, 8) = Tag16
     
    'Printing MaxVol
    Summary.Cells(5, 2) = Tag_10
    Summary.Cells(5, 3) = Tag_11
    Summary.Cells(5, 4) = Tag_12
    Summary.Cells(5, 5) = Tag_13
    Summary.Cells(5, 6) = Tag_14
    Summary.Cells(5, 6).NumberFormat = "0.00%"
    Summary.Cells(5, 7) = Tag_14a
    Summary.Cells(5, 7).NumberFormat = "#,###"
    Summary.Cells(5, 7).Interior.ColorIndex = 4
    'Summary.Cells(5, 8) = Tag17

End Sub

Attribute VB_Name = "Module1"
'Created by David Pasquale
'Last edit on 6/30/16
'if website was https certified, links wouldn't work.
'Added request for http(s)://www.site.com

Sub Test()
    firstRow = 6
    Charlimit = 150
    Dim ws As Worksheet, ws2 As Worksheet
    Set ws = Sheets("Content Map - DDC")
    Set ws2 = Sheets("DDC")
    
    'Edit Protected Worksheet
    'ws.Protect Password:="Gabriel1", _
    'UserInterFaceOnly:=True
    
    'Last Column # in Content Map
    With ws
        LastColMap = .Cells(1, .Columns.Count).End(xlToLeft).Column
    End With
    
    'Last Column # in DDCtest
    With ws2
        LastColDDC = .Cells(1, .Columns.Count).End(xlToLeft).Column
    End With

    'Last Row # in DDCtest
    With ws2
        LastRowDDC = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    End With
        
    
    'Content Map loop
    For i = 1 To LastColMap
        searchItem = ws.Cells(1, i).Value
        'Sheets loop
        For j = 1 To LastColDDC
            foundItem = ws2.Cells(1, j).Value
            If searchItem = foundItem Then
            'copy all data in col to Content Map
                For x = 1 To LastRowDDC
                    ws.Cells(x + firstRow - 1, i).Value = ws2.Cells(x + 1, j)
                    'limit cell length
                    If Len(ws.Cells(x + firstRow - 1, i).Value) > Charlimit Then
                        ws.Cells(x + firstRow - 1, i).Value = Left(ws.Cells(x + firstRow - 1, i).Value, Charlimit)
                    End If
                Next x
            Else
            'Do nothing
            End If
        Next j
    Next i
    
    'Points Loop
    For i = firstRow To LastRowDDC + 5
        Path = ws.Cells(i, "K").Value
        If IsEmpty(Path) Then
            'Do Nothing
        Else
            Points = Right(Path, 3)
            If Points = "- 5" Then
                'Check if field has more than one point value
                Result = InStr(Path, ",")
                'If not zero
                If Result <> 0 Then
                    ws.Cells(i, "K").Value = "#ERROR"
                Else
                    ws.Cells(i, "K").Value = 5
                    ws.Cells(i, "BH").Value = "X"
                End If
            
            ElseIf Points = 10 Then
                Result = InStr(Path, ",")
                If Result <> 0 Then
                    ws.Cells(i, "K").Value = "#ERROR"
                Else
                    ws.Cells(i, "K").Value = 10
                    ws.Cells(i, "BG").Value = "X"
                End If
                  
            ElseIf Points = 20 Then
                Result = InStr(Path, ",")
                If Result <> 0 Then
                    ws.Cells(i, "K").Value = "#ERROR"
                Else
                    ws.Cells(i, "K").Value = 20
                    ws.Cells(i, "BI").Value = "X"
                End If
                
            ElseIf Points = 50 Then
                Result = InStr(Path, ",")
                If Result <> 0 Then
                    ws.Cells(i, "K").Value = "#ERROR"
                Else
                    ws.Cells(i, "K").Value = 50
                    ws.Cells(i, "BF").Value = "X"
                End If
                
            ElseIf Points = 100 Then
                Result = InStr(Path, ",")
                If Result <> 0 Then
                    ws.Cells(i, "K").Value = "#ERROR"
                Else
                    ws.Cells(i, "K").Value = 100
                    ws.Cells(i, "BE").Value = "X"
                End If
            Else
                'Do Nothing
            End If
        End If
    Next i

    Columns("BE:BI").Font.Bold = True
    
    'Link loop
    website = ws.Range("P4").Value
    For i = firstRow To LastRowDDC + 5
        Path = ws.Cells(i, "P").Value
        If IsEmpty(Path) Then
            'Do Nothing
        ElseIf ws2.Cells(i - 4, "J").Value = "ddc-landing-page-pro" Then
            concat = website & "lp2/" & Path
            ws.Cells(i, "P").Value = concat
            ws.Hyperlinks.Add Anchor:=ws.Cells(i, "P"), Address:=concat, _
            TextToDisplay:=concat
        Else
            If ws2.Cells(i - 4, "J").Value = "ddc-landing-page" Then
                concat = website & "lp/" & Path
                ws.Cells(i, "P").Value = concat
                ws.Hyperlinks.Add Anchor:=ws.Cells(i, "P"), Address:=concat, _
                TextToDisplay:=concat
            End If
        End If
    Next i
    
    'Left align content
    ws.Cells.HorizontalAlignment = xlLeft
    Columns("BE:BI").HorizontalAlignment = xlCenter
End Sub



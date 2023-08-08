Attribute VB_Name = "Main"
Sub Update_Projection_Table()
    On Error GoTo 0

    'EACH PROJECT
    Dim vRow As Range 'for Row Value
    Dim vRowIndex As Long
    Dim vTable As ListObject
    
    Set vTable = Worksheets("Projects").ListObjects("Projects_Table")
    
    'EACH DATE
    Dim dRow As Range 'for Row Value
    Dim dRowIndex As Long
    Dim dTable As ListObject
    
    Set dTable = Worksheets("Projection").ListObjects("Projection_Table")

    Current_Year = Year(Date)
    Current_Month = Month(Date)
    'Debug.Print Current_Year & "-" & Current_Month
    
    'CREATE MONTHS
    On Error Resume Next
    With Sheets("Projection")
        .Rows(2 & ":" & .Rows.Count).Delete
    End With
    
    On Error GoTo 0
    
    
    latestDate = WorksheetFunction.Max(vTable.ListColumns("Finish Date").Range)
    
    
    
    
    totalMonths = (Year(latestDate) - Year(Date)) * 12 + (Month(latestDate) - Month(Date) + 1)
    
    'Debug.Print totalMonths
    



    
    vRowIndex = 0
    For Each vRow In vTable.ListColumns("Project Name").DataBodyRange.Rows
        vRowIndex = vRowIndex + 1
        ' Use vRow if you only need the value from that column
        ' Use comment below for different row values based on header name
        cProject_Name = vTable.DataBodyRange.Cells(vRowIndex, vTable.ListColumns("Project Name").Index)
        cProject_Client = vTable.DataBodyRange.Cells(vRowIndex, vTable.ListColumns("Client").Index)
        cProject_Start = vTable.DataBodyRange.Cells(vRowIndex, vTable.ListColumns("Start Date").Index)
        cProject_Finish = vTable.DataBodyRange.Cells(vRowIndex, vTable.ListColumns("Finish Date").Index)
        cProject_Value = vTable.DataBodyRange.Cells(vRowIndex, vTable.ListColumns("Projected Value").Index)
        cProject_Billed = vTable.DataBodyRange.Cells(vRowIndex, vTable.ListColumns("Billed To Date").Index)
        
        
        'If is Project Month then add month line.....
        

        cProject_RemainingRev = cProject_Value - cProject_Billed
        
        If Date > cProject_Start Then
            cProject_RemainingMonths = (Year(cProject_Finish) - Year(Date)) * 12 + (Month(cProject_Finish) - Month(Date) + 1)
            newRow_Date = Date
        Else
            cProject_RemainingMonths = (Year(cProject_Finish) - Year(cProject_Start)) * 12 + (Month(cProject_Finish) - Month(cProject_Start) + 1)
            newRow_Date = cProject_Start
        End If
        
        
        cProject_RemainingMonthlyRev = cProject_RemainingRev / cProject_RemainingMonths
        
        
        
        'newRow_Date = cProject_Start
        For rm = 1 To cProject_RemainingMonths
            newRow_Date = DateAdd("m", 1, newRow_Date)
            
            Set newRow = dTable.ListRows.Add()
            With newRow
                .Range(1) = newRow_Date
                .Range(2) = newRow_Date
                .Range(3) = cProject_RemainingMonthlyRev
                .Range(4) = cProject_Name
            End With
        Next rm
        


        
        
        
        
    Next vRow
    ''''''''END vLoop
        
ActiveWorkbook.RefreshAll
Exit Sub
errHandle:
        


End Sub

Sub test()
    dRowIndex = 0
    For Each dRow In dTable.ListColumns("Project Name").DataBodyRange.Rows
        dRowIndex = dRowIndex + 1
        ' Use vRow if you only need the value from that column
        ' Use comment below for different row values based on header name
        cpj_Year = dTable.DataBodyRange.Cells(dRowIndex, dTable.ListColumns("Year").Index)
        cpj_Month = dTable.DataBodyRange.Cells(dRowIndex, dTable.ListColumns("Month").Index)
        cpj_ProjectedRev = dTable.DataBodyRange.Cells(dRowIndex, dTable.ListColumns("Projected Rev").Index)
        cpj_Name = dTable.DataBodyRange.Cells(dRowIndex, dTable.ListColumns("Project Name").Index)
    Next dRow
    ''''''''END dLoop
End Sub



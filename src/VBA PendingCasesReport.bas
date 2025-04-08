Public Sub createAllPendingCasesReport()
    Dim myDicTypes As Object
    Dim theDate As String
    Dim criteria1 As String
    Dim criteria2 As String
    Dim CaseTypes As Range
    Dim ctrlText As String
    Dim key As String
    
    On Error GoTo messageError
    
    Call UnfiltermyTable
    
    Set CaseTypes = myActiveSheet.Range("CaseTypes")
    
    'Populating the dictionary with the values in the CaseType Column, the 2nd for the names of the tables in the document
    Set myDicTypes = CreateObject("Scripting.Dictionary")
      For t = 1 To CasesTypes.Rows.Count
        key = CaseTypes(t, 1).Value
        myDicTypes.Add key, key
      Next t
      
    'Filtering Column Current Status  of myTable excluding cases which status is "Finisihed and issued"
      criteria1 = "<>" & checkBoxesDictionary("CheckBox_CurrentStatusFinishedAndIssued")
      Call filterFieldsByCriteria(ColumnCurrrentStatus, criteria1)
      
      theDate = InputBox("Enter the date until which you want the report to be issued")
        If theDate <> "" Then
          Dim longDate As Long
          longDate = DateSerial(year(theDate), Month(theDate), Day(theDate))
          criteria1 = "<=" & longDate
          Call filterFieldsByCriteria(myTable.ListColumns(theColumnIndex), criteria1)
        End If
        
    'Populating a dictionary that will hold all the Case Types that are still shown after the table is filtered
    Set myDicExisting = CreateObject("Scripting.Dictionary")
     Dim vCells As Range
     Set vCells = myTable.DataBodyRange.SpecialCells(xlCellTypeVisible)
      For A = 1 To vCells.Areas.Count
        For r = 1 To vCells.Areas(A).Rows.Count
          key = vCells.Areas(A).Rows(r).Cells(ColumnCaseType.index).Value
            If Not (myDicExisting.exists(key)) And key <> "" Then
              myDicExisting.Add key, key
            End If
        Next r
      Next A
    
    
    If theColumnIndex = ColumnReceiptionDate.index Then
      Call generateAllPendingCasesDoc("_PendingCasesReport_Arranged by Case Date", myDicTypes, theDate, ColumnJudiciaryCaseYear, ColumnCaseYear, ColumnCaseNumber)
    End If
    
    Call generateAllPendingCasesDoc("_PendingCasesReport_Arranged by Register Number", myDicTypes, theDate, ColumnReceiptDate, ColumnReceiptNumber, ColumnRegisterNumber)
    
    If theColumnIndex = ColumnReceiptionDate.index Then
      MsgBox "Finished issuing the 2 sets of reports successfuly"
    Else
      MsgBox "Finished issuing the report successfuly"
      theColumnIndex = ColumnReceiptionDate.index
    End If
    
    
    Exit Sub
messageError:     Call onErrorMessage(ColumnCaseType)

End Sub

Private Sub generateAllPendingCasesDoc(docName As String, _
                                      myDicTypes As Object, _
                                      theDate As Variant, _
                                      Column1 As ListColumn, _
                                      Column2 As ListColumn, _
                                      Optional Column3 As ListColumn)
  Dim wd As Object
  Dim AllPendingCasesReportDoc As Word.Document
  Dim AllPendingCasesTemplatePath As Variant
  Dim myTblsCollection As New Collection
  
  'We select the template
  AllPendingCasesTemplatePath = PickFileWhenPrompted("Choose the template for the report")
    Set wd = CreateObject("Word.Application")
    If wd.Visible = False Then
      wd.Visible = True
    End If
    
        'We create a new document from the tempalte
      Set AllPendingCasesReportDoc = wd.Documents.Add(AllPendingCasesTemplatePath, , wdWordDocument)
          
      'We save the file with the appropriate file name
      myDate = Format(Date, "yyyy.mm.dd") '& "." & Format(Now(), "hh") & "." & Format(Now(), "mm")
      fileName = "D:\" & myDate & docName & ".docx"
      AllPendingCasesReportDoc.SaveAs2 (fileName)
        
          
  
  'Populating a collection of all the tables
  With AllPendingCasesReportDoc
    For t = 1 To .Tables.Count
      myTblsCollection.Add .Tables(t)
    Next
  End With
  
  
    'Civil Appealed
      Call prepareAllPendingCasesReport(myDicTypes.items()(0), myTblsCollection(1), Column1, Column2, Column3)
    
    'Administrative
      Call prepareAllPendingCasesReport(myDicTypes.items()(1), myTblsCollection(1), Column1, Column2, Column3)
    
    'Civil
      Call prepareAllPendingCasesReport(myDicTypes.items()(2), myTblsCollection(1), Column1, Column2, Column3)
    
    'Labour Appealed
    
      Call prepareAllPendingCasesReport(myDicTypes.items()(3), myTblsCollection(2), Column1, Column2, Column3)
    
    'Labour
      Call prepareAllPendingCasesReport(myDicTypes.items()(4), myTblsCollection(2), Column1, Column2, Column3)
    
    'Tax, Persons, Criminal, Public Funds
      For t = 5 To myDicTypes.Count - 1
          Call prepareAllPendingCasesReport(myDicTypes.items()(t), myTblsCollection(t - 2), Column1, Column2, Column3)
      Next t
  
    Dim ctrl As ContentControl
    
    Call deleteEmptyContentControls(AllPendingCasesReportDoc, theDate)
  
    AllPendingCasesReportDoc.Range.Font.TextColor = wdColorAutomatic
    AllPendingCasesReportDoc.Save
  
End Sub

Private Sub prepareAllPendingCasesReport(criteria As Variant, _
                                        tbl As Object, _
                                        Column1 As ListColumn, _
                                        Column2 As ListColumn, _
                                        Optional Column3 As ListColumn)
  Dim AllPendingCasesReportTable As Word.Table
  
      If myDicExisting.exists(criteria) Then
        criteria1 = "=" & criteria
        Call filterFieldsByCriteria(ColumnCaseType, criteria1)
        Call sortMyTable(Column1, xlAscending, Column2, xlAscending, Column3, xlAscending)
        Call processAllPendingCasesReport(tbl)
      End If
End Sub

Private Function processAllPendingCasesReport(tbl As Word.Table)
  Dim visibleCells As Range
  Dim AllPendingCasesReportDoc As Word.Document
  Dim AllPendingCasesReportTable As Word.Table
  Dim myArray() As Variant
  Dim ctrlText As String
  Dim ReceiptNumber As String
  Dim CaseYear As String
  Dim securityChecker As String
  Dim entryChecker As String
  Dim r As Integer
  Dim w As Integer
  
  
  'Using the visible cells after filtering the table in order to fill the report
  Set visibleCells = myTable.DataBodyRange.SpecialCells(xlCellTypeVisible)
  
  
  'Populating the date RichText field in the Report
  
  
  counter = 1 'this counter counts the rows in the visible celles of the table
  r = 3
  
  ReDim myArray(myTable.ListColumns.Count, 1) 'Redming the array as a multidimensional array with rows = the number of columns in the table
  
  'populating an arry representing all the visible cells excluding redundant values of receiptnumber
  For i = 1 To visibleCells.Areas.Count
    For n = 1 To visibleCells.Areas(i).Rows.Count
      'In order to avoid repetition,we check if the new row is for a case with the same claimant before populating the array
            With visibleCells.Areas(i).Rows(n)
             entryChecker = .Cells(ColumnReceiptNumber.index).Value & .Cells(ColumnReceiptDate.index).Value & .Cells(ColumnRegisterNumber.index).Value
            End With
        
        If entryChecker <> securityChecker Then
          ReDim Preserve myArray(UBound(myArray, 1), UBound(myArray, 2) + 1)
            For c = 1 To myTable.ListColumns.Count
              myArray(c, UBound(myArray, 2) - 1) = visibleCells.Areas(i).Rows(n).Cells(myTable.ListColumns(c).index).Value
              
              With visibleCells.Areas(i).Rows(n)
               securityChecker = .Cells(ColumnReceiptNumber.index).Value & .Cells(ColumnReceiptDate.index).Value & .Cells(ColumnRegisterNumber.index).Value
              End With
            Next c
        End If
    Next n
  Next i
  
    For col = 1 To UBound(myArray, 2) - 1
    
        tbl.cell(tbl.Rows.Count, 8).Range.Rows.Add
                                                                                   
        'Serial Number
          'Call insertContentControlInTable(tbl, 1, wdContentControlRichText, "RT_SerialNumber", tbl.Rows.Count - r)
          Call fillTable(tbl, 1, tbl.Rows.Count - r)
  
          
        'Receipt Number & Receipt Date
          If myArray(ColumnReceiptNumber.index, col) <> "" Then
            ReceiptNumber = myArray(ColumnReceiptNumber.index, col)
          ElseIf myArray(ColumnReceiptNumber.index, col) = "" Then
            ReceiptNumber = myArray(ColumnRegisterNumber.index, col)
          End If
          If theColumnIndex = ColumnReceiptDate.index Then
            ctrlText = ReceiptNumber & vbNewLine & Format(myArray(ColumnReceiptDate.index, col), "dd/MM/yyyy")
          Else
            ctrlText = ReceiptNumber & "/" & Format(myArray(ColumnReceiptDate.index, col), "yyyy")
          End If
          'Call insertContentControlInTable(tbl, 2, wdContentControlRichText, "RT_ReceiptNumber", ctrlText)
          Call fillTable(tbl, 2, ctrlText)
        
        'Case Number, Case Year And Case Court
          If myArray(ColumnJudiciaryCaseYear.index, col) <> "" Then
            CaseYear = myArray(ColumnJudiciaryCaseYear.index, col)
          ElseIf myArray(ColumnJudiciaryCaseYear.index, col) = "" Then
            CaseYear = myArray(ColumnCaseYear.index, col)
          End If
          ctrlText = myArray(ColumnCaseNumber.index, col) & " " & ColumnCaseYear.name & " " & CaseYear & " " & ColumnCaseCourt.name & " " & myArray(ColumnCaseCourt.index, col)
          'Call insertContentControlInTable(tbl, 3, wdContentControlRichText, "RT_CaseNumber", ctrlText)
          Call fillTable(tbl, 3, ctrlText)
  
   
        'Claimant & Defendant
          ctrlText = myArray(ColumnClaimantName.index, col) & " x " & myArray(ColumnDefendantName.index, col)
          'Call insertContentControlInTable(tbl, 4, wdContentControlRichText, "RT_Parties", ctrlText)
          Call fillTable(tbl, 4, ctrlText)
        
        'Transfer Date
          ctrlText = myArray(ColumnTransferDate.index, col)
          'Call insertContentControlInTable(tbl, 5, wdContentControlRichText, "RT_TransferDate", ctrlText)
          Call fillTable(tbl, 5, ctrlText)
  
        
        'Receiption Date
          ctrlText = myArray(ColumnReceiptionDate.index, col)
          'Call insertContentControlInTable(tbl, 6, wdContentControlRichText, "RT_ReceiptionDate", ctrlText)
          Call fillTable(tbl, 6, ctrlText)
  
        
        'First Meeting Date
          ctrlText = myArray(ColumnFirstMeetingDate.index, col)
          'Call insertContentControlInTable(tbl, 7, wdContentControlRichText, "RT_FirstMeetingDate", ctrlText)
          Call fillTable(tbl, 7, ctrlText)
  
        
        'Current Status
        
          If myArray(ColumnCurrrentStatus.index, col) = checkBoxesDictionary("CheckBox_CurrentStatusPostponed") Then
            ctrlText = ChrW("&H00FC")
            'Call insertContentControlInTable(tbl, 8, wdContentControlRichText, "RT_CurrentStatusPostopned", ctrlText, "Wingdings", 14)
            Call fillTable(tbl, 8, ctrlText, "Wingdings", 14)
  
          
          ElseIf myArray(ColumnCurrrentStatus.index, col) = checkBoxesDictionary("CheckBox_CurrentStatusOngoing") Then
            ctrlText = myArray(ColumnLastMeetingDate.index, col)
            'Call insertContentControlInTable(tbl, 9, wdContentControlDate, "RT_CurrentStatusOngoing", ctrlText, , , , "dd/MM/yyyy")
            Call fillTable(tbl, 9, Format(ctrlText, "dd/MM/yyyy"))
          
          ElseIf myArray(ColumnCurrrentStatus.index, col) = checkBoxesDictionary("CheckBox_CurrentStatusFinishedNotIssued") Then
            ctrlText = myArray(ColumnEndOfExaminationDate.index, col)
            'Call insertContentControlInTable(tbl, 10, wdContentControlDate, "RT_FinishedNotIssued", ctrlText, , , , "dd/MM/yyyy")
            Call fillTable(tbl, 10, Format(ctrlText, "dd/MM/yyyy"))
          End If
        
        'Achievement Type
          ctrlText = myArray(ColumnIfReturnedAchievementType.index, col)
          'Call insertContentControlInTable(tbl, 11, wdContentControlRichText, "RT_AchievementType", ctrlText)
          Call fillTable(tbl, 11, ctrlText)
        
        'Name of the previous expert
          ctrlText = myArray(ColumnIfReturnedNameOfPreviousExpert.index, col)
          'Call insertContentControlInTable(tbl, 12, wdContentControlRichText, "RT_PreviousExpertName", ctrlText)
          Call fillTable(tbl, 12, ctrlText)
        
        'Date the previous expert rendered his report
          ctrlText = myArray(ColumnIfReturnedRegisterNumber.index, col)
          'Call insertContentControlInTable(tbl, 13, wdContentControlDate, "RT_ReportDate", ctrlText)
           Call fillTable(tbl, 13, ctrlText)
          
  
                                            'ctrlText = myArray(ColumnCaseType.index, col) & "/" & myArray(ColumnCurrrentStatus.index, col) & "/" & myArray(ColumnReceiptionDate.index, col)
                                            'Call insertContentControlInTable(tbl, 14, wdContentControlRichText, "Trash", ctrlText)
    Next col
   

End Function

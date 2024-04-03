Attribute VB_Name = "expense_schedule_lookup"
Sub expense_sched_all_depts()

  Application.DisplayStatusBar = True
  
  Dim depts As Variant
  depts = Array("Q:\FPO Business Development\Business Plans\NYP Review\Westchester Avenue\Westchester Avenue Business Plan Template 2.22.2024 - Cardiology.xlsb") ', _
'                "Q:\FPO Business Development\Business Plans\NYP Review\Westchester Avenue\Westchester Avenue Business Plan Template 1.29.2024 - Dermatology.xlsb", _
'                "Q:\FPO Business Development\Business Plans\NYP Review\Westchester Avenue\Westchester Avenue Business Plan Template 1.29.2024 - ENT.xlsb", _
'                "Q:\FPO Business Development\Business Plans\NYP Review\Westchester Avenue\Westchester Avenue Business Plan Template 1.29.2024 - Medical Oncology.xlsb", _
'                "Q:\FPO Business Development\Business Plans\NYP Review\Westchester Avenue\Westchester Avenue Business Plan Template 1.29.2024 - Medicine.xlsb", _
'                "Q:\FPO Business Development\Business Plans\NYP Review\Westchester Avenue\Westchester Avenue Business Plan Template 1.29.2024 - Neurology.xlsb", _
'                "Q:\FPO Business Development\Business Plans\NYP Review\Westchester Avenue\Westchester Avenue Business Plan Template 1.29.2024 - OBGYN.xlsb", _
'                "Q:\FPO Business Development\Business Plans\NYP Review\Westchester Avenue\Westchester Avenue Business Plan Template 1.29.2024 - Ophthalmology.xlsb", _
'                "Q:\FPO Business Development\Business Plans\NYP Review\Westchester Avenue\Westchester Avenue Business Plan Template 1.29.2024 - Ortho.xlsb", _
'                "Q:\FPO Business Development\Business Plans\NYP Review\Westchester Avenue\Westchester Avenue Business Plan Template 1.29.2024 - Pain Management.xlsb", _
'                "Q:\FPO Business Development\Business Plans\NYP Review\Westchester Avenue\Westchester Avenue Business Plan Template 1.29.2024 - Pediatrics.xlsb", _
'                "Q:\FPO Business Development\Business Plans\NYP Review\Westchester Avenue\Westchester Avenue Business Plan Template 1.29.2024 - PM&R.xlsb", _
'                "Q:\FPO Business Development\Business Plans\NYP Review\Westchester Avenue\Westchester Avenue Business Plan Template 1.29.2024 - Primary Care.xlsb", _
'                "Q:\FPO Business Development\Business Plans\NYP Review\Westchester Avenue\Westchester Avenue Business Plan Template 1.29.2024 - Psychiatry.xlsb", _
'                "Q:\FPO Business Development\Business Plans\NYP Review\Westchester Avenue\Westchester Avenue Business Plan Template 1.29.2024 - Surgery.xlsb", _
'                "Q:\FPO Business Development\Business Plans\NYP Review\Westchester Avenue\Westchester Avenue Business Plan Template 1.29.2024 - Urology.xlsb")


  ' status bar variables
  Dim x As Long
  x = 0
  
  Dim y As Long
  y = UBound(depts) - LBound(depts) + 1
  
  For Each i In depts
    Application.StatusBar = Left(Right(Mid(i, InStrRev(i, " - ")), Len(Mid(i, InStrRev(i, " - "))) - 3), _
                              Len(Right(Mid(i, InStrRev(i, " - ")), Len(Mid(i, InStrRev(i, " - "))) - 3)) - 5) & _
                              ": " & Format(x / y, "0%")
                              
    expense_schedule_lookup (i)
    
    x = x + 1
  Next i
  
  Application.DisplayStatusBar = False

End Sub


Sub expense_schedule_lookup(origFilePath)

  ' source data
  Dim sourceFilePath As String
  sourceFilePath = "Q:\FPO Business Development\Business Plans\NYP Review\Westchester Avenue\Comp Model\Test Expense Schedule Macro\"
  
  Dim sourceWb As Workbook
  Set sourceWb = Workbooks.Open(sourceFilePath & "Comp Model Salary Schedules.xlsx")
  
  Dim baseWs As Worksheet
  Set baseWs = sourceWb.Worksheets("Schedule (Non Prorated)-Base")
  
  Dim bonusWs As Worksheet
  Set bonusWs = sourceWb.Worksheets("Schedule (Non Prorated)-Bonus")
  
  
  ' destination workbook
'  Dim origFilePath As String
'  origFilePath = "Q:\FPO Business Development\Business Plans\NYP Review\Westchester Avenue\Comp Model\Test Expense Schedule Macro\"
  
  Dim origWb As Workbook
  Set origWb = Workbooks.Open(origFilePath)
  
  ' save copy of destination workbook with today's date
  Dim dept As String
  dept = Left(Right(Mid(origWb.Name, InStrRev(origWb.Name, " - ")), Len(Mid(origWb.Name, InStrRev(origWb.Name, " - "))) - 3), _
           Len(Right(Mid(origWb.Name, InStrRev(origWb.Name, " - ")), Len(Mid(origWb.Name, InStrRev(origWb.Name, " - "))) - 3)) - 5)
  
  Dim newFileName As String
  newFileName = "Westchester Avenue Business Plan Template " & Format(Date, "m.d.yyyy") & " - " & dept & ".xlsb"
  newFileNameV2 = "Westchester Avenue Business Plan Template " & Format(Date, "m.d.yyyy") & " - " & dept & " v2.xlsb"
  
  Dim destFilePath As String
  destFilePath = "Q:\FPO Business Development\Business Plans\NYP Review\Westchester Avenue\Comp Model\Test Expense Schedule Macro\"
  
  ' if newFileName exists in folder, save as newFileName_v2
  If Dir(destFilePath & newFileName) <> "" Then
    origWb.SaveAs FileName:=destFilePath & newFileNameV2
  Else:
    origWb.SaveAs FileName:=destFilePath & newFileName
  End If
  
  
  ' close origWb
  origWb.Close SaveChanges = False
  
  Dim newWb As Workbook
  If Dir(destFilePath & newFileNameV2) <> "" Then
    Set newWb = Workbooks.Open(destFilePath & newFileNameV2)
  Else:
    Set newWb = Workbooks.Open(destFilePath & newFileName)
  End If
  
  Dim expenseSchedWs As Worksheet
  Set expenseSchedWs = newWb.Worksheets("Expense Schedule")
                      
           
  ' variables for incremental xlookup
  Dim k As Long
  k = 9

  ' xlookup values from base & bonus to expenseSched
  For Each i In expenseSchedWs.Range("A8:A175")
    If Not IsEmpty(i.Value) _
      And Not (i.Value Like "*Unique*") _
      And IsNumeric(i.Value) = False _
      And Len(i) > 5 Then
    For j = 13 To 42
        i.Offset(-1, j).Value = Format(Application.WorksheetFunction.XLookup(i, baseWs.Range("A:A"), baseWs.Columns(k), ""), "###,###")  ' base Year k - 8
        i.Offset(0, j).Value = Format(Application.WorksheetFunction.XLookup(i, bonusWs.Range("A:A"), bonusWs.Columns(k), ""), "###,###") ' bonus Year k - 8

        k = k + 1
    Next j
    End If
    k = 9
  Next i
           
'  newWb.Close SaveChanges = True

End Sub


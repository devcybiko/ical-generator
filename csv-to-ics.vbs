'Input Excel File's Full Path
  ExcelFilePath = "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"

'Input Module/Macro name within the Excel File
  MacroPath = "Module1.PrintRecurring"

'Create an instance of Excel
  Set ExcelApp = CreateObject("Outlook.Application")
  MsgBox ExcelApp

'Do you want this Excel instance to be visible?
'   ExcelApp.Visible = True  'or "False"

'Prevent any App Launch Alerts (ie Update External Links)
'   ExcelApp.DisplayAlerts = False

'Open Excel File
'   Set wb = ExcelApp.Workbooks.Open(ExcelFilePath)

'Execute Macro Code
  ExcelApp.Run MacroPath

'Save Excel File (if applicable)
'   wb.Save

'Reset Display Alerts Before Closing
'   ExcelApp.DisplayAlerts = True

'Close Excel File
'   wb.Close

'End instance of Excel
  ExcelApp.Quit

'Leaves an onscreen message!
  MsgBox "Your Automated Task successfully ran at " & TimeValue(Now), vbInformation


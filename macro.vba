Public Sub StartExeWithArgument()
    Dim strProgramName As String
    Dim strArgument As String

    strProgramName = "C:\Users\GSMIT615\git\ical-generator\csv-to-ics.bat"
    strArgument = ""

    ''' Call Shell("""" & strProgramName & """ """ & strArgument & """", vbNormalFocus)
    Call Shell(strProgramName, vbNormalFocus)
End Sub

Sub ShowCalendar()

    Dim olApp As Outlook.Application
    Dim olNs As NameSpace


    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")


    If Err.Number = 429 Then
        Set olApp = CreateObject("Outlook.Application")
    End If


    On Error GoTo 0


    Set olNs = olApp.GetNamespace("MAPI")


    If olApp.ActiveExplorer Is Nothing Then
        olApp.Explorers.Add _
            (olNs.GetDefaultFolder(olFolderCalendar)).Activate
    Else
        Set olApp.ActiveExplorer.CurrentFolder = _
            olNs.GetDefaultFolder(olFolderCalendar)
        olApp.ActiveExplorer.Display
    End If


    Set olNs = Nothing
    Set olApp = Nothing


End Sub

Sub PrintRecurring()
    Dim CalFolder As Outlook.MAPIFolder
    Dim CalItems As Outlook.Items
    Dim ResItems As Outlook.Items
    Dim sFilter, strSubject, strOccur As String
    Dim itm, ListAppt As Object
    Dim tStart, tEnd As Date
    
    Call ShowCalendar
    
    '''
    ''' constants
    '''
    outfname = "C:\Users\GSMIT615\git\ical-generator\calendar.csv"
    Attachment = "C:\Users\GSMIT615\git\ical-generator\calendar.ics"
    comma = "|"
    
    '''
    ''' query the calendar
    '''
    Set CalFolder = Application.ActiveExplorer.CurrentFolder ' Use the selected calendar folder
    Set CalItems = CalFolder.Items
    CalItems.Sort "[Start]" ' Sort all of the appointments based on the start time
    CalItems.IncludeRecurrences = True
    tStart = Format(Now, "Short Date") ' Set an end date
    tEnd = Format(Now + 14, "Short Date")
    sFilter = "[Start] >= '" & tStart & "' And [End] < '" & tEnd & "' And  [IsRecurring]  = True" 'create the Restrict filter by day and recurrence
    Set ResItems = CalItems.Restrict(sFilter)

    '''
    ''' Loop through the items in the collection.
    '''
    strOccur = "Subject|Start Date|Start Time|End Date|End Time|All Day|Description|Location|UID|Busy Status"
    For Each itm In ResItems
        StartDate = Format(itm.Start, "yyyy-mm-dd")
        EntryID = Mid(itm.EntryID, Len(itm.EntryID) - 16, 16)
        strOccur = strOccur & vbCrLf
        strOccur = strOccur & "FORD:" & itm.Subject
        strOccur = strOccur & comma & Format(itm.Start, "mm/dd/yyyy|hh:mm AM/PM")
        strOccur = strOccur & comma & Format(itm.End, "mm/dd/yyyy|hh:mm AM/PM")
        strOccur = strOccur & comma & "FALSE"
        strOccur = strOccur & comma & itm.Subject
        strOccur = strOccur & comma & "FORD"
        strOccur = strOccur & comma & StartDate & "-" & EntryID
        strOccur = strOccur & comma & "BUSY"
    Next
   
    '''
    ''' write the file
    '''
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set FileToCreate = FSO.CreateTextFile(outfname)
    FileToCreate.Write strOccur
    FileToCreate.Close
    
    '''
    ''' convert to ical / ics
    '''
    Call StartExeWithArgument
    
    '''
    ''' send it as an email
    '''
    SendEmail (Attachment)
  
    '''
    ''' clean up
    '''
    Set itm = Nothing
    Set ListAppt = Nothing
    Set ResItems = Nothing
    Set CalItems = Nothing
    Set CalFolder = Nothing
    
    
    '''
    ''' alert
    '''
    MsgBox "Email sent with attachment " & outfname
End Sub

Sub SendEmail(Attachment)
    Set ListAppt = Application.CreateItem(olMailItem)
    ListAppt.Body = "calendar.ics attached..."
    ListAppt.Subject = "Greg Smith Calendar"
    ListAppt.To = "gregory@greg-smith.com"
    ListAppt.Attachments.Add Attachment
    ListAppt.Send
    'ListAppt.Display
End Sub

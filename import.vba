Option Compare Database

'Use Windows File Browser to select file for input

Private Sub fileBrowse_Click()
    
    Const msoFileDialogFilePicker As Long = 3
    Dim f As Object
    Set f = Application.FileDialog(msoFileDialogFilePicker)
    f.AllowMultiSelect = False
    
    If f.Show Then
        fileP = f.SelectedItems(1)
    End If

End Sub

'Executes when "Import" button is clicked, Imports file to database
'First makes a temporary copy of the database, then checks for input errors and finally transfers over temp data to master table

Private Sub import_Click()
        
    'Check to see if user has entered all the fields
    
    If IsNull(Me.fileP) Then
        MsgBox "Please Enter a File Path!", vbOKOnly
        Exit Sub
    End If
    
    If IsNull(Me.dateS) Then
        MsgBox "Please Enter a Date Submitted!", vbOKOnly
        Exit Sub
    End If
    
    If Me.dateS > Date Then
        'MsgBox "Please enter a valid date (must be today or older).", vbOKOnly
        'Exit Sub
    End If
    
    If IsNull(Me.templateT) Then
        MsgBox "Please Enter a Template Type!", vbOKOnly
        Exit Sub
    End If
    
    transferToDB 'prep table
    wipeErrors
    checkComplete (Me.templateT) 'Check to see BCUS fields are complete
    getSleep 'make sure prev finishes executing
    sanitize 'Check data types
    getSleep 'make sure prev finishes executing
    tableName = importTODB 'import
    
    DoCmd.Close acTable, tableName, acSaveYes
    DoCmd.OpenTable tableName
    
End Sub

'Wipe the contents of the import errors from before

Private Function wipeErrors()

    DoCmd.Close acTable, "Import Errors", acSaveYes
    wipeError = "DELETE * FROM [Import Errors]"
    CurrentDb.Execute wipeError, dbFailOnError

End Function

'Integrate Fraud templates (only) w Fraud master table by transferring over data in an appropriate fashion (see cases)

Private Function importTODBF()
    
    'Set up Recordsets
    Dim db As DAO.Database
    Dim rsSrc As DAO.Recordset
    Dim rsDest As DAO.Recordset
    Dim errorTable As DAO.Recordset
    Dim found As Boolean
    
    'Set db variables
    Set db = CurrentDb
    Set rsSrc = db.OpenRecordset(Me.templateT, dbOpenSnapshot)
    Set rsDest = db.OpenRecordset("Fraud Table", _
        dbOpenTable, dbAppendOnly)
    Set errorTable = db.OpenRecordset("Import Errors", _
        dbOpenTable, dbAppendOnly)
        
    counter = 4
    tableError = False
    
    'Compare each record in DB to record in upload file
    'loop through each record in upload temp table
    With rsSrc
        Do Until rsSrc.EOF
                                   
            'compare to each record in master table
            With rsDest
                
                found = False
                
                'If there are records in the table, start comparing
                If Not DCount("*", "Fraud Table") = 0 Then
                
                    rsDest.MoveFirst
                    
                    Do Until rsDest.EOF
                         
                        rsDest.Edit
                        rowError = False
                        updateRow = False
                         
                        'Record already exists in the database
                        If rsSrc![Transaction ID] = rsDest![Transaction ID] And Me.templateT = rsDest![Request Type] And rsSrc![Cardmember Number] <> "ERROR" Then
                            found = True
                            
                            If rsDest![Request Type] = "Fraud Research" Then
                                
                                'Case 1a: During retrieval request, AX gets back to BCUS w info, add info
                                If Not IsNull(rsSrc![Fees]) Then
                                    
                                    'Add fields if not a dupe
                                    If IsNull(rsDest![Fees]) Then
                                        rsDest![Fees] = rsSrc![Fees]
                                        rsDest![Credits] = rsSrc![Credits]
                                        rsDest![Finance Charges] = rsSrc![Finance Charges]
                                        rsDest![Late Payment] = rsSrc![Late Payment]
                                        rsDest![Re-Tier] = rsSrc![Re-Tier]
                                        rsDest![Re-Age] = rsSrc![Re-Age]
                                        updateRow = True
                                    
                                    'Duplicate record w BCUS and AXP fields
                                    Else
                                        errorMsg = "This record is a duplicate with BCUS and AXP info."
                                        rowError = True
                                    End If
                                
                                'Duplicate record w BCUS fields
                                Else
                                    errorMsg = "This record is a duplicate with BCUS info."
                                    rowError = True
                                End If
                            
                            ElseIf rsDest![Request Type] = "Fraud Chargeback" Then
                            
                                'Case 1b: During chargeback request, AX gets back to BCUS w info, add info
                                If Not IsNull(rsSrc![Chargeback Outcome]) Then
                                    
                                    'Add fields if not dupe
                                    If IsNull(rsDest![Chargeback Outcome]) Then
                                        rsDest![Chargeback Outcome] = rsSrc![Chargeback Outcome]
                                        updateRow = True
                                    
                                    'Dupe w BCUS and AXP fields
                                    Else
                                        errorMsg = "This record is a duplicate with BCUS and AXP info."
                                        rowError = True
                                    End If
                                    
                                'Dupe w BCUS info only
                                Else
                                    errorMsg = "This record is a duplicate with BCUS info."
                                    rowError = True
                                End If
                                                            
                            End If
                            
                        End If
                        
                        'if error then add error to table
                        If rowError Then
                            
                            tableError = True
                            errorTable.AddNew
                            errorTable![row] = counter
                            errorTable![error] = errorMsg
                            errorTable![File Name] = getFileName(Me.fileP)
                            If Me.templateT <> "Fraud Research" And Me.templateT <> "Fraud Chargeback" And Me.templateT <> "Fraud Application" Then
                                errorTable![Cardmember Name] = tableToCheck![Cardmember Name]
                            End If
                            errorTable![Cardmember Number] = rsDest![Cardmember Number]
                            errorTable.Update
                        ElseIf updateRow Then
                            rsDest![File Name] = getFileName(Me.fileP)
                            rsDest![Status] = "Closed"
                            rsDest![Week Closed] = Int((Date - CDate("3/21/2016")) / 7) + 1
                            rsDest![Days Active] = Date - rsDest![Date Received from BCUS]
                            rsDest![Date Closed] = Date
                            rsDest![Date Last Updated] = Now
                        End If
                        
                        rsDest.Update
                        rsDest.MoveNext
                        
                    Loop
                End If
                
                'Case 0: Brand new record that does not exist in master DB
                If found = False And rsSrc![Cardmember Number] <> "ERROR" Then
    
                    'Fields that are the same for Research & Chargeback, import them
                    rsDest.AddNew
                    rsDest![Cardmember Number] = rsSrc![Cardmember Number]
                    rsDest![Transaction Amt] = rsSrc![Transaction Amt]
                    rsDest![Transaction Date] = rsSrc![Transaction Date]
                    rsDest![Auth Code] = rsSrc![Auth Code *]
                    rsDest![SE Name] = rsSrc![SE Name *]
                    rsDest![SE Number] = rsSrc![SE Number]
                    rsDest![Transaction ID] = rsSrc![Transaction ID]
                    
                    'Input User Submitted fields
                    rsDest![Date Received from BCUS] = Me.dateS
                    rsDest![Date Last Updated] = Date
                    rsDest![Week Opened] = Int((Me.dateS - CDate("3/21/2016")) / 7) + 1
                    rsDest![Request Type] = Me.templateT
                    rsDest![File Name] = getFileName(Me.fileP)
                    
                    'Calculate Fields
                    rsDest![Status] = "Open"
                    If Me.templateT = "Fraud Research" Then 'If research
                        rsDest![AXP Case Due Date] = Me.dateS + 10
                    ElseIf Me.templateT = "Fraud Chargeback" Then 'If chargeback
                        rsDest![AXP Case Due Date] = Me.dateS + 45
                    End If
                                      
                    'Case 0: Fraud Chargeback Table
                    If templateT = "Fraud Chargeback" Then
                        
                        rsDest![Date Reported Fraud] = rsSrc![Date Reported Fraud]
                        rsDest![Reason for Claiming Fraud] = rsSrc![Reason for Claiming Fraud?]
                        rsDest![Card in Possession] = rsSrc![Card in Possession?]
                        rsDest![Date Lost/ Stolen] = rsSrc![Date Lost/Stolen *]
                        rsDest![Dispute being Withdrawn] = rsSrc![Dispute being Withdrawn?]
                        rsDest![Credits] = rsSrc![Credits]
                        
                    End If
                    
                rsDest.Update
                
                End If
            End With
            
            counter = counter + 1
            rsSrc.MoveNext
        
        Loop
    End With
    
    'Close DB data stream
    rsDest.Close
    Set rsDest = Nothing
    rsSrc.Close
    Set rsSrc = Nothing
    Set db = Nothing
    
    'Alert user & open import error tables if there is an error
    If tableError Then
        alertImportError ("There are one or more errors in the data uploaded. Please see Import Errors table.")
    End If
    MsgBox "Import Successful!", vbOKOnly
    
End Function

'Integrate Fraud Application template (only) w Fraud Application table by transferring over data in an appropriate fashion (see cases)

Private Function importTODBMISC(requestType As String)

    'Set up Recordsets
    Dim db As DAO.Database
    Dim rsSrc As DAO.Recordset
    Dim rsDest As DAO.Recordset
    Dim found As Boolean
    Dim errorTable As DAO.Recordset
    
    'Set db variables
    Set db = CurrentDb
    Set rsSrc = db.OpenRecordset(Me.templateT, dbOpenSnapshot)
    Set rsDest = db.OpenRecordset("Additional Request Table", _
        dbOpenTable, dbAppendOnly)
    Set errorTable = db.OpenRecordset("Import Errors", _
        dbOpenTable, dbAppendOnly)
    tableError = False
    counter = 4
    
    'Compare each record in DB to record in upload file
    'loop through each record in upload temp table
    With rsSrc
        Do Until rsSrc.EOF
                        
            'compare to each record in master table
            With rsDest
                               
                'If there are records in the table, start comparing
                If Not DCount("*", "Additional Request Table") = 0 Then
                    
                    rsDest.MoveFirst
                    
                    found = False
                    Do Until rsDest.EOF
                        
                        rowError = False
                        updateRow = False
                        'Record already exists in the database
                        If rsSrc![Cardmember Number] = rsDest![Cardmember Number] And rsDest![Request Type] = requestType And rsSrc![Cardmember Number] <> "ERROR" Then
                            rsDest.Edit
                            
                            If requestType <> "Statement Request" And requestType <> "Cardmember Correspondence" Then
                                found = True
                                
                                'Add AXP return fields if any
                                'Fraud App
                                If requestType = "Fraud Application" Then
                                    If Not IsNull(rsSrc![Tradeline Deleted]) Then
                                        
                                        'Add if not dupe
                                        If IsNull(rsDest![Tradeline Deleted]) Then
                                            rsDest![Tradeline Deleted] = rsSrc![Tradeline Deleted]
                                            updateRow = True
                                        
                                        'Dupe w BCUS and AXP fields
                                        Else
                                            errorMsg = "This record is a duplicate with BCUS and AXP info."
                                            rowError = True
                                        End If
                                    
                                    'Dupe w BCUS info only
                                    Else
                                        errorMsg = "This record is a duplicate with BCUS info."
                                        rowError = True
                                    End If
                                
                                'CM App Request
                                ElseIf requestType = "CM Application Request" Then
                                    If Not IsNull(rsSrc![Date CM Application Sent]) Then
                                        'Add if not dupe
                                        If IsNull(rsDest![Date CM Application Sent]) Then
                                            updateRow = True
                                            rsDest![Date CM Application Sent] = rsSrc![Date CM Application Sent]
                                        
                                        'Dupe w BCUS and AXP fields
                                        Else
                                            errorMsg = "This record is a duplicate with BCUS and AXP info."
                                            rowError = True
                                        End If
                                    
                                    'Dupe w BCUS info only
                                    Else
                                        errorMsg = "This record is a duplicate with BCUS info."
                                        rowError = True
                                    End If
                                
                                'General Inquiry
                                ElseIf requestType = "General Inquiry" Then
                                    If Not IsNull(rsSrc![Reply Document Indicator (Y/N)]) Then
                                        
                                        'Add if not dupe
                                        If IsNull(rsDest![Reply Document Indicator]) Then
                                            updateRow = True
                                            rsDest![Reply Document Indicator] = rsSrc![Reply Document Indicator (Y/N)]
                                            rsDest![Date General Inquiry Response Sent] = rsSrc![Date General Inquiry Response Sent]
                                        
                                        'Dupe w BCUS and AXP fields
                                        Else
                                            errorMsg = "This record is a duplicate with BCUS and AXP info."
                                            rowError = True
                                        End If
                                    
                                    'Dupe w BCUS info only
                                    Else
                                        errorMsg = "This record is a duplicate with BCUS info."
                                        rowError = True
                                    End If
                                End If
                            Else
                                'Check statement request
                                If rsDest![Request Type] = "Statement Request" Then
                                    If CStr(rsDest![Statement Month]) = CStr(rsSrc![Statement Month] And CStr(rsSrc![Statement Year]) = CStr(rsDest![Statement Year])) Then
                                        
                                        found = True
                                        If Not IsNull(rsSrc![Date CM Statement Sent]) Then
                                            
                                            'Add if not dupe
                                            If IsNull(rsDest![Date CM Statement Sent]) Then
                                                
                                                updateRow = True
                                                rsDest![Date CM Statement Sent] = rsSrc![Date CM Statement Sent]
                                                rsDest![Status] = "Closed"
                                                rsDest![Week Closed] = Int((Date - CDate("3/21/2016")) / 7) + 1
                                                rsDest![Days Active] = Date - rsDest![Date Received from BCUS]
                                                rsDest![Date Last Updated] = Now
                                                rsDest![Date Closed] = Date
                                                rsDest![File Name] = getFileName(Me.fileP)
                                            
                                            'Dupe w BCUS and AXP fields
                                            Else
                                                errorMsg = "This record is a duplicate with BCUS and AXP info."
                                                rowError = True
                                            End If
                                        
                                        'Dupe w BCUS info only
                                        Else
                                            errorMsg = "This record is a duplicate with BCUS info."
                                            rowError = True
                                        End If
                                    End If
                                ElseIf rsDest![Request Type] = "Cardmember Correspondence" Then
                                    If CStr(rsDest![Date Correspondence Sent]) = CStr(rsSrc![Date Correspondence Sent]) Then
                                        found = True
                                        rowError = True
                                        errorMsg = "This record is a duplicate with BCUS info."
                                    End If
                                End If
                                    
                            End If
                        rsDest.Update
                            
                            'if we are updating the records update the common fields
                            If updateRow Then
                                'add fields in common
                                rsDest.Edit
                                rsDest![Status] = "Closed"
                                rsDest![Week Closed] = Int((Date - CDate("3/21/2016")) / 7) + 1
                                rsDest![Days Active] = Date - rsDest![Date Received from BCUS]
                                rsDest![Date Last Updated] = Now
                                rsDest![Date Closed] = Date
                                rsDest![File Name] = getFileName(Me.fileP)
                                rsDest.Update
                            End If
                        
                        End If
                                                   
                        'if error then add error to table
                        If rowError Then
                            
                            tableError = True
                            errorTable.AddNew
                            errorTable![row] = counter
                            errorTable![error] = errorMsg
                            errorTable![File Name] = getFileName(Me.fileP)
                            If Me.templateT <> "Fraud Research" And Me.templateT <> "Fraud Chargeback" And Me.templateT <> "Fraud Application" Then
                                errorTable![Cardmember Name] = rsSrc![Cardmember Name]
                            End If
                            errorTable![Cardmember Number] = rsDest![Cardmember Number]
                            errorTable.Update
                        End If
                    
                        rsDest.MoveNext
                    Loop
                End If
                
            
                'Case 0: Brand new record that does not exist in master DB
                If found = False And rsSrc![Cardmember Number] <> "ERROR" Then
                    
                    'Add common fields in all templates
                    rsDest.AddNew
                    rsDest![Cardmember Number] = rsSrc![Cardmember Number]
                    rsDest![Request Type] = requestType
                    rsDest![File Name] = getFileName(Me.fileP)
                    rsDest![Date Received from BCUS] = dateS
                    rsDest![Date Last Updated] = Now
                    rsDest![Status] = "Open"
                    rsDest![Week Opened] = Int((dateS - CDate("3/21/2016")) / 7) + 1
        
                    If requestType <> "Fraud Application" Then
                        rsDest![Cardmember Name] = rsSrc![Cardmember Name]
                        
                        'Add unique fields from each template
                        'Statement Request
                        If requestType = "Statement Request" Then
                            rsDest![Statement Month] = rsSrc![Statement Month]
                            rsDest![Statement Year] = rsSrc![Statement Year]
                            rsDest![AXP Case Due Date] = dateS + 7
                        
                        'Cm Application Request
                        ElseIf requestType = "CM Application Request" Then
                            rsDest![Affidavit Provided] = rsSrc![Affadavit Provided (Y/N)]
                            rsDest![AXP Case Due Date] = dateS + 30
                        
                        'Cardmember Correspondence
                        ElseIf requestType = "Cardmember Correspondence" Then
                            rsDest![Reply Document Indicator] = rsSrc![Reply Document Indicator (Y/N)]
                            rsDest![Date Correspondence Sent] = rsSrc![Date Correspondence Sent]
                            rsDest![Status] = "Closed"
                            rsDest![Date Closed] = Date
                            rsDest![Week Closed] = Int((Date - CDate("3/21/2016")) / 7) + 1
                            rsDest![Days Active] = Date - rsDest![Date Received from BCUS]
                            rsDest![AXP Case Due Date] = dateS + 2
                        
                        'General Inquiry
                        ElseIf requestType = "General Inquiry" Then
                            rsDest![Document Indicator] = rsSrc![Document Indicator (Y/N)]
                            rsDest![Free Text Box] = rsSrc![Free Text Box *]
                            rsDest![AXP Case Due Date] = dateS + 7
                        End If
                        
                    Else 'If Fraud Application
                        rsDest![Date Reported Fraud] = rsSrc![Date Reported Fraud]
                        rsDest![AXP Case Due Date] = dateS + 10
                    End If
                    
                    rsDest.Update
                    
                End If
            End With
            
            counter = counter + 1
            rsSrc.MoveNext
        
        Loop
    End With
    
    rsDest.Close
    Set rsDest = Nothing
    rsSrc.Close
    Set rsSrc = Nothing
    Set db = Nothing
    
    'Alert user & open import error tables if there is an error
    If tableError Then
        alertImportError ("There are one or more errors in the data uploaded. Please see Import Errors table.")
    End If
    MsgBox "Import Successful!", vbOKOnly
    

End Function

'Integrate Non Fraud templates (only) w Non Fraud master table by transferring over data in an appropriate fashion (see cases)

Private Function importTODBNF()
    
    'Set up Recordsets
    Dim db As DAO.Database
    Dim rsSrc As DAO.Recordset
    Dim rsDest As DAO.Recordset
    Dim errorTable As DAO.Recordset
    Dim found As Boolean
    
    'Set db variables
    Set db = CurrentDb
    Set rsSrc = db.OpenRecordset(Me.templateT, dbOpenSnapshot)
    Set rsDest = db.OpenRecordset("Non Fraud Table", _
        dbOpenTable, dbAppendOnly)
    Set errorTable = db.OpenRecordset("Import Errors", _
        dbOpenTable, dbAppendOnly)
    
    counter = 4
    tableError = False
    'Compare each record in DB to record in upload file
    'loop through each record in upload temp table
    With rsSrc
        Do Until rsSrc.EOF
                        
            
            'compare to each record in master table
            With rsDest
                
                found = False
                
                'If there are records in the table, start comparing
                If Not DCount("*", "Non Fraud Table") = 0 Then
                
                    rsDest.MoveFirst
                    
                    Do Until rsDest.EOF
                        
                        rsDest.Edit
                        rowError = False
                        updateRow = False
                        
                        'Record already exists in the database
                        If rsSrc![Transaction ID] = rsDest![Transaction ID] And Me.templateT = rsDest![Request Type] And rsSrc![Cardmember Number] <> "ERROR" Then
                            found = True
                                                        
                            If rsDest![Request Type] = "Non Fraud Retrieval" Then
                                
                                'Case 1a: During retrieval request, AX gets back to BCUS w info
                                If Not IsNull(rsSrc![Data Capture Date]) Then
                                    
                                    'Add fields if not dupe
                                    If IsNull(rsDest![Data Capture Date]) Then
                                        rsDest![AXP Case ID Number] = rsSrc![Amex Case ID #]
                                        rsDest![Data Capture Date] = rsSrc![Data Capture Date]
                                        rsDest![Reply Document Indicator] = rsSrc![Reply Document Indicator]
                                        rsDest![Reply Message] = rsSrc![Reply Message *]
                                        rsDest![Dispute Stage] = "Fulfillment"
                                        updateRow = True
                                    
                                    'This is an AX response dupe
                                    Else
                                        errorMsg = "This record is a duplicate with BCUS and AXP info."
                                        rowError = True
                                    End If
                                    
                                'This is a BCUS request dupe
                                Else
                                    errorMsg = "This record is a duplicate with BCUS info only."
                                    rowError = True
                                End If
                            
                            ElseIf rsDest![Request Type] = "Non Fraud Chargeback" Then
                            
                                'Case 1b: During chargeback request, AX gets back to BCUS w info
                                If Not IsNull(rsSrc![Data Capture Date]) Then

                                    'Add fields if not dupe
                                    If IsNull(rsDest![Data Capture Date]) Then
                                        
                                        rsDest![Case Due Date] = rsSrc![Case Due Date]
                                        rsDest![Data Capture Date] = rsSrc![Data Capture Date]
                                        rsDest![Dispute Received Date] = rsSrc![Dispute Received Date]
                                        rsDest![Reply Document Indicator] = rsSrc![Reply Document Indicator]
                                        rsDest![Reply Message] = rsSrc![Reply Message *]
                                        rsDest![Dispute Stage] = "Chargeback Confirmation"
                                        updateRow = True
                                    
                                    'This is an AXP response dupe
                                    Else
                                        errorMsg = "This record is a duplicate with BCUS and AXP chargeback info."
                                        rowError = True
                                    End If
                                    
                                'This is a BCUS request dupe
                                ElseIf IsNull(rsSrc![Representment Date]) Then
                                    errorMsg = "This record is a duplicate with BCUS info only ded."
                                    rowError = True
                                End If
                                
                                'Case 1c: During chargeback request, merchant represents, AX gives BCUS representment info
                                If Not IsNull(rsSrc![Representment Date]) Then
                                    
                                    'Add if not dupe
                                    If IsNull(rsDest![Representment Date]) Then
                                        rsDest![Representment Amount] = rsSrc![Representment Amount]
                                        rsDest![Representment Date] = rsSrc![Representment Date]
                                        rsDest![Dispute Stage] = "Representment"
                                        updateRow = True
                                    
                                    'This is an AXP representment dupe
                                    Else
                                        errorMsg = "This record is a duplicate with BCUS and AXP representment infoxxxx."
                                        rowError = True
                                    End If
                                End If

                            End If
                            
                            If rsSrc![Dispute/Retrieval Reason] = "CM No longer disputing" Then
                                rowError = False
                                rsDest![Dispute/Retrieval Reason] = rsSrc![Dispute/Retrieval Reason]
                                rsDest![Status] = "Closed"
                                rsDest![Week Closed] = Int((Date - CDate("3/21/2016")) / 7) + 1
                                rsDest![Days Active] = Date - rsDest![Date Received from BCUS]
                                rsDest![Date Closed] = Date
                                rsDest![File Name] = getFileName(Me.fileP)
                            End If
                            
                        End If
                            
                        'If there is a dupe add to error table
                        If rowError Then
                            tableError = True
                            errorTable.AddNew
                            errorTable![row] = counter
                            errorTable![error] = errorMsg
                            errorTable![File Name] = getFileName(Me.fileP)
                            errorTable![Cardmember Name] = rsDest![Cardmember Name]
                            errorTable![Cardmember Number] = rsDest![Cardmember Number]
                            errorTable.Update
                        ElseIf updateRow Then
                            rsDest![Status] = "Closed"
                            rsDest![Week Closed] = Int((Date - CDate("3/21/2016")) / 7) + 1
                            rsDest![Days Active] = Date - rsDest![Date Received from BCUS]
                            rsDest![Date Closed] = Date
                            rsDest![Date Last Updated] = Now
                            rsDest![File Name] = getFileName(Me.fileP)
                        End If
                            
                        rsDest.Update
                        rsDest.MoveNext
                    Loop
                End If
                
                'Case 0: Brand new record that does not exist in master DB
                If found = False And rsSrc![Cardmember Number] <> "ERROR" Then
                        
                        'Fields that are the same for Retrieval & Chargeback, import them
                        rsDest.AddNew
                        rsDest![Cardmember Name] = rsSrc![Cardmember Name]
                        rsDest![Cardmember Number] = rsSrc![Cardmember Number]
                        rsDest![Chargeback Reason] = rsSrc![Chargeback Reason]
                        rsDest![Country] = rsSrc![Country]
                        rsDest![Dispute Stage] = rsSrc![Dispute Stage]
                        rsDest![Dispute/Retrieval Reason] = rsSrc![Dispute/Retrieval Reason]
                        rsDest![Disputed Amount] = rsSrc![Disputed Amount]
                        rsDest![Document Indicator] = rsSrc![Document Indicator]
                        rsDest![Financial Reference/Case Num] = rsSrc![Financial Reference/Case #]
                        rsDest![Post Date] = rsSrc![Post Date]
                        rsDest![Representment Indicator] = rsSrc![Representment Indicator *]
                        rsDest![Representment Reason] = rsSrc![Representment Reason *]
                        rsDest![SE Name] = rsSrc![SE Name *]
                        rsDest![SE Number] = rsSrc![SE Number]
                        rsDest![Statement Date] = rsSrc![Statement Date]
                        rsDest![Transaction Amt] = rsSrc![Transaction Amt]
                        rsDest![Transaction Amt in Original Currency] = rsSrc![Transaction Amt in Original Currency]
                        rsDest![Transaction Cd] = rsSrc![Transaction Cd]
                        rsDest![Transaction Date] = rsSrc![Transaction Date]
                        rsDest![Transaction ID] = rsSrc![Transaction ID]
                        
                        'Input User Submitted fields
                        rsDest![Date Received from BCUS] = Me.dateS
                        rsDest![Request Type] = Me.templateT
                        rsDest![File Name] = getFileName(Me.fileP)
                        rsDest![Week Opened] = Int((Me.dateS - CDate("3/21/2016")) / 7) + 1
                        
                        'Calculate Fields
                        If (rsDest![Country] = "USA") Or (rsDest![Country] = "US") Then 'If country is USA
                            rsDest![AXP Case Due Date] = Me.dateS + 30
                        Else 'If international
                            rsDest![AXP Case Due Date] = Me.dateS + 45
                        End If
                        rsDest![Status] = "Open"
                        rsDest![Date Last Updated] = Now
                        
                        If rsSrc![Dispute/Retrieval Reason] = "CM No longer disputing" Then
                            rowError = False
                            rsDest![Dispute/Retrieval Reason] = rsSrc![Dispute/Retrieval Reason]
                            rsDest![Status] = "Closed"
                            rsDest![Week Closed] = Int((Date - CDate("3/21/2016")) / 7) + 1
                            rsDest![Days Active] = Date - rsDest![Date Received from BCUS]
                            rsDest![Date Closed] = Date
                            rsDest![File Name] = getFileName(Me.fileP)
                        End If
                    
                    'Case 0a: Non Fraud Retrieval Table
                    If templateT = "Non Fraud Retrieval" Then
                        
                        'Copy over the fields
                        rsDest![Case Due Date] = rsSrc![Case Due Date]
                        rsDest![Dispute Received Date] = rsSrc![Dispute Received Date]
                        
                        'if this is in flight w AXP fields already
                        If Not IsNull(rsSrc![Amex Case ID #]) Then
                            rsDest![AXP Case ID Number] = rsSrc![Amex Case ID #]
                            rsDest![Data Capture Date] = rsSrc![Data Capture Date]
                            rsDest![Reply Document Indicator] = rsSrc![Reply Document Indicator]
                            rsDest![Reply Message] = rsSrc![Reply Message *]
                            rsDest![Dispute Stage] = "Fulfillment"
                            rsDest![Status] = "Closed"
                            rsDest![Week Closed] = Int((Date - CDate("3/21/2016")) / 7) + 1
                            rsDest![Days Active] = Date - rsDest![Date Received from BCUS]
                            rsDest![Date Closed] = Date
                            rsDest![Date Last Updated] = Now
                        End If
                    
                    'Case 0b: Non Fraud Chargeback Table
                    ElseIf templateT = "Non Fraud Chargeback" Then
                        
                        'Copy over the fields
                        rsDest![AXP Case ID Number] = rsSrc![Amex Case ID # *]
                        rsDest![Case Opened Date] = rsSrc![Case Opened Date]
                        rsDest![Chargeback Amount] = rsSrc![Chargeback Amount ($)]
                        rsDest![Current Valid Cardmember Number] = rsSrc![Current Valid Cardmember Number]
                        rsDest![Document Type] = rsSrc![Document Type *]
                        
                        'if this is in flight w AXP chargeback fields already
                        If Not IsNull(rsSrc![Data Capture Date]) Then
                            rsDest![Case Due Date] = rsSrc![Case Due Date]
                            rsDest![Data Capture Date] = rsSrc![Data Capture Date]
                            rsDest![Dispute Received Date] = rsSrc![Dispute Received Date]
                            rsDest![Reply Document Indicator] = rsSrc![Reply Document Indicator]
                            rsDest![Reply Message] = rsSrc![Reply Message *]
                            rsDest![Dispute Stage] = "Chargeback Confirmation"
                            rsDest![Status] = "Closed"
                            rsDest![Week Closed] = Int((Date - CDate("3/21/2016")) / 7) + 1
                            rsDest![Days Active] = Date - rsDest![Date Received from BCUS]
                            rsDest![Date Closed] = Date
                            rsDest![Date Last Updated] = Now
                        End If
                        
                        'if this is in flight w AXP representment fields already
                        If Not IsNull(rsSrc![Representment Amount]) Then
                            rsDest![Representment Amount] = rsSrc![Representment Amount]
                            rsDest![Representment Date] = rsSrc![Representment Date]
                            rsDest![Dispute Stage] = "Representment"
                            rsDest![Status] = "Closed"
                            rsDest![Week Closed] = Int((Date - CDate("3/21/2016")) / 7) + 1
                            rsDest![Days Active] = Date - rsDest![Date Received from BCUS]
                            rsDest![Date Closed] = Date
                            rsDest![Date Last Updated] = Now
                        End If
                        
                    End If
                    
                    rsDest.Update
                    
                End If
            End With
            
            counter = counter + 1
            rsSrc.MoveNext
        
        Loop
    End With
    
    'Close DB data flow
    rsDest.Close
    Set rsDest = Nothing
    rsSrc.Close
    errorTable.Close
    Set rsSrc = Nothing
    Set db = Nothing
    Set errorTable = Nothing
    
    'Alert user & open import error tables if there is an error
    If tableError Then
        alertImportError ("There are one or more errors in the data uploaded. Please see Import Errors table.")
    End If
    MsgBox "Import Successful!", vbOKOnly
    
End Function

'Check to see that the data records uploaded are complete. Checks all BCUS fields and selectively checks AXP fields if at least
'one AXP field is populated

Private Function checkComplete(template As String)
    
    'Set up Recordsets
    Dim db As DAO.Database
    Dim tableToCheck As DAO.Recordset
    Dim errorTable As DAO.Recordset
    
    'Set db variables
    Set db = CurrentDb
    Set tableToCheck = db.OpenRecordset(template)
    Set errorTable = db.OpenRecordset("Import Errors", _
        dbOpenTable, dbAppendOnly)
    
    counter = 4
    tableError = False
    
    Do Until tableToCheck.EOF
        
        tableToCheck.Edit
        errorMsg = ""
        rowError = False
        
        'Delete the fields that are blank
        If Me.templateT = "Non Fraud Retrieval" Or Me.templateT = "Non Fraud Research" Or Me.templateT = "Fraud Research" Or Me.templateT = "Fraud Chargeback" Then
            If Trim(tableToCheck![Cardmember Number]) = "" And Trim(tableToCheck![Transaction ID]) = "" And Trim(tableToCheck![SE Number]) = "" Then
                tableToCheck![Cardmember Number] = "ERROR"
            ElseIf IsNull(tableToCheck![Cardmember Number]) And Trim(tableToCheck![Transaction ID]) = "" And Trim(tableToCheck![SE Number]) = "" Then
                tableToCheck![Cardmember Number] = "ERROR"
            ElseIf Trim(tableToCheck![Cardmember Number]) = "" And IsNull(tableToCheck![Transaction ID]) And Trim(tableToCheck![SE Number]) = "" Then
                tableToCheck![Cardmember Number] = "ERROR"
            ElseIf Trim(tableToCheck![Cardmember Number]) = "" And Trim(tableToCheck![Transaction ID]) = "" And IsNull(tableToCheck![SE Number]) Then
                tableToCheck![Cardmember Number] = "ERROR"
            ElseIf IsNull(tableToCheck![Cardmember Number]) And IsNull(tableToCheck![Transaction ID]) And Trim(tableToCheck![SE Number]) = "" Then
                tableToCheck![Cardmember Number] = "ERROR"
            ElseIf IsNull(tableToCheck![Cardmember Number]) And Trim(tableToCheck![Transaction ID]) = "" And IsNull(tableToCheck![SE Number]) Then
                tableToCheck![Cardmember Number] = "ERROR"
            ElseIf Trim(tableToCheck![Cardmember Number]) = "" And IsNull(tableToCheck![Transaction ID]) And IsNull(tableToCheck![SE Number]) Then
                tableToCheck![Cardmember Number] = "ERROR"
            End If
        ElseIf Me.templateT = "Fraud Application" Then
            If (Trim(tableToCheck![Cardmember Number]) = "" And Trim(tableToCheck![Date Reported Fraud]) = "") Or (IsNull(tableToCheck![Date Reported Fraud]) And Trim(tableToCheck![Cardmember Number]) = "") Then
                tableToCheck![Cardmember Number] = "ERROR"
            End If
        Else
            If Trim(tableToCheck![Cardmember Number]) = "" And Trim(tableToCheck![Cardmember Name]) = "" Then
                tableToCheck![Cardmember Number] = "ERROR"
            End If
        End If
        
        If IsNull(tableToCheck![Cardmember Number]) Then
            errorMsg = errorMsg + "Cardmember Number is blank or improperly formatted. "
            rowError = True
            tableToCheck![Cardmember Number] = "ERROR"
            
        End If
        
        If tableToCheck![Cardmember Number] <> "ERROR" Then
                
            'NF Retrieval
            If template = "Non Fraud Retrieval" Then
                
                'Check BCUS fields complete
                Dim NFRArray As Variant
                NFRArray = Array("Cardmember Name", "Cardmember Number", "Case Due Date", "Chargeback Reason", "Country", "Dispute Received Date", "Dispute Stage", "Dispute/Retrieval Reason", "Disputed Amount", "Document Indicator", "Financial Reference/Case #", "Post Date", "SE Number", "Statement Date", "Transaction Amt", "Transaction Amt in Original Currency", "Transaction Cd", "Transaction Date", "Transaction ID")
                
                For Each element In NFRArray
                 
                    If Not IsNull(tableToCheck.Fields(element)) Then
                        If Trim(tableToCheck.Fields(element)) = "" Then
                            errorMsg = errorMsg + element + " is blank or improperly formatted. "
                            rowError = True
                        End If
                    Else
                        errorMsg = errorMsg + element + " is blank or improperly formatted. "
                        rowError = True
                    End If
                
                Next element
                
                'Check if AXP fields are complete
                'If at least one AXP field is filled in...
                If (Not IsNull(tableToCheck![Amex Case ID #]) Or Not IsNull(tableToCheck![Data Capture Date]) Or Not IsNull(tableToCheck![Reply Document Indicator])) Then
                    
                    '...check if AXP fields are complete
                    Dim NFRAXPArray As Variant
                    NFRAXPArray = Array("Amex Case ID #", "Data Capture Date", "Reply Document Indicator")
                    
                    For Each element In NFRAXPArray
                        
                        If Not IsNull(tableToCheck.Fields(element)) Then
                            If Trim(tableToCheck.Fields(element)) = "" Then
                                errorMsg = errorMsg + element + " is blank or improperly formatted. "
                                rowError = True
                            End If
                        Else
                            errorMsg = errorMsg + element + " is blank or improperly formatted. "
                            rowError = True
                        End If
                        
                    Next element
                End If
                
            End If
                
            'NF Chargeback
            If template = "Non Fraud Chargeback" Then
                
                'Check BCUS fields
                Dim NFCArray As Variant
                NFCArray = Array("Cardmember Name", "Cardmember Number", "Case Opened Date", "Chargeback Amount ($)", "Chargeback Reason", "Country", "Current Valid Cardmember Number", "Dispute Stage", "Dispute/Retrieval Reason", "Disputed Amount", "Document Indicator", "Financial Reference/Case #", "Post Date", "SE Number", "Statement Date", "Transaction Amt", "Transaction Amt in Original Currency", "Transaction Cd", "Transaction Date", "Transaction ID")
                
                For Each element In NFCArray
                    
                    If Not IsNull(tableToCheck.Fields(element)) Then
                        If Trim(tableToCheck.Fields(element)) = "" Then
                            errorMsg = errorMsg + element + " is blank or improperly formatted. "
                            rowError = True
                        End If
                    Else
                        errorMsg = errorMsg + element + " is blank or improperly formatted. "
                        rowError = True
                    End If
                 
                 Next element
                 
                 'Check AXP Chargeback Fields
                 'If at least one AXP field is filled in...
                 If (Not IsNull(tableToCheck![Case Due Date]) Or Not IsNull(tableToCheck![Data Capture Date]) Or Not IsNull(tableToCheck![Dispute Received Date]) Or Not IsNull(tableToCheck![Reply Document Indicator])) Then
                     
                     '...Check if AXP fields are complete
                     Dim NFCAXPArray As Variant
                     NFCAXPArray = Array("Case Due Date", "Data Capture Date", "Dispute Received Date", "Reply Document Indicator")
                     
                     For Each element In NFCAXPArray
                         
                         If Not IsNull(tableToCheck.Fields(element)) Then
                             If Trim(tableToCheck.Fields(element)) = "" Then
                                 errorMsg = errorMsg + element + " is blank or improperly formatted. "
                                 rowError = True
                             End If
                         Else
                             errorMsg = errorMsg + element + " is blank or improperly formatted. "
                             rowError = True
                         End If
                         
                     Next element
                 End If
                 
                 'Check AXP Representment Fields
                 'If at least one AXP representment field is filled in...
                 If (Not IsNull(tableToCheck![Representment Amount]) Or Not IsNull(tableToCheck![Representment Date])) Then
                     
                     '...check if AXP representment fields are completed
                     Dim NFCAXPRArray As Variant
                     NFCAXPRArray = Array("Representment Amount", "Representment Date")
                     
                     For Each element In NFCAXPRArray
                         
                         If Not IsNull(tableToCheck.Fields(element)) Then
                             If Trim(tableToCheck.Fields(element)) = "" Then
                                 errorMsg = errorMsg + element + " is blank or improperly formatted. "
                                 rowError = True
                             End If
                         Else
                             errorMsg = errorMsg + element + " is blank or improperly formatted. "
                             rowError = True
                         End If
                         
                     Next element
                 End If
                
             End If
        
            
             'F Research
             If template = "Fraud Research" Then
                 
                 'Check BCUS fields are complete
                 Dim FRArray As Variant
                 FRArray = Array("Cardmember Number", "SE Number", "Transaction Amt", "Transaction Date", "Transaction ID")
                 
                 For Each element In FRArray
                     
                     If Not IsNull(tableToCheck.Fields(element)) Then
                         If Trim(tableToCheck.Fields(element)) = "" Then
                             errorMsg = errorMsg + element + " is blank or improperly formatted. "
                             rowError = True
                         End If
                     Else
                         errorMsg = errorMsg + element + " is blank or improperly formatted. "
                         rowError = True
                     End If
                     
                 Next element
                 
                 'Check AXP fields
                 'If at least one AXP field is filled in...
                 If (Not IsNull(tableToCheck![Fees]) Or Not IsNull(tableToCheck![Credits]) Or Not IsNull(tableToCheck![Finance Charges]) Or Not IsNull(tableToCheck![Late Payment]) Or Not IsNull(tableToCheck![Re-Tier]) Or Not IsNull(tableToCheck![Re-Age])) Then
                     
                     '...check if AXP representment fields are completed
                     Dim FRAXPArray As Variant
                     FRAXPArray = Array("Fees", "Credits", "Finance Charges", "Late Payment", "Re-Tier", "Re-Age")
                     
                     For Each element In FRAXPArray
                     
                        If Not IsNull(tableToCheck.Fields(element)) Then
                            If Trim(tableToCheck.Fields(element)) = "" Then
                                errorMsg = errorMsg + element + " is blank or improperly formatted. "
                                rowError = True
                            End If
                        Else
                            errorMsg = errorMsg + element + " is blank or improperly formatted. "
                            rowError = True
                        End If
                        
                    Next element
                End If
            End If
            
            'F Chargeback
            If template = "Fraud Chargeback" Then
                
                'Check BCUS fields
                Dim FCArray As Variant
                FCArray = Array("Cardmember Number", "Date Reported Fraud", "Reason for Claiming Fraud?", "Card in Possession?", "Dispute being Withdrawn?", "Credits", "SE Number", "Transaction Amt", "Transaction Date", "Transaction ID")
                
                For Each element In FCArray
                    
                    If Not IsNull(tableToCheck.Fields(element)) Then
                        If Trim(tableToCheck.Fields(element)) = "" Then
                            errorMsg = errorMsg + element + " is blank or improperly formatted. "
                            rowError = True
                        End If
                    Else
                        errorMsg = errorMsg + element + " is blank or improperly formatted. "
                        rowError = True
                    End If
                Next element
                
            End If
            
            'FRAP
            If template = "Fraud Application" Then
                Dim FRAPArray As Variant
                FRAPArray = Array("Cardmember Number", "Date Reported Fraud")
                For Each element In FRAPArray
                    
                    If Not IsNull(tableToCheck.Fields(element)) Then
                        If Trim(tableToCheck.Fields(element)) = "" Then
                            errorMsg = errorMsg + element + " is blank or improperly formatted. "
                            rowError = True
                        End If
                    Else
                        errorMsg = errorMsg + element + " is blank or improperly formatted. "
                        rowError = True
                    End If
                Next element
            End If
            
            'Statement Request
            If template = "Statement Request" Then
                Dim SRArray As Variant
                SRArray = Array("Cardmember Name", "Cardmember Number", "Statement Month", "Statement Year")
                For Each element In SRArray
                    
                    If Not IsNull(tableToCheck.Fields(element)) Then
                        If Trim(tableToCheck.Fields(element)) = "" Then
                            errorMsg = errorMsg + element + " is blank or improperly formatted. "
                            rowError = True
                        End If
                    Else
                        errorMsg = errorMsg + element + " is blank or improperly formatted. "
                        rowError = True
                    End If
                Next element
            End If
            
            'CM Application Request
            If template = "CM Application Request" Then
                Dim CMARArray As Variant
                CMARArray = Array("Cardmember Name", "Cardmember Number", "Affadavit Provided (Y/N)")
                For Each element In CMARArray
                    
                    If Not IsNull(tableToCheck.Fields(element)) Then
                        If Trim(tableToCheck.Fields(element)) = "" Then
                            errorMsg = errorMsg + element + " is blank or improperly formatted. "
                            rowError = True
                        End If
                    Else
                        errorMsg = errorMsg + element + " is blank or improperly formatted. "
                        rowError = True
                    End If
                Next element
            End If
            
            'Cardmember Correspondence
            If template = "Cardmember Correspondence" Then
                Dim CCArray As Variant
                
                
                CCArray = Array("Cardmember Name", "Cardmember Number", "Reply Document Indicator (Y/N)", "Date Correspondence Sent")
                For Each element In CCArray
                    
                    If Not IsNull(tableToCheck.Fields(element)) Then
                        If Trim(tableToCheck.Fields(element)) = "" Then
                            errorMsg = errorMsg + element + " is blank or improperly formatted. "
                            rowError = True
                        End If
                    Else
                        errorMsg = errorMsg + element + " is blank or improperly formatted. "
                        rowError = True
                    End If
                Next element
            End If
            
            'General Inquiries
            If template = "General Inquiry" Then
                
                Dim GIArray As Variant
                GIArray = Array("Cardmember Name", "Cardmember Number", "Document Indicator (Y/N)")
                For Each element In GIArray
                    
                    If Not IsNull(tableToCheck.Fields(element)) Then
                        If Trim(tableToCheck.Fields(element)) = "" Then
                            errorMsg = errorMsg + element + " is blank or improperly formatted. "
                            rowError = True
                        End If
                    Else
                        errorMsg = errorMsg + element + " is blank or improperly formatted. "
                        rowError = True
                    End If
                Next element
                
                'Check AXP fields
                'If at least one AXP field is filled in...
                If (Not IsNull(tableToCheck![Reply Document Indicator (Y/N)]) Or Not IsNull(tableToCheck![Date General Inquiry Response Sent])) Then
                    
                    '...check if AXP representment fields are completed
                    Dim GIAXPArray As Variant
                    GIAXPArray = Array("Reply Document Indicator (Y/N)", "Date General Inquiry Response Sent")
                    
                    For Each element In GIAXPArray
                        
                        If Not IsNull(tableToCheck.Fields(element)) Then
                            If Trim(tableToCheck.Fields(element)) = "" Then
                                errorMsg = errorMsg + element + " is blank or improperly formatted. "
                                rowError = True
                            End If
                        Else
                            errorMsg = errorMsg + element + " is blank or improperly formatted. "
                            rowError = True
                        End If
                        
                    Next element
                End If
                
            End If
        End If
            
         'if there is an error add to error table
        If rowError Then
            tableError = True
            errorTable.AddNew
            errorTable![row] = counter
            errorTable![error] = errorMsg
            errorTable![File Name] = getFileName(Me.fileP)
            
            'If template has Cardmember name add Cardmember name to error output
            If Me.templateT <> "Fraud Research" And Me.templateT <> "Fraud Chargeback" And Me.templateT <> "Fraud Application" Then
                    errorTable![Cardmember Name] = tableToCheck![Cardmember Name]
            End If
             
            errorTable![Cardmember Number] = tableToCheck![Cardmember Number]
            errorTable.Update
            tableToCheck![Cardmember Number] = "ERROR"
        End If
    
        tableToCheck.Update
        counter = counter + 1
        tableToCheck.MoveNext
    Loop
    
    'Close DB data stream
    tableToCheck.Close
    Set tableToCheck = Nothing
    errorTable.Close
    Set errorTable = Nothing
    Set db = Nothing
    
    'If there is an error alert user
    If tableError Then
        alertImportError ("There are missing values in one or more row(s) you are trying to upload. Please see Import Errors table.")
    End If
    
End Function

'Check to see that import data meets requirements and throws errors into ImportError table if errors are detected

Private Function sanitize()
    
    'Set up Recordsets
    Dim db As DAO.Database
    Dim tableToCheck As DAO.Recordset
    Dim errorTable As DAO.Recordset
    
    'Set db variables
    Set db = CurrentDb
    Set tableToCheck = db.OpenRecordset(Me.templateT)
    Set errorTable = db.OpenRecordset("Import Errors", _
        dbOpenTable, dbAppendOnly)
            
    counter = 4 'Counter starts from 4 because of headers in excel
    hasTransactionID = False
    tableError = False
    
    Do Until tableToCheck.EOF
        
        If tableToCheck![Cardmember Number] <> "ERROR" Then 'if the row data is complete
        
            rowError = False
            errorMsg = ""
            
            'Check cardmember number is 15 digit integer
            If Not isXDigitNumber(tableToCheck![Cardmember Number], 15) Then
                rowError = True
                errorMsg = errorMsg + "Cardmember Number: " + tableToCheck![Cardmember Number] + " is not a 15 digit integer. "
            End If
                        
            'Data check for common fields in NF templates & Fraud Retrieval & Fraud Chargeback
            If Me.templateT = "Non Fraud Retrieval" Or Me.templateT = "Non Fraud Chargeback" Or Me.templateT = "Fraud Research" Or Me.templateT = "Fraud Chargeback" Then
            
                'Transaction ID
                If Not isXDigitNumber(tableToCheck![Transaction ID], 18) Then
                    rowError = True
                    errorMsg = errorMsg + "Transaction ID: " + tableToCheck![Transaction ID] + " is not a 18 digit numeric value. "
                End If
                
                'SE Number
                If Not isXDigitNumber(tableToCheck![SE Number], 10) Then
                    rowError = True
                    errorMsg = errorMsg + "SE Number: " + tableToCheck![SE Number] + " is not a 10 digit integer. "
                End If
                
                'Transaction Date
                If Not IsDate(tableToCheck![Transaction Date]) Then
                    rowError = True
                    errorMsg = errorMsg + "Transaction Date: " + tableToCheck![Transaction Date] + " is not a date. "
                ElseIf (Date - CDate(tableToCheck![Transaction Date])) > 165 Then
                    rowError = True
                    errorMsg = errorMsg + "Transaction Date: " & tableToCheck![Transaction Date] & " is more than 165 days ago. "
                End If
                
            End If
            
            'Non Fraud Common Fields
            If Me.templateT = "Non Fraud Retrieval" Or Me.templateT = "Non Fraud Chargeback" Then
                
                'Transaction Amt
                If (Not IsNumeric(tableToCheck![Transaction Amt])) Or Len(tableToCheck![Transaction Amt]) >= 12 Then
                    rowError = True
                    errorMsg = errorMsg + "Transaction Amt: " + tableToCheck![Transaction Amt] + " is not a valid number. "
                ElseIf CDbl(tableToCheck![Transaction Amt] <= 30) Then
                    rowError = True
                    errorMsg = errorMsg + "Transaction Amt: " + CStr(tableToCheck![Transaction Amt]) + " is not over $30.00. "
                End If
                
                'Disputed Amt
                If (Not IsNumeric(tableToCheck![Disputed Amount])) Or (Len(tableToCheck![Disputed Amount]) >= 12) Then
                    rowError = True
                    errorMsg = errorMsg + "Disputed Amount: " + tableToCheck![Disputed Amount] + " is not a valid number. "
                End If
                
                'Post Date
                If Not IsDate(tableToCheck![Post Date]) Then
                    rowError = True
                    errorMsg = errorMsg + "Post Date: " + tableToCheck![Post Date] + " is not a date. "
                End If
                
                'Statement Date
                If Not IsDate(tableToCheck![Statement Date]) Then
                    rowError = True
                    errorMsg = errorMsg + "Statement Date: " + tableToCheck![Statement Date] + " is not a date. "
                End If
                
                'Transaction Amt in Original Currency
                If (Not IsNumeric(tableToCheck![Transaction Amt in Original Currency])) Or (Len(tableToCheck![Transaction Amt in Original Currency]) >= 12) Then
                    rowError = True
                    errorMsg = errorMsg + "Transaction Amt in Original Currency: " + tableToCheck![Transaction Amt in Original Currency] + " is not a valid number. "
                End If
                
                'Transaction Code
                If Not isXDigitNumber(tableToCheck![Transaction Cd], 4) Then
                    rowError = True
                    errorMsg = errorMsg + "Transaction Cd: " + tableToCheck![Transaction Cd] + " is not a 4 digit numeric value. "
                End If
                
                'Document Indicator
                If Not isBooleanVal(tableToCheck![Document Indicator]) Then
                    rowError = True
                    errorMsg = errorMsg + "Document Indicator: " + tableToCheck![Document Indicator] + " is not a Y/N value. "
                End If
                
                'Reply Document Indicator
                If Not IsNull(tableToCheck![Reply Document Indicator]) Then
                    If Not isBooleanVal(tableToCheck![Reply Document Indicator]) Then
                          rowError = True
                          errorMsg = errorMsg + "Reply Document Indicator: " + tableToCheck![Reply Document Indicator] + " is not a Y/N value. "
                    End If
                End If
                
                'Chargeback Reason
                If Not validChargebackReason(tableToCheck![Chargeback Reason]) Then
                    rowError = True
                    errorMsg = errorMsg + "Chargeback Reason: " + tableToCheck![Chargeback Reason] + " is not valid. "
                End If
                
                'Country
                If (Len(tableToCheck![Country]) > 3) Or (Len(tableToCheck![Country]) < 2) Or (IsNumeric(tableToCheck![Country])) Then
                    rowError = True
                    errorMsg = errorMsg + "Country: " + tableToCheck![Country] + " is not a valid country code. "
                End If
                
                'Dispute Stage
                If Not validDisputeStage(tableToCheck![Dispute Stage]) Then
                    rowError = True
                    errorMsg = errorMsg + "Dispute Stage: " + tableToCheck![Dispute Stage] + " is not valid. "
                End If
                
                'Dispute Retrieval Reason
                If Not validDisputeRetrievalReason(tableToCheck![Dispute/Retrieval Reason]) Then
                    rowError = True
                    errorMsg = errorMsg + "Dispute/ Retrieval Reason: " + tableToCheck![Dispute/Retrieval Reason] + " is not valid. "
                End If
                
                'Representment Indicator
                If Not IsNull(tableToCheck![Representment Indicator *]) Then
                    If Not isBooleanVal(tableToCheck![Representment Indicator *]) Then
                        rowError = True
                        errorMsg = errorMsg + "Representment Indicator: " + tableToCheck![Representment Indicator *] + " is not a Y/N value. "
                    End If
                End If
                
                'Representment Reason
                If Not IsNull(tableToCheck![Representment Reason *]) Then
                    If Not validRepresentmentReason(tableToCheck![Representment Reason *]) Then
                        rowError = True
                        errorMsg = errorMsg + "Representment Reason: " + tableToCheck![Representment Reason *] + " is not valid. "
                    End If
                End If
                
                'Data Capture Date
                If Not IsNull(tableToCheck![Data Capture Date]) Then
                    If Not IsDate(tableToCheck![Data Capture Date]) Then
                        rowError = True
                        errorMsg = errorMsg + "Data Capture Date: " + tableToCheck![Data Capture Date] + " is not a date. "
                    End If
                End If
                
                'NFR only fields
                If Me.templateT = "Non Fraud Retrieval" Then
                    'Dispute Received Date
                    If Not IsDate(tableToCheck![Dispute Received Date]) Then
                        rowError = True
                        errorMsg = errorMsg + "Dispute Received Date: " + tableToCheck![Transaction Date] + " is not a date. "
                    End If
                    
                    'Case Due Date
                    If Not IsDate(tableToCheck![Case Due Date]) Then
                        rowError = True
                        errorMsg = errorMsg + "Case Due Date: " + tableToCheck![Case Due Date] + " is not a date. "
                    End If
                    
                    'Amex Case ID
                    If Not IsNull(tableToCheck![Amex Case ID #]) Then
                        If Len(tableToCheck![Amex Case ID #]) <> 7 Then
                            rowError = True
                            errorMsg = errorMsg & "Amex Case ID: " & tableToCheck![Amex Case ID #] & " is not valid. "
                        End If
                    End If
                    
                End If
                
                'NFC only fields
                If Me.templateT = "Non Fraud Chargeback" Then
                    
                    'Amex Case ID
                    If Not IsNull(tableToCheck![Amex Case ID # *]) Then
                        If Len(tableToCheck![Amex Case ID # *]) <> 7 Then
                            rowError = True
                            errorMsg = errorMsg & "Amex Case ID: " & tableToCheck![Amex Case ID # *] & " is not valid. "
                        End If
                    End If
                    
                    'Case Opened Date
                    If Not IsDate(tableToCheck![Case Opened Date]) Then
                        rowError = True
                        errorMsg = errorMsg + "Case Opened Date: " + tableToCheck![Case Opened Date] + " is not a date. "
                    End If
                    
                    'Chargeback Amount
                    If (Not IsNumeric(tableToCheck![Chargeback Amount ($)])) Or (Len(tableToCheck![Chargeback Amount ($)]) >= 12) Then
                        rowError = True
                        errorMsg = errorMsg + "Chargeback Amount: " + tableToCheck![Chargeback Amount ($)] + " is not a valid number. "
                    End If
                    
                    'Current Valid Cardmember Number
                    If Not isXDigitNumber(tableToCheck![Current Valid Cardmember Number], 15) Then
                        rowError = True
                        errorMsg = errorMsg + "Current Valid Cardmember Number: " + tableToCheck![Current Valid Cardmember Number] + " is not a 15 digit integer. "
                    End If
                    
                    'NFC Chargeback confirmation fields
                    Dim arraynfc As Variant
                    arraynfc = Array("Case Due Date", "Dispute Received Date")
                    If Not IsNull(tableToCheck![Case Due Date]) Then
                        For Each element In arraynfc
                            If Not IsDate(tableToCheck.Fields(element)) Then
                                errorMsg = errorMsg & element & ": " & tableToCheck.Fields(element) & " is not a date. "
                                rowError = True
                                MsgBox errorMsg, vbOKOnly
                            End If
                        Next element
                        
                        If Not IsNull(tableToCheck![Representment Date]) Then
                        
                            'Representment Date
                            If Not IsDate(tableToCheck![Representment Date]) Then
                                rowError = True
                                errorMsg = errorMsg & "Representment Date: " & tableToCheck![Representment Date] & " is not a date. "
                            End If
                            
                            'Representment Amt
                            If (Not IsNumeric(tableToCheck![Representment Amount])) Or (Len(tableToCheck![Representment Amount]) >= 12) Then
                                rowError = True
                                errorMsg = errorMsg + "Representment Amount: " + tableToCheck![Representment Amount] + " is not a valid number. "
                            End If
                        End If
                    
                    End If
                    
                End If
            
            End If
            
            'FR & FC
            If Me.templateT = "Fraud Chargeback" Or Me.templateT = "Fraud Research" Then
                
                'Transaction Amt
                If (Not IsNumeric(tableToCheck![Transaction Amt])) Or Len(tableToCheck![Transaction Amt]) >= 12 Then
                    rowError = True
                    errorMsg = errorMsg + "Transaction Amt: " + tableToCheck![Transaction Amt] + " is not a valid number. "
                ElseIf CDbl(tableToCheck![Transaction Amt] <= 25) Then
                    rowError = True
                    errorMsg = errorMsg + "Transaction Amt: " + CStr(tableToCheck![Transaction Amt]) + " is not over $25.00. "
                End If
                
                If Not IsNull(tableToCheck![Auth Code *]) Then
                    If Not isXDigitNumber(tableToCheck![Auth Code *], 6) Then
                        rowError = True
                        errorMsg = errorMsg & "Auth Code: " & tableToCheck![Auth Code *] & " is not 6 digit numeric value."
                    End If
                End If
                
                'Fraud Research
                If Me.templateT = "Fraud Research" Then
                
                    If Not IsNull(tableToCheck![Fees]) Then
                    
                        arrayFR = Array("Fees", "Credits", "Finance Charges", "Late Payment")
                        For Each element In arrayFR
                            If (Not IsNumeric(tableToCheck.Fields(element))) Or (Len(tableToCheck.Fields(element)) >= 12) Then
                                errorMsg = errorMsg + element + ": " + tableToCheck.Fields(element) + " is not a valid number. "
                                rowError = True
                            End If
                        Next element
                        
                        'Retier
                        If Not isBooleanVal(tableToCheck![Re-Tier]) Then
                            rowError = True
                            errorMsg = "Re-Tier: " + tableToCheck![Re-Tier] + " is not a Y/N value. "
                        End If
                        
                    End If
                    
                End If
                
                If Me.templateT = "Fraud Chargeback" Then
                
                    'Reason for Claiming Fraud
                    If Not validFraudReason(tableToCheck![Reason for Claiming Fraud?]) Then
                        rowError = True
                        errorMsg = errorMsg + "Reason for Claiming Fraud: " + tableToCheck![Reason for Claiming Fraud?] + " is not valid. "
                    End If
                    
                    
                    'Date Lost/ Stolen
                    If tableToCheck![Reason for Claiming Fraud?] = "Lost" Or tableToCheck![Reason for Claiming Fraud?] = "Stolen" Then
                    
                        If Not IsNull(tableToCheck![Date Lost/Stolen *]) Then
                        
                            If Not IsDate(tableToCheck![Date Lost/Stolen *]) Then
                                rowError = True
                                errorMsg = errorMsg + "Date Lost/Stolen: " + tableToCheck![Date Lost/Stolen *] + " is not a date. "
                            End If
                        Else
                            rowError = True
                            errorMsg = "Date Lost/ Stolen is blank but Reason for Claiming Fraud is Lost/ Stolen. "
                        End If
                    End If
                    
                    'Credits
                    If (Not IsNumeric(tableToCheck![Credits])) Or (Len(tableToCheck![Credits]) >= 12) Then
                        rowError = True
                        errorMsg = errorMsg + "Credits: " + tableToCheck![Credits] + " is not a valid number. "
                    End If
                    
                    'Date reported fraud
                    If Not IsDate(tableToCheck![Date Reported Fraud]) Then
                        rowError = True
                        errorMsg = errorMsg + "Date Reported Fraud: " + tableToCheck![Date Reported Fraud] + " is not a date. "
                    ElseIf (Date - CDate(tableToCheck![Date Reported Fraud]) > 165) Then
                        rowError = True
                        errorMsg = errorMsg + "Date Reported Fraud: " & tableToCheck![Date Reported Fraud] & " is more than 165 days ago. "
                    End If
                    
                    'Chargeback Outcome
                    If Not IsNull(tableToCheck![Chargeback Outcome]) Then
                        If Not validChargebackOutcome(tableToCheck![Chargeback Outcome]) Then
                            rowError = True
                            errorMsg = errorMsg + "Chargeback Outcome: " + tableToCheck![Chargeback Outcome] + " is not valid. "
                        End If
                    End If
                
                End If
            
            End If
            
            If Me.templateT = "Statement Request" Then
                
                'Statement Month
                If Not intBetween(tableToCheck![Statement Month], 12, 1) Then
                    rowError = True
                    errorMsg = errorMsg + "Statement Month: " + tableToCheck![Statement Month] + " is not a valid month. "
                End If
                
                'Statement Year
                If Not isXDigitNumber(tableToCheck![Statement Year], 4) Then
                    rowError = True
                    errorMsg = errorMsg + "Statement Year: " + tableToCheck![Statement Year] + " is not a valid year. "
                End If
                
                'Date CM Statement Sent
                If Not IsNull(tableToCheck![Date CM Statement Sent]) Then
                    If Not IsDate(tableToCheck![Date CM Statement Sent]) Then
                        rowError = True
                        errorMsg = errorMsg + "Date CM Statement Sent: " + tableToCheck![Date CM Statement Sent] + " is not a date. "
                    End If
                End If
                
            End If
            
            If Me.templateT = "CM Application Request" Then
                
                'Affidavit Provided
                If Not isBooleanVal(tableToCheck![Affadavit Provided (Y/N)]) Then
                    rowError = True
                    errorMsg = errorMsg + "Affidavit Provided: " + tableToCheck![Affadavit Provided (Y/N)] + " is not a Y/N value. "
                ElseIf tableToCheck![Affadavit Provided (Y/N)] = "N" Then
                    rowError = True
                    errorMsg = errorMsg + "Affidavit Provided: " + tableToCheck![Affadavit Provided (Y/N)] + ", however Affidavit of Identity is required for application requests. "
                End If
                
                'Date CM Application Sent
                If Not IsNull(tableToCheck![Date CM Application Sent]) Then
                    If Not IsDate(tableToCheck![Date CM Application Sent]) Then
                        rowError = True
                        errorMsg = errorMsg + "Date CM Application Sent: " + tableToCheck![Date CM Application Sent] + " is not a date. "
                    End If
                End If
                
            End If
            
            If Me.templateT = "Cardmember Correspondence" Then
                
                'Reply Document Indicator
                If Not isBooleanVal(tableToCheck![Reply Document Indicator (Y/N)]) Then
                    rowError = True
                    errorMsg = errorMsg + "Reply Document Indicator: " + tableToCheck![Reply Document Indicator (Y/N)] + " is not a Y/N value. "
                End If
                
                'Date Correspondence Sent
                If Not IsDate(tableToCheck![Date Correspondence Sent]) Then
                    rowError = True
                    errorMsg = errorMsg + "Date Correspondence Sent: " + tableToCheck![Date Correspondence Sent] + " is not a date. "
                End If
            
            End If
            
            If Me.templateT = "General Inquiry" Then
                
                'Document Indicator
                If Not isBooleanVal(tableToCheck![Document Indicator (Y/N)]) Then
                    rowError = True
                    errorMsg = errorMsg + "Document Indicator: " + tableToCheck![Document Indicator (Y/N)] + " is not a Y/N value. "
                End If
                
                'Date General Inquiry Response Sent
                If Not IsNull(tableToCheck![Date General Inquiry Response Sent]) Then
                    If Not IsDate(tableToCheck![Date General Inquiry Response Sent]) Then
                        rowError = True
                        errorMsg = errorMsg + "Date General Inquiry Response Sent: " + tableToCheck![Date General Inquiry Response Sent] + " is not a date. "
                    End If
                End If
                
                'Reply Document Indicator
                If Not IsNull(tableToCheck![Reply Document Indicator (Y/N)]) Then
                    If Not isBooleanVal(tableToCheck![Reply Document Indicator (Y/N)]) Then
                        rowError = True
                        errorMsg = errorMsg + "Reply Document Indicator: " & tableToCheck![Reply Document Indicator (Y/N)] & " is not a Y/N value. "
                    End If
                End If
                
            End If
                    
            If Me.templateT = "Fraud Application" Then
                
                'Date Reported Fraud
                If Not IsDate(tableToCheck![Date Reported Fraud]) Then
                    rowError = True
                    errorMsg = errorMsg + "Date Reported Fraud: " + tableToCheck![Date Reported Fraud] + " is not a date. "
                ElseIf (Date - CDate(tableToCheck![Date Reported Fraud])) > 365 Then
                        rowError = True
                        errorMsg = errorMsg + "Date Reported Fraud: " & tableToCheck![Date Reported Fraud] & " is more than 365 days ago. "
                End If
                
                'Tradeline deleted
                If Not IsNull(tableToCheck![Tradeline Deleted]) Then
                    If Not isBooleanVal(tableToCheck![Tradeline Deleted]) Then
                        rowError = True
                        errorMsg = errorMsg + "Tradeline Deleted: " + tableToCheck![Tradeline Deleted] + " is not a Y/N value. "
                    End If
                End If
                
            
            End If
            
            'If there is an error add it to output table
            If rowError Then
                tableError = True
                errorTable.AddNew
                errorTable![row] = counter
                errorTable![error] = errorMsg
                errorTable![Cardmember Number] = tableToCheck![Cardmember Number]
                If Me.templateT <> "Fraud Research" And Me.templateT <> "Fraud Chargeback" And Me.templateT <> "Fraud Application" Then
                    errorTable![Cardmember Name] = tableToCheck![Cardmember Name]
                End If
                errorTable![File Name] = getFileName(Me.fileP)
                errorTable.Update
                tableToCheck.Edit
                tableToCheck![Cardmember Number] = "ERROR"
                tableToCheck.Update
            End If
        End If
            counter = counter + 1
            tableToCheck.MoveNext
            
    Loop
    
    'Close DB data stream
    tableToCheck.Close
    Set tableToCheck = Nothing
    errorTable.Close
    Set errorTable = Nothing
    Set db = Nothing
    
    'If there is an error alert user
    If tableError Then
        alertImportError ("There is a data error in one or more row(s) you are uploading. Please see Import Errors table.")
    End If
    
End Function

'Create a temporary table in Access based on the template type. This will later be integrated with the master DB

Private Function transferToDB()

    'Wipe the contents of the old DB primer
    wipeSQL = "DELETE * FROM [" & Me.templateT & "];"
    CurrentDb.Execute wipeSQL, dbFailOnError
    
    'Transfer info to DB primer
    Dim Row1, Row2 As Integer
    Col1 = "B"
    Col2 = "AG"
    Row1 = 3
    Row2 = 1500
    SheetName = getSheetName(Me.templateT)

    Dim Range As String
    Range = SheetName & "!" & Col1 & Row1 & ":" & Col2 & Row2

    DoCmd.TransferSpreadsheet acImport, 9, Me.templateT.value, Me.fileP.value, True, Range
    
    'Delete empty rows
    If Me.templateT = "Non Fraud Retrieval" Or Me.templateT = "Non Fraud Chargeback" Or Me.templateT = "Fraud Research" Or Me.templateT = "Fraud Chargeback" Then
        delSQL = "DELETE * FROM [" & Me.templateT & "] WHERE ISNULL([Transaction ID])And ISNULL([Cardmember Number]) And ISNULL([SE Number]);"
        CurrentDb.Execute delSQL, dbFailOnError
    ElseIf Me.templateT = "Fraud Application" Then
        delSQL = "DELETE * FROM [" & Me.templateT & "] WHERE ISNULL([Cardmember Number])And ISNULL([Date Reported Fraud]);"
        CurrentDb.Execute delSQL, dbFailOnError
    Else
        delSQL = "DELETE * FROM [" & Me.templateT & "] WHERE ISNULL([Cardmember Number])And ISNULL([Cardmember Name]);"
        CurrentDb.Execute delSQL, dbFailOnError
    End If
    
End Function

'Find out which table in the database the data should be uploaded to

Private Function importTODB() As String
    
    Dim tableName As String
    
    If Me.templateT = "Non Fraud Retrieval" Or Me.templateT = "Non Fraud Chargeback" Then
        importTODBNF
        tableName = "Non Fraud Table"
    ElseIf Me.templateT = "Fraud Research" Or Me.templateT = "Fraud Chargeback" Then
        importTODBF
        tableName = "Fraud Table"
    Else
        importTODBMISC (Me.templateT)
        tableName = "Additional Request Table"
    End If
    
    importTODB = tableName
    
End Function

'When given a path name to file, extract file name

Private Function getFileName(path As String) As String
    
    If Right(path, 1) <> "\" And Len(path) > 0 Then
        getFileName = getFileName(Left(path, Len(path) - 1)) + Right(path, 1)
    End If
    
End Function

'Check if value is x digit number

Private Function isXDigitNumber(value As String, digits As Integer) As Boolean
        
    isXDigitNumber = Len(value) = digits And IsNumeric(value)
        
End Function

'Check if value is between high and low

Private Function intBetween(value As String, high As Integer, low As Integer) As Boolean
    
    intBetween = False
    
    If IsNumeric(value) Then
        intVal = CInt(value)
        intBetween = value <= high And value >= low
    End If
    
End Function


'Check if value is Y/N

Private Function isBooleanVal(value As String) As Boolean

    isBooleanVal = False
    If value = "Y" Or value = "N" Then
        isBooleanVal = True
    End If

End Function

'Check if reason to claim fraud is valid (Lost, Stolen, Not Received, Confirmed Identity Takeover, Fraud Application, None of the Above)

Private Function validFraudReason(value As String) As Boolean

    validFraudReason = False
    If value = "Lost" Or value = "Stolen" Or value = "Not Received" Or value = "Confirmed Identity Takeover" Or value = "Fraud Application" Or value = "None of the Above" Then
        validFraudReason = True
    End If
    
End Function

'Check if Dispute Stage is valid

Private Function validDisputeStage(value As String) As Boolean
    
    validDisputeStage = False
    If value = "Retrieval Request" Or value = "Fulfillment" Or value = "Chargeback" Or value = "Chargeback Confirmation" Or value = "Representment" Or value = "Final Chargeback" Then
        validDisputeStage = True
    End If
    
End Function

'Check if Dispute/Retrieval Reason is valid

Private Function validDisputeRetrievalReason(value As String) As Boolean
    
    validDisputeRetrievalReason = False
    If value = "N/A" Or value = "CM Dispute" Or value = "CM Needs for personal records" Or value = "CM No longer disputing" Then
        validDisputeRetrievalReason = True
    End If
    
End Function

'Check if Representment Reason is valid

Private Function validRepresentmentReason(value As String) As Boolean

    validRepresentmentReason = False
    If value = "NA" Or value = "General Invalid Chargeback" Or value = "Credit Previously Issued" Then
        validRepresentmentReason = True
    End If
    
End Function

'Check if Chargeback Reason is valid

Private Function validChargebackReason(value As String) As Boolean

    validChargebackReason = False
    If validChargebackReasonHelper(value) Or value = "Paid through other means" Or value = "Currency Discrepancy" Or value = "Credit/Debit Presentment Error" Or value = "Cancellation of Recurring Goods and Services" Or value = "Not as described or defective merchandise " Or value = "Goods and Services not received" Then
        validChargebackReason = True
    End If
    
End Function

'Helper function for validChargebackReason

Private Function validChargebackReasonHelper(value As String) As Boolean
    
    validChargebackReasonHelper = False
    If value = "NA" Or value = "Request for support not fulfilled" Or value = "Request for support illegible/incomplete" Or value = "Invalid Authorization" Or value = "Incorrect Transaction Amount" Or value = "Multiple Processing" Or value = "Late Presentment" Or value = "Credit not presented " Then
        validChargebackReasonHelper = True
    End If
    
End Function

'Check if Chargeback Outcome is valid

Private Function validChargebackOutcome(value As String) As Boolean
    
    validChargebackOutcome = False
    If value = "Chargeback Processed" Or value = "Cannot Chargeback" Or value = "Ineligible for Chargeback" Or value = "Chargeback Reversed" Then
        validChargebackOutcome = True
    End If
    
End Function

'Alerts users to an error and directs them to Import Errors table

Private Function alertImportError(error As String)

    MsgBox error, vbOKOnly
    DoCmd.Close acTable, "Import Errors", acSaveYes
    DoCmd.OpenTable "Import Errors"

End Function

'Returns the sheet name where the data is in the excel

Private Function getSheetName(template As String)

    If template = "Non Fraud Retrieval" Then
        getSheetName = "Dispute Retrieval"
    ElseIf template = "Non Fraud Chargeback" Then
        getSheetName = "Dispute Chargeback"
    ElseIf template = "Fraud Research" Then
        getSheetName = "Fraud Info"
    ElseIf template = "Fraud Chargeback" Then
        getSheetName = "Fraud Chargeback"
    ElseIf template = "Statement Request" Then
        getSheetName = "Statement"
    ElseIf template = "CM Application Request" Then
        getSheetName = "Applications"
    ElseIf template = "General Inquiry" Then
        getSheetName = "General Inquiries"
    ElseIf template = "Fraud Application" Then
        getSheetName = "Fraud Application"
    Else 'Cardmember Correspondence
        getSheetName = "Cardmember Correspondence"
    End If
        
End Function

'Causes program to sleep for 1 second

Private Function getSleep()
    Dim started As Single: started = Timer
    Do: DoEvents: Loop Until Timer - started >= 0.1
End Function
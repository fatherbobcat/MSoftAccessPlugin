Option Compare Database

Private Sub browseButton_Click()
    
    Const msoFileDialogFilePicker As Long = 3
    Dim f As Object
    Set f = Application.FileDialog(msoFileDialogFilePicker)
    f.AllowMultiSelect = False
    
    If f.Show Then
        fileP = f.SelectedItems(1)
    End If

End Sub

Private Sub browsePull_Click()
    
    Const msoFileDialogFilePicker As Long = 4
    Dim f As Object
    Set f = Application.FileDialog(msoFileDialogFilePicker)
    f.AllowMultiSelect = False
    
    If f.Show Then
        fileP1 = f.SelectedItems(1)
    End If

End Sub


Private Sub export_Click()

    If IsNull(Me.fileP) Then
        MsgBox "Please Enter a File Path!", vbOKOnly
        Exit Sub
    End If
    
    If IsNull(Me.dateCalc) Then
        MsgBox "Please Enter a Week Of Date!", vbOKOnly
        Exit Sub
    End If
    
    updateDaysActive
    
    'export all
    DoCmd.TransferSpreadsheet acExport, , "pull_NF", fileP, True, "Non Fraud Raw"
    DoCmd.TransferSpreadsheet acExport, , "pull_F", fileP, True, "Fraud Raw"
    DoCmd.TransferSpreadsheet acExport, , "pull_AR", fileP, True, "Additional Request Raw"
    
    weekOfCalcDate = CStr(Int((Me.dateCalc - CDate("3/21/2016")) / 7) + 1)
    
    'export week of
    Dim pullArray As Variant
    pullArray = Array("Non Fraud", "Fraud", "Additional Request")
                
    For Each element In pullArray
        Dim strSQL As String
        Dim qdf As DAO.QueryDef
        strSQL = "SELECT * FROM [" & element & " Table] WHERE [Week Opened] = " & weekOfCalcDate & " OR [Week Closed] = " & weekOfCalcDate & ";"
        Set qdf = CurrentDb.CreateQueryDef("tmpQdf", strSQL)
        DoCmd.TransferSpreadsheet acExport, , "tmpQdf", fileP, True, element & " Raw Week Of"
        CurrentDb.QueryDefs.Delete ("tmpQdf")
        qdf.Close
    Next element
        
    MsgBox "Exported to Reporting Dashboard!", vbOKOnly
    
End Sub


Private Function updateDaysActive()

    'Update the Days Active
    updateDatesNF = "UPDATE [Non Fraud Table] SET [Days Active] = Date() - [Date Received from BCUS] WHERE [Status] = ""Open"";"
    CurrentDb.Execute updateDatesNF, dbFailOnError
    updateDatesF = "UPDATE [Fraud Table] SET [Days Active] = Date() - [Date Received from BCUS] WHERE [Status] = ""Open"";"
    CurrentDb.Execute updateDatesF, dbFailOnError
    updateDatesAR = "UPDATE [Additional Request Table] SET [Days Active] = Date() - [Date Received from BCUS] WHERE [Status] = ""Open"";"
    CurrentDb.Execute updateDatesAR, dbFailOnError
    
End Function

Private Sub exportAll_Click()
    
    If IsNull(Me.fileP1) Then
        MsgBox "Please Enter a File Path!", vbOKOnly
    Exit Sub
    End If
    updateDaysActive
    
    exportFile "Non Fraud Data ", "pull_NF"
    exportFile "Fraud Data ", "pull_F"
    exportFile "Additional Requests Data ", "pull_AR"
    
    MsgBox "Successfully exported!", vbOKOnly
    
End Sub

Private Sub exportFRAP_Click()
        
    If IsNull(Me.fileP1) Then
        MsgBox "Please Enter a File Path!", vbOKOnly
    Exit Sub
    End If
    
    updateDaysActive
    exportFile "Fraud Application Data ", "pull_FRAP"
    
    MsgBox "Successfully exported!", vbOKOnly
    
End Sub

Private Sub exportFraud_Click()

    If IsNull(Me.fileP1) Then
        MsgBox "Please Enter a File Path!", vbOKOnly
    Exit Sub
    End If
    
    updateDaysActive
    exportFile "Fraud Data ", "pull_F"
    
    MsgBox "Successfully exported!", vbOKOnly
    
End Sub

Private Sub exportMisc_Click()
    
    If IsNull(Me.fileP1) Then
        MsgBox "Please Enter a File Path!", vbOKOnly
    Exit Sub
    End If
    
    updateDaysActive
    exportFile "Additional Requests Data ", "pull_AR"
    
    MsgBox "Successfully exported!", vbOKOnly
    
End Sub

Private Sub exportNF_Click()
    
    If IsNull(Me.fileP1) Then
        MsgBox "Please Enter a File Path!", vbOKOnly
    Exit Sub
    End If
    
    updateDaysActive
    exportFile "Non Fraud Data ", "pull_NF"
    
    MsgBox "Successfully exported!", vbOKOnly
    
End Sub

Private Sub exportFile(typeFile As String, query As String)
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    newFile = Me.fileP1 & "\" & typeFile & Format(Date, "mmddyyyy") & ".txt"
    Set a = fs.CreateTextFile(newFile, True)
    a.Close
    DoCmd.TransferText acExportDelim, , query, newFile, True

End Sub

Private Sub Command0_Click()
    Dim rs As DAO.Recordset
    Dim sFolder As String
    Dim sFile As String
    Const sReportName = "NameOfYourReport" 'Fill in your name on the report
 
    On Error GoTo Error_Handler
 
    'The folder in which to save the PDFs
    sFolder = Application.CurrentProject.Path & "\"
 
    'Define the Records that you will use to filtered the report with
    Set rs = CurrentDb.OpenRecordset("SELECT ContactID, FirstName FROM Contacts;", dbOpenSnapshot)
    With rs
        If .RecordCount <> 0 Then 'Make sure we have record to generate PDF with
            .MoveFirst
            Do While Not .EOF
                'Build the PDF filename we are going to use to save the PDF with
                sFile = sFolder & Nz(![FirstName], "") & ".pdf"
                'Open the report filtered to the specific record or criteria we want in hidden mode
                DoCmd.OpenReport sReportName, acViewPreview, , "[ContactID]=" & ![ContactID], acHidden
                'Print it out as a PDF
                DoCmd.OutputTo acOutputReport, sReportName, acFormatPDF, sFile, , , , acExportQualityPrint
                'Close the report now that we're done with this criteria
                DoCmd.Close acReport, sReportName
                'If you wanted to create an e-mail and include an individual report, you would do so now
                .MoveNext
            Loop
        End If
    End With
 
    'Open the folder housing the PDF files (Optional)
    Application.FollowHyperlink sFolder
 
Error_Handler_Exit:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Exit Sub
 
Error_Handler:
    If Err.Number <> 2501 Then    'Let's ignore user cancellation of this action!
        MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
               "Error Number: " & Err.Number & vbCrLf & _
               "Error Source: Command0_Click" & vbCrLf & _
               "Error Description: " & Err.Description & _
               Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
               , vbOKOnly + vbCritical, "An Error has Occurred!"
    End If
    Resume Error_Handler_Exit
Set rs = Nothing
rs.Close
End Sub
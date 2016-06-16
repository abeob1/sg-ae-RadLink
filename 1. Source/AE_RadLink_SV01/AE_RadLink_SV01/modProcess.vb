
Module modProcess

    Public Sub Start()

        Dim sFuncName As String = "Start()"
        Dim DirInfo As New System.IO.DirectoryInfo(p_oCompDef.sInboxDir)
        Dim files() As System.IO.FileInfo
        Dim sErrdesc As String = String.Empty

        Try
            files = DirInfo.GetFiles("*.xlsx")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Uploadfiles()", sFuncName)
           
            Uploadfiles(files)
            'Send Error Email if Datable has rows.
            If p_oDtError.Rows.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EmailTemplate_Error()", sFuncName)
                EmailTemplate_Error()
            End If
            p_oDtError.Rows.Clear()

            'Send Success Email if Datable has rows..
            If p_oDtSuccess.Rows.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EmailTemplate_Success()", sFuncName)
                EmailTemplate_Success()
            End If
            p_oDtSuccess.Rows.Clear()

            'Send SMS failure email if datatable has rows.

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in upload", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try

    End Sub

    Private Sub Uploadfiles(ByVal files() As System.IO.FileInfo)

        Dim sFuncName As String = "Uploadfiles()"
        Dim sErrDesc As String = String.Empty
        Dim bIsFilesExist As Boolean = False

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function..", sFuncName)

            p_oDtSuccess = CreateDataTable("FileName", "Status")
            p_oDtError = CreateDataTable("FileName", "Status", "ErrDesc")
  

            For Each File As System.IO.FileInfo In files
               
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File name is : " & File.Name.ToUpper, sFuncName)

                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling UploadDocument()", sFuncName)
                If UploadDocument(File.FullName, File.Name, sErrDesc) <> RTN_SUCCESS Then
                    'Insert Error Description into Table
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
                    AddDataToTable(p_oDtError, File.Name, "Error", sErrDesc)
                    'error condition
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Moving " & File.Name & " to " & p_oCompDef.sSuccessDir, sFuncName)
                    Dim UploadedFileName As String = Mid(File.Name, 1, File.Name.Length - 5) & "_" & Now.ToString("yyyyMMddhhmmss") & ".txt"
                    File.MoveTo(p_oCompDef.sFailDir & "\" & Replace(UploadedFileName, ".txt", ".xlsx"))
                Else

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Moving " & File.Name & " to " & p_oCompDef.sSuccessDir, sFuncName)
                    Dim UploadedFileName As String = Mid(File.Name, 1, File.Name.Length - 5) & "_" & Now.ToString("yyyyMMddhhmmss") & ".txt"
                    File.MoveTo(p_oCompDef.sSuccessDir & "\" & Replace(UploadedFileName, ".txt", ".xlsx"))

                    'Insert Success Notificaiton into Table..
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
                    AddDataToTable(p_oDtSuccess, File.Name, "Success")
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File successfully uploaded" & File.FullName, sFuncName)

                End If

                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Disconnecting from SAP Databases", sFuncName)
                'p_oCompany.Disconnect()
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Disconnected from SAP Databases", sFuncName)
              
            Next File


            If bIsFilesExist = False Then If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No files found to upload in INUPUT Folder.", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Funcation complete successfully.", sFuncName)

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in upload setup", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Sub

    Public Function UploadDocument(ByVal sFileNamePath As String, ByVal sFileName As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "UploadDocument()"
        Dim myfile As New System.IO.FileInfo(sFileName)
        Dim sSheet1 As String = String.Empty
        Dim sSheet2 As String = String.Empty
        Dim sFileType As String = String.Empty
        Dim bIsError As Boolean = False
        Dim sPatientType As String = String.Empty


        Try

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            sFileType = sFileName

            sFileType = Microsoft.VisualBasic.Left(sFileType, 7)


            If UCase(sFileType) = "INVOICE" Then

                Console.WriteLine("Reading Invoice Listing file..")

                sSheet1 = "Sheet1"
                Dim oDv1 As DataView = Nothing

                bIsError = False
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling ReadInvoiceListing", sFuncName)
                ReadInvoiceListing(sFileNamePath, sSheet1, bIsError, oDv1, sErrDesc)

                If bIsError = True Then
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Invalid Invoice Listing Excel Worksheet", sFuncName)
                    WriteToLogFile("Invalid Invoice Listing Excel Worksheet " & sFileName, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If

                Console.WriteLine("Processing Invoice Listing..")

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessInvoiceFile()", sFuncName)
                If oDv1.Count > 1 Then
                    If Not oDv1(1)(0).ToString = String.Empty Then
                        If ProcessInvoiceFile(oDv1, sFileName, sErrDesc) <> RTN_SUCCESS Then
                            Throw New ArgumentException(sErrDesc)
                        End If
                    End If
                End If
            ElseIf UCase(sFileType) = "RECEIPT" Then
                sSheet1 = "Sheet1"
                Dim oDv2 As DataView = Nothing

                bIsError = False
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling ReadBillToClient", sFuncName)
                ReadReceiptListing(sFileNamePath, sSheet1, bIsError, oDv2, sErrDesc)

                If bIsError = True Then
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Invalid Billl To Client Excel Worksheet", sFuncName)
                    WriteToLogFile("Invalid Billl To Client Excel Worksheet " & sFileName, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If


                Console.WriteLine("Processing Receipt Listing..")

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessReceiptFile()", sFuncName)
                If oDv2.Count > 1 Then
                    If Not oDv2(1)(0).ToString = String.Empty Then
                        If ProcessReceiptFile(oDv2, sFileName, sErrDesc) <> RTN_SUCCESS Then
                            Throw New ArgumentException(sErrDesc)
                        End If
                    End If
                End If
            Else
                sErrDesc = "File name is Invalid. Please check the file name ::" & sFileName
                WriteToLogFile(sErrDesc, sFuncName)
                If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with SUCCESS", sFuncName)
            UploadDocument = RTN_SUCCESS

        Catch ex As Exception
            UploadDocument = RTN_ERROR
            If RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with ERROR", sFuncName)
        Finally

        End Try
    End Function









End Module

Imports System.Globalization


Module modInvoice

#Region "Invoice"

    Public Sub ReadInvoiceListing(ByVal sFileName As String, _
                                 ByVal sSheet As String, _
                                 ByRef bIsError As Boolean, _
                                 ByRef dv As DataView, _
                                 ByRef sErrdesc As String)

        Dim iHeaderRow As Integer
        Dim sCardName As String = String.Empty
        Dim sFuncName As String = "ReadInvoiceListing"
        Dim sBatchNo As String = String.Empty
        Dim k As Integer

        iHeaderRow = 0

        dv = GetDataViewFromExcel(sFileName, sSheet)

        If IsNothing(dv) Then
            Exit Sub
        End If


        If dv(iHeaderRow)(0).ToString <> "locationShortName" Then
            sErrdesc = "Invalid Excel file Format - ([locationShortName] not found at Column 1"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(1).ToString <> "e_registrationDate" Then
            sErrdesc = "Invalid Excel file Format - ([e_registrationDate] not found at Column 2"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If


        If dv(iHeaderRow)(2).ToString <> "registrationNumber" Then
            sErrdesc = "Invalid Excel file Format - ([registrationNumber] not found at Column 3"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(3).ToString <> "accessionNumber" Then
            sErrdesc = "Invalid Excel file Format - ([accessionNumber] not found at Column 4"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(4).ToString <> "patientName" Then
            sErrdesc = "Invalid Excel file Format - ([patientName] not found at Column 5"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(5).ToString <> "referringDoctorName" Then
            sErrdesc = "Invalid Excel file Format - ([referringDoctorName] not found at Column 6"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If


        If dv(iHeaderRow)(6).ToString <> "referringClinicName" Then
            sErrdesc = "Invalid Excel file Format - ([referringClinicName] not found at Column 7"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(7).ToString <> "modalityName" Then
            sErrdesc = "Invalid Excel file Format - ([modalityName] not found at Column 8"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(8).ToString <> "examinationDescription" Then
            sErrdesc = "Invalid Excel file Format - ([examinationDescription] not found at Column 9"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(9).ToString <> "listedBillPrice" Then
            sErrdesc = "Invalid Excel file Format - ([listedBillPrice] not found at Column 10"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(10).ToString <> "SalesAmount" Then
            sErrdesc = "Invalid Excel file Format - ([SalesAmount] not found at Column 11"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(11).ToString <> "registrationDate" Then
            sErrdesc = "Invalid Excel file Format - ([registeredDate] not found at Column 12"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(12).ToString <> "userName" Then
            sErrdesc = "Invalid Excel file Format - ([userName] not found at Column 13"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(13).ToString <> "radiographerName" Then
            sErrdesc = "Invalid Excel file Format - ([radiographerName] not found at Column 14"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(14).ToString <> "scanStartTime" Then
            sErrdesc = "Invalid Excel file Format - ([scanStartTime] not found at Column 15"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(15).ToString <> "scanEndTime" Then
            sErrdesc = "Invalid Excel file Format - ([scanEndTime] not found at Column 16"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If


        If dv(iHeaderRow)(16).ToString <> "stenoGrapherName" Then
            sErrdesc = "Invalid Excel file Format - ([stenoGrapherName] not found at Column 17"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(17).ToString <> "transcriptStartTime" Then
            sErrdesc = "Invalid Excel file Format - ([transcriptStartTime] not found at Column 18"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(18).ToString <> "transcriptEndTime" Then
            sErrdesc = "Invalid Excel file Format - [transcriptEndTime] not found at Column 19"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(19).ToString <> "radiologistName" Then
            sErrdesc = "Invalid Excel file Format - ([radiologistName] not found at Column 20"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(20).ToString <> "signedOffTime" Then
            sErrdesc = "Invalid Excel file Format - [signedOffTime] not found at Column 21"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(21).ToString <> "billToTypeId" Then
            sErrdesc = "Invalid Excel file Format - [billToTypeId] not found at Column 20"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(22).ToString <> "billToId" Then
            sErrdesc = "Invalid Excel file Format - [billToId] not found at Column 21"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(23).ToString <> "customerCode" Then
            sErrdesc = "Invalid Excel file Format - ([customerCode] not found at Column 22"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(24).ToString <> "documentInvoiceNumber" Then
            sErrdesc = "Invalid Excel file Format - [documentInvoiceNumber] not found at Column 23"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(25).ToString <> "receiptNumber" Then
            sErrdesc = "Invalid Excel file Format - [receiptNumber] not found at Column 24"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(26).ToString <> "referringDocumentInvoiceNumber" Then
            sErrdesc = "Invalid Excel file Format - [referringDocumentInvoiceNumber] not found at Column 25"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

    End Sub

    Public Function ProcessInvoiceFile(ByVal oDv As DataView, ByVal sFileName As String, _
                                        ByRef sErrdesc As String) As Long

        Dim sFuncName As String = "ProcessInvoiceFile"
        Dim oDtHdr As New DataTable
        Dim oDIComp() As SAPbobsCOM.Company = Nothing
        Dim oDt As DataTable
        Dim oDtMain As DataTable
        Dim sCostCenter As String = String.Empty

        Dim sSQL As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("getting distinct DBs in the file.", sFuncName)

            oDtMain = oDv.Table.Clone
            oDtMain = oDv.Table.Copy

            oDtMain.Columns.Add("DBCode", GetType(String))
            oDtMain.Columns.Add("Type", GetType(String))

            oDtMain.Rows(0).Delete()
            oDtMain.AcceptChanges()

            For Each row As DataRow In oDtMain.Rows
                row.Item("DBCode") = Microsoft.VisualBasic.Left(row.Item(24).ToString, 2).Trim
                row.Item("Type") = Microsoft.VisualBasic.Mid(row.Item(24).ToString, 3, 2).Trim
            Next

            oDt = oDtMain.DefaultView.ToTable(True, "DBCode")

            ReDim oDIComp(oDt.Rows.Count)

            For j As Integer = 0 To oDt.Rows.Count - 1

                oDIComp(j) = New SAPbobsCOM.Company



                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
                If ConnectToTargetCompany(oDIComp(j), oDt.Rows(j).Item(0).ToString, sCostCenter, sErrdesc) <> RTN_SUCCESS Then
                    Throw New ArgumentException(sErrdesc)
                End If

                Console.WriteLine("Successfully connected")

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting transaction on company database " & oDIComp(j).CompanyDB, sFuncName)
                oDIComp(j).StartTransaction()

                Dim oInvoiceDT As DataTable
                oInvoiceDT = oDtMain.DefaultView.ToTable(True, "DBCode", "Type", "F25")

                For Each row As DataRow In oInvoiceDT.Rows
                    'Create A/R Invoice
                    Dim DBRows() As DataRow = oDtMain.Select("DBCode='" & oDt.Rows(j).Item(0).ToString & "' and Type='IN' and F25='" & row.Item(2).ToString & "'")
                    If DBRows.Length > 0 Then

                        Console.WriteLine("Creating A/R Invoice - Document Invoice No.:: " & row.Item(2).ToString)

                        If AddARInvoice(DBRows, sFileName, sCostCenter, oDIComp(j), sErrdesc) <> RTN_SUCCESS Then
                            Throw New ArgumentException(sErrdesc)
                        End If

                        Console.WriteLine("Successfully created A/R Invoice.")

                    End If

                    'Create A/R CreditNote
                    Dim CNRows() As DataRow = oDtMain.Select("DBCode='" & oDt.Rows(j).Item(0).ToString & "' and Type='CN' and F25='" & row.Item(2).ToString & "'")
                    If CNRows.Length > 0 Then
                        Console.WriteLine("Creating A/R Credit Note - Document Invoice No.:: " & row.Item(2).ToString)
                        If AddARCreditNote(CNRows, sFileName, sCostCenter, oDIComp(j), sErrdesc) <> RTN_SUCCESS Then
                            Throw New ArgumentException(sErrdesc)
                        End If
                        Console.WriteLine("Successfully created A/R Credit Note.")
                    End If

                Next

            Next

            Console.WriteLine("Commit all transaction...")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit all transaction..........", sFuncName)
            For lCounter As Integer = 0 To UBound(oDIComp)
                If Not oDIComp(lCounter) Is Nothing Then
                    If oDIComp(lCounter).Connected = True Then
                        If oDIComp(lCounter).InTransaction = True Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit transaction on company database " & oDIComp(lCounter).CompanyDB, sFuncName)
                            oDIComp(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        End If
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDIComp(lCounter).CompanyDB, sFuncName)
                        oDIComp(lCounter).Disconnect()
                        oDIComp(lCounter) = Nothing
                    End If
                End If
            Next

            Console.WriteLine(".................. COMPLETED .............")

            ProcessInvoiceFile = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion completed successfully.", sFuncName)

        Catch ex As Exception
            Console.WriteLine(".................. ERROR .............")
            For lCounter As Integer = 0 To UBound(oDIComp)
                If Not oDIComp(lCounter) Is Nothing Then
                    If oDIComp(lCounter).Connected = True Then
                        If oDIComp(lCounter).InTransaction = True Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDIComp(lCounter).CompanyDB, sFuncName)
                            oDIComp(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDIComp(lCounter).CompanyDB, sFuncName)
                        oDIComp(lCounter).Disconnect()
                        oDIComp(lCounter) = Nothing
                    End If
                End If
            Next



            ProcessInvoiceFile = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in Uploading AR File", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Function

    Private Function AddARInvoice(ByVal oBPRow() As DataRow, _
                                  ByVal sFileName As String, _
                                  ByVal sCostCenter As String, _
                                  ByVal oCompany As SAPbobsCOM.Company, _
                                  ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim oDoc As SAPbobsCOM.Documents
        Dim sSQL As String = String.Empty
        Dim oRS As SAPbobsCOM.Recordset = Nothing
        Dim iCnt As Integer
        Dim lRetCode As Long
        Dim lErrCode As Long
        Dim oDs As DataSet
        Dim sTaxcode As String = String.Empty
        Try

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI ARI Object", sFuncName)
            oDoc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sSQL = "Select T0.""CardCode"",IFNULL(""ECVatGroup"",'') AS ""ECVatGroup"" from OINV T0 INNER JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" WHERE T0.""NumAtCard""='" & oBPRow(0).Item(24).ToString & "'"
            oDs = ExecuteSQLQuery(sSQL, oCompany.CompanyDB)

            If oBPRow(0).Item(24).ToString = "JPIN21107" Then
                MsgBox(oBPRow(0).Item(24).ToString)
            End If

            If oDs.Tables(0).Rows.Count > 0 Then
                sTaxcode = oDs.Tables(0).Rows(0).Item("ECVatGroup").ToString
                sErrDesc = oCompany.CompanyDB & "- " & "Invoice Document NO ::" & oBPRow(0).Item(24).ToString & " already exist in SAP."
                Console.WriteLine(sErrDesc)
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            If oBPRow(0).Item(21) = 2 Then
                sCardCode = p_oCompDef.sCustomerCode
            Else
                sCardCode = "C" & oBPRow(0).Item(23).ToString
            End If

            oDoc.CardCode = sCardCode
            oDoc.DocDate = CDate(oBPRow(0).Item(1))
            oDoc.TaxDate = CDate(oBPRow(0).Item(1))
            oDoc.NumAtCard = oBPRow(0).Item(24).ToString
            oDoc.UserFields.Fields.Item("U_Source").Value = sFileName

            For Each row As DataRow In oBPRow
                iCnt += 1
                If iCnt > 1 Then
                    oDoc.Lines.Add()
                End If
                oDoc.Lines.ItemCode = row.Item(7).ToString
                oDoc.Lines.Quantity = 1
                oDoc.Lines.UnitPrice = row.Item(10)
                oDoc.Lines.COGSCostingCode2 = sCostCenter
                oDoc.Lines.CostingCode2 = sCostCenter

                If Not sTaxcode = String.Empty Then
                    oDoc.Lines.VatGroup = sTaxcode
                End If
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding ARI.", sFuncName)
            lRetCode = oDoc.Add
            If lRetCode <> 0 Then
                oCompany.GetLastError(lErrCode, sErrDesc)
                sErrDesc = oCompany.CompanyDB & "::" & sErrDesc
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding API failed.", sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion Completed Succesfully.", sFuncName)
            AddARInvoice = RTN_SUCCESS

        Catch ex As Exception
            AddARInvoice = RTN_ERROR
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function

    Private Function AddARCreditNote(ByVal oCRRow() As DataRow, _
                                      ByVal sFileName As String, _
                                      ByVal sCostCenter As String, _
                                      ByVal oCompany As SAPbobsCOM.Company, _
                                      ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim oDoc As SAPbobsCOM.Documents
        Dim sSQL As String = String.Empty
        Dim oRS As SAPbobsCOM.Recordset = Nothing
        Dim iCnt As Integer
        Dim lRetCode As Long
        Dim lErrCode As Long
        Dim oDs As DataSet
        Dim sDocType As String = String.Empty

        Try

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI AR Credit Note Object", sFuncName)
            oDoc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
            oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'sSQL = "SELECT T0.""CardCode"",T0.""DocNum"", T0.""NumAtCard"", T1.""DocEntry"", T1.""LineNum"", T1.""ItemCode"",T1.""VatGroup"" FROM OINV T0  INNER JOIN INV1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" " & _
            '       " WHERE T0.""NumAtCard"" ='" & oCRRow(0).Item(26).ToString & "' order by T1.""LineNum"" "

            sSQL = "SELECT T0.""CardCode"",T0.""DocNum"", T0.""NumAtCard"",T0.""DocType"", T1.""DocEntry"", T1.""LineNum"", T1.""ItemCode"",T1.""Dscription"",T1.""VatGroup"",T1.""AcctCode"" " & _
                   " FROM OINV T0  INNER JOIN INV1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" " & _
                   " WHERE T0.""NumAtCard"" ='" & oCRRow(0).Item(26).ToString & "' AND T0.""DocStatus"" = 'O' AND T0.""CANCELED"" = 'N' order by T1.""LineNum"" "
            oDs = ExecuteSQLQuery(sSQL, oCompany.CompanyDB)


            If oDs.Tables(0).Rows.Count > 0 Then
                oDoc.CardCode = oDs.Tables(0).Rows(0).Item("CardCode").ToString
                oDoc.DocDate = CDate(oCRRow(0).Item(1))
                oDoc.TaxDate = CDate(oCRRow(0).Item(1))
                oDoc.NumAtCard = oCRRow(0).Item(24).ToString
                sDocType = oDs.Tables(0).Rows(0).Item("DocType").ToString
                oDoc.UserFields.Fields.Item("U_Source").Value = sFileName

                If sDocType = "I" Then  '**************ITEM INVOICE
                    oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
                    For Each row As DataRow In oCRRow
                        iCnt += 1
                        If iCnt > 1 Then
                            oDoc.Lines.Add()
                        End If
                        Dim dAmount As Double = 0.0
                        dAmount = row.Item(10)
                        oDoc.Lines.ItemCode = row.Item(7).ToString
                        oDoc.Lines.UnitPrice = Math.Abs(dAmount)
                        oDoc.Lines.COGSCostingCode2 = sCostCenter

                        Dim DBRows() As DataRow = oDs.Tables(0).Select("ItemCode='" & row.Item(7).ToString & "'")
                        If DBRows.Length > 0 Then
                            oDoc.Lines.BaseEntry = DBRows(0).Item("DocEntry")
                            oDoc.Lines.BaseLine = DBRows(0).Item("LineNum")
                            oDoc.Lines.BaseType = 13
                            oDoc.Lines.VatGroup = DBRows(0).Item("VatGroup")
                            oDs.Tables(0).Rows(0).Delete()
                            oDs.Tables(0).AcceptChanges()
                        End If

                    Next
                Else '**************SERVICE INVOICE 
                    oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                    Dim dAmount As Double = 0.0
                    For Each row As DataRow In oCRRow
                        dAmount = dAmount + CDbl(row.Item(10))
                    Next
                    oDoc.Lines.ItemDescription = oDs.Tables(0).Rows(0).Item("Dscription").ToString
                    oDoc.Lines.AccountCode = oDs.Tables(0).Rows(0).Item("AcctCode").ToString
                    oDoc.Lines.UnitPrice = Math.Abs(dAmount)
                    oDoc.Lines.COGSCostingCode2 = sCostCenter

                    oDoc.Lines.BaseEntry = oDs.Tables(0).Rows(0).Item("DocEntry").ToString 'DBRows(0).Item("DocEntry")
                    oDoc.Lines.BaseLine = oDs.Tables(0).Rows(0).Item("LineNum").ToString 'DBRows(0).Item("LineNum")
                    oDoc.Lines.BaseType = 13
                    oDoc.Lines.VatGroup = oDs.Tables(0).Rows(0).Item("VatGroup").ToString ''DBRows(0).Item("VatGroup")
                End If
            Else
                sErrDesc = "Invoice not found/Check invoice for Ref.No " & oCRRow(0).Item(26).ToString
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding AR CR.", sFuncName)
            lRetCode = oDoc.Add
            If lRetCode <> 0 Then
                oCompany.GetLastError(lErrCode, sErrDesc)
                sErrDesc = oCompany.CompanyDB & "::" & sErrDesc
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding AP CR failed.", sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion Completed Succesfully.", sFuncName)
            AddARCreditNote = RTN_SUCCESS

        Catch ex As Exception
            AddARCreditNote = RTN_ERROR
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function

    Private Function AddARCreditNote_Backup(ByVal oCRRow() As DataRow, _
                                      ByVal sCostCenter As String, _
                                      ByVal oCompany As SAPbobsCOM.Company, _
                                      ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim oDoc As SAPbobsCOM.Documents
        Dim sSQL As String = String.Empty
        Dim oRS As SAPbobsCOM.Recordset = Nothing
        Dim iCnt As Integer
        Dim lRetCode As Long
        Dim lErrCode As Long
        Dim oDs As DataSet
        Dim sDocType As String = String.Empty

        Try

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI AR Credit Note Object", sFuncName)
            oDoc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
            oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sSQL = "SELECT T0.""CardCode"",T0.""DocNum"", T0.""NumAtCard"", T1.""DocEntry"", T1.""LineNum"", T1.""ItemCode"",T1.""VatGroup"" FROM OINV T0  INNER JOIN INV1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" " & _
                   " WHERE T0.""NumAtCard"" ='" & oCRRow(0).Item(26).ToString & "' order by T1.""LineNum"" "

            oDs = ExecuteSQLQuery(sSQL, oCompany.CompanyDB)


            If oDs.Tables(0).Rows.Count > 0 Then
                oDoc.CardCode = oDs.Tables(0).Rows(0).Item("CardCode").ToString
                oDoc.DocDate = CDate(oCRRow(0).Item(1))
                oDoc.TaxDate = CDate(oCRRow(0).Item(1))
                oDoc.NumAtCard = oCRRow(0).Item(24).ToString
               
                For Each row As DataRow In oCRRow
                    iCnt += 1
                    If iCnt > 1 Then
                        oDoc.Lines.Add()
                    End If

                    oDoc.Lines.ItemCode = row.Item(7).ToString
                    oDoc.Lines.UnitPrice = -1 * row.Item(10)
                    oDoc.Lines.COGSCostingCode2 = sCostCenter

                    Dim DBRows() As DataRow = oDs.Tables(0).Select("ItemCode='" & row.Item(7).ToString & "'")
                    If DBRows.Length > 0 Then
                        oDoc.Lines.BaseEntry = DBRows(0).Item("DocEntry")
                        oDoc.Lines.BaseLine = DBRows(0).Item("LineNum")
                        oDoc.Lines.BaseType = 13
                        oDoc.Lines.VatGroup = DBRows(0).Item("VatGroup")
                        oDs.Tables(0).Rows(0).Delete()
                        oDs.Tables(0).AcceptChanges()
                    End If

                Next
            Else
                sErrDesc = "Invoice not found/Check invoice for Ref.No " & oCRRow(0).Item(26).ToString
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding AR CR.", sFuncName)
            lRetCode = oDoc.Add
            If lRetCode <> 0 Then
                oCompany.GetLastError(lErrCode, sErrDesc)
                sErrDesc = oCompany.CompanyDB & "::" & sErrDesc
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding AP CR failed.", sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion Completed Succesfully.", sFuncName)
            AddARCreditNote_Backup = RTN_SUCCESS

        Catch ex As Exception
            AddARCreditNote_Backup = RTN_ERROR
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function

#End Region

#Region "Receipt"

    Public Sub ReadReceiptListing(ByVal sFileName As String, _
                                 ByVal sSheet As String, _
                                 ByRef bIsError As Boolean, _
                                 ByRef dv As DataView, _
                                 ByRef sErrdesc As String)

        Dim iHeaderRow As Integer
        Dim sCardName As String = String.Empty
        Dim sFuncName As String = "ReadReceiptListing"
        Dim sBatchNo As String = String.Empty
        Dim k As Integer

        iHeaderRow = 0

        dv = GetDataViewFromExcel(sFileName, sSheet)

        If IsNothing(dv) Then
            Exit Sub
        End If


        If dv(iHeaderRow)(0).ToString <> "locationId" Then
            sErrdesc = "Invalid Excel file Format - ([locationId] not found at Column 1"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(1).ToString <> "receiptDate" Then
            sErrdesc = "Invalid Excel file Format - ([receiptDate] not found at Column 2"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If


        If dv(iHeaderRow)(2).ToString <> "invoiceBillToType" Then
            sErrdesc = "Invalid Excel file Format - ([invoiceBillToType] not found at Column 3"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(3).ToString <> "invoiceBillTo" Then
            sErrdesc = "Invalid Excel file Format - ([invoiceBillTo] not found at Column 4"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(4).ToString <> "receiptNumber" Then
            sErrdesc = "Invalid Excel file Format - ([receiptNumber] not found at Column 5"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(5).ToString <> "status" Then
            sErrdesc = "Invalid Excel file Format - ([status] not found at Column 6"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(6).ToString <> "paymentModeDescription" Then
            sErrdesc = "Invalid Excel file Format - ([paymentModeDescription] not found at Column 7"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(7).ToString <> "paymentModeReference" Then
            sErrdesc = "Invalid Excel file Format - ([paymentModeReference] not found at Column 8"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(8).ToString <> "paymentPaymentAmount" Then
            sErrdesc = "Invalid Excel file Format - ([paymentPaymentAmount] not found at Column 9"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(9).ToString <> "amountSign" Then
            sErrdesc = "Invalid Excel file Format - ([amountSign] not found at Column 10"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(10).ToString <> "amountPaid" Then
            sErrdesc = "Invalid Excel file Format - ([amountPaid] not found at Column 11"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(11).ToString <> "toPerson" Then
            sErrdesc = "Invalid Excel file Format - ([toPerson] not found at Column 12"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(12).ToString <> "invoiceDate" Then
            sErrdesc = "Invalid Excel file Format - ([invoiceDate] not found at Column 13"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(13).ToString <> "invoiceNumber" Then
            sErrdesc = "Invalid Excel file Format - ([invoiceNumber] not found at Column 14"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(14).ToString <> "createdDate" Then
            sErrdesc = "Invalid Excel file Format - ([createdDate] not found at Column 15"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(15).ToString <> "netSalesAmount" Then
            sErrdesc = "Invalid Excel file Format - ([netSalesAmount] not found at Column 16"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If


        If dv(iHeaderRow)(16).ToString <> "balanceAmount" Then
            sErrdesc = "Invalid Excel file Format - ([balanceAmount] not found at Column 17"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(17).ToString <> "creditAmount" Then
            sErrdesc = "Invalid Excel file Format - ([creditAmount] not found at Column 18"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(18).ToString <> "refundAmount" Then
            sErrdesc = "Invalid Excel file Format - [refundAmount] not found at Column 19"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(19).ToString <> "locationName" Then
            sErrdesc = "Invalid Excel file Format - ([locationName] not found at Column 20"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

        If dv(iHeaderRow)(20).ToString <> "registrationNumber" Then
            sErrdesc = "Invalid Excel file Format - [registrationNumber] not found at Column 21"
            WriteToLogFile(False, sErrdesc)
            bIsError = True
        End If

    End Sub

    Public Function ProcessReceiptFile(ByVal oDv As DataView, ByVal sFileName As String, _
                                       ByRef sErrdesc As String) As Long

        Dim sFuncName As String = "ProcessReceiptFile"
        Dim oDtHdr As New DataTable
        Dim oDIComp() As SAPbobsCOM.Company = Nothing
        Dim oDt As DataTable
        Dim oDtMain As DataTable
        Dim sCostCenter As String = String.Empty

        Dim sSQL As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("getting distinct DBs in the file.", sFuncName)

            oDtMain = oDv.Table.Clone
            oDtMain = oDv.Table.Copy

            oDtMain.Columns.Add("DBCode", GetType(String))

            oDtMain.Rows(0).Delete()
            oDtMain.AcceptChanges()

            For Each row As DataRow In oDtMain.Rows
                row.Item("DBCode") = Microsoft.VisualBasic.Left(row.Item(4).ToString, 2).Trim
            Next

            oDt = oDtMain.DefaultView.ToTable(True, "DBCode")

            ReDim oDIComp(oDt.Rows.Count)

            For j As Integer = 0 To oDt.Rows.Count - 1

                oDIComp(j) = New SAPbobsCOM.Company

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
                If ConnectToTargetCompany(oDIComp(j), oDt.Rows(j).Item(0).ToString, sCostCenter, sErrdesc) <> RTN_SUCCESS Then
                    Throw New ArgumentException(sErrdesc)
                End If

                Console.WriteLine("Successfully connected")

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting transaction on company database " & oDIComp(j).CompanyDB, sFuncName)
                oDIComp(j).StartTransaction()

                Dim oReceiptDT As DataTable
                'oReceiptDT = oDtMain.DefaultView.ToTable(True, "DBCode", "F5")

                Dim oDvNew As DataView = New DataView(oDtMain)
                oReceiptDT = oDvNew.Table.DefaultView.ToTable(True, "DBCode", "F5")
                For i As Integer = 0 To oReceiptDT.Rows.Count - 1
                    If Not (oReceiptDT.Rows(i).Item(0).ToString.Trim() = String.Empty Or oReceiptDT.Rows(i).Item(0).ToString.ToUpper().Trim() = "DBCODE") Then
                        If oDt.Rows(j).Item(0).ToString = oReceiptDT.Rows(i).Item(0).ToString.Trim() Then
                            oDvNew.RowFilter = "DBCode ='" & oReceiptDT.Rows(i).Item(0).ToString.Trim() & "' and F5 = '" & oReceiptDT.Rows(i).Item(1).ToString.Trim() & "' "
                            If oDvNew.Count > 0 Then
                                Console.WriteLine("Creating Incoming Payment - Receipt No.:: " & oReceiptDT.Rows(i).Item(1).ToString)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateIncomingPayment()", sFuncName)
                                Dim oDt_Payment As DataTable
                                oDt_Payment = oDvNew.ToTable
                                Dim oDv_Payment As DataView = New DataView(oDt_Payment)
                                If CreateIncomingPayment(oDv_Payment, oDIComp(j), sFileName, sErrdesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrdesc)
                                Console.WriteLine("Successfully created Incoming Payment.")
                            End If
                        End If
                    End If
                Next


            Next
            Console.WriteLine("Commit all transaction...")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit all transaction..........", sFuncName)
            For lCounter As Integer = 0 To UBound(oDIComp)
                If Not oDIComp(lCounter) Is Nothing Then
                    If oDIComp(lCounter).Connected = True Then
                        If oDIComp(lCounter).InTransaction = True Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit transaction on company database " & oDIComp(lCounter).CompanyDB, sFuncName)
                            oDIComp(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        End If
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDIComp(lCounter).CompanyDB, sFuncName)
                        oDIComp(lCounter).Disconnect()
                        oDIComp(lCounter) = Nothing
                    End If
                End If
            Next

            Console.WriteLine(".................. COMPLETED .............")

            ProcessReceiptFile = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion completed successfully.", sFuncName)

        Catch ex As Exception
            Console.WriteLine(".................. ERROR .............")
            For lCounter As Integer = 0 To UBound(oDIComp)
                If Not oDIComp(lCounter) Is Nothing Then
                    If oDIComp(lCounter).Connected = True Then
                        If oDIComp(lCounter).InTransaction = True Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDIComp(lCounter).CompanyDB, sFuncName)
                            oDIComp(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDIComp(lCounter).CompanyDB, sFuncName)
                        oDIComp(lCounter).Disconnect()
                        oDIComp(lCounter) = Nothing
                    End If
                End If
            Next

            ProcessReceiptFile = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in Uploading AR File", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Function

    Private Function CreateIncomingPayment_Datarow(ByVal oReceiptRow() As DataRow, _
                                  ByVal oCompany As SAPbobsCOM.Company, _
                                  ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim sNumAtCard As String = String.Empty
        Dim oRS As SAPbobsCOM.Recordset = Nothing
        Dim oPayment As SAPbobsCOM.IPayments
        Dim bIsLineAdded As Boolean = False
        Dim lRetCode As Long
        Dim lErrCode As Long
        Dim oDt As New DataTable
        Dim oDs As New DataSet

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI ARI Object", sFuncName)
            oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oPayment = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)

            For Each row As DataRow In oReceiptRow
                sNumAtCard = row.Item(13).ToString.Trim()

                sSQL = "SELECT ""CardCode"" FROM ""OINV"" WHERE ""NumAtCard"" = '" & sNumAtCard & "' AND ""DocStatus"" = 'O' AND ""CANCELED"" = 'N'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                oRS.DoQuery(sSQL)
                If oRS.RecordCount > 0 Then
                    sCardCode = oRS.Fields.Item("CardCode").Value
                End If

                Dim iIndex As Integer = row.Item(1).ToString.IndexOf(" ")
                Dim sDate As String = row.Item(1).ToString.Substring(0, iIndex)
                Dim dtDocDate As Date = CDate(row.Item(1))

                Dim sPayType As String = row.Item(6).ToString.Trim()
                Dim sPayRef As String = row.Item(7).ToString.Trim()
                Dim dPayAmount As Double = CDbl(row.Item(8).ToString.Trim())
                Dim dAmount As Double = CDbl(row.Item(10).ToString.Trim())
                Dim sDbCode As String = row.Item(21).ToString.Trim()

                Dim sDB As String = String.Empty
                sSQL = "SELECT ""U_Database"" FROM ""@AE_BRANCH"" WHERE UPPER(""U_Prefix"") = '" & sDbCode.ToUpper() & "'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                oDs = ExecuteSQLQuery(sSQL, p_oCompDef.sSAPDBName)
                If oDs.Tables(0).Rows.Count > 0 Then
                    sDB = oDs.Tables(0).Rows(0).Item(0).ToString
                End If

                Dim sType As String = String.Empty
                sSQL = "SELECT ""U_Mode"" FROM ""@AE_PAYMENT"" WHERE UPPER(""U_Method"") = '" & sPayType.ToUpper() & "' AND UPPER(""U_Branch"") = '" & sDB.ToUpper() & "'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                oDs = ExecuteSQLQuery(sSQL, p_oCompDef.sSAPDBName)
                If oDs.Tables(0).Rows.Count > 0 Then
                    sType = oDs.Tables(0).Rows(0).Item(0).ToString
                End If

                Dim sAccount As String = String.Empty
                sSQL = "SELECT ""U_GL_Account"" FROM ""@AE_PAYMENT"" WHERE UPPER(""U_Method"") = '" & sPayType.ToUpper() & "' AND UPPER(""U_Branch"") = '" & sDB.ToUpper() & "'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                oDs = ExecuteSQLQuery(sSQL, p_oCompDef.sSAPDBName)
                If oDs.Tables(0).Rows.Count > 0 Then
                    sAccount = oDs.Tables(0).Rows(0).Item(0).ToString
                End If

                oPayment.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
                oPayment.CardCode = sCardCode
                oPayment.DocDate = dtDocDate
                oPayment.Remarks = row.Item(4).ToString.Trim()

                If sType.ToUpper() <> "CASH" Then
                    oPayment.TransferReference = sPayRef
                End If

                Dim sDocEntry As String = String.Empty
                sSQL = "SELECT ""DocEntry"" FROM ""OINV"" WHERE ""NumAtCard"" = '" & sNumAtCard & "' AND ""CardCode"" = '" & sCardCode & "' AND ""DocStatus"" = 'O' AND ""CANCELED"" = 'N'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                oRS.DoQuery(sSQL)
                If Not (oRS.BoF And oRS.EoF) Then
                    oRS.MoveFirst()
                    Do Until oRS.EoF
                        sDocEntry = oRS.Fields.Item("DocEntry").Value

                        If sDocEntry <> "" Then
                            oPayment.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
                            oPayment.Invoices.DocEntry = sDocEntry
                            oPayment.Invoices.SumApplied = dAmount

                            If sType.ToUpper() = "BANK TRANSFER" Then
                                oPayment.TransferAccount = sAccount
                                oPayment.TransferSum = dAmount
                            ElseIf sType.ToUpper() = "CASH" Then
                                oPayment.CashAccount = sAccount
                                oPayment.CashSum = dAmount
                            ElseIf sType.ToUpper() = "CHEQUE" Then
                                oPayment.Checks.CheckNumber = sPayRef
                                oPayment.Checks.CheckSum = dAmount
                                oPayment.Checks.CheckAccount = sAccount
                            End If
                            bIsLineAdded = True
                            oPayment.Invoices.Add()
                        End If

                        oRS.MoveNext()
                    Loop
                End If

            Next

            If bIsLineAdded = True Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Document", sFuncName)

                lRetCode = oPayment.Add()
                If lRetCode <> 0 Then
                    oCompany.GetLastError(lErrCode, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                Else
                    Dim iDocNo As Integer
                    oCompany.GetNewObjectCode(iDocNo)
                    Console.WriteLine("Document Created Successfully :: " & iDocNo)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oPayment)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion Completed Succesfully.", sFuncName)
            CreateIncomingPayment_Datarow = RTN_SUCCESS
        Catch ex As Exception
            CreateIncomingPayment_Datarow = RTN_ERROR
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        End Try
    End Function

    Private Function CreateIncomingPayment(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByVal sFileName As String, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateIncomingPayment"
        Dim sCardCode As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim sNumAtCard As String = String.Empty
        Dim oRS As SAPbobsCOM.Recordset = Nothing
        Dim oPayment As SAPbobsCOM.IPayments
        Dim bIsLineAdded As Boolean = False
        Dim lRetCode As Long
        Dim lErrCode As Long
        Dim oDt As New DataTable
        Dim oDs As New DataSet

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating DIAPI ARI Object", sFuncName)
            oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oPayment = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)

            Dim oDtGroup As DataTable
            oDtGroup = oDv.Table.DefaultView.ToTable(True, "F14")
            For i As Integer = 0 To oDtGroup.Rows.Count - 1
                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "INVOICENUMBER") Then
                    oDv.RowFilter = "F14 ='" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' "
                    If oDv.Count > 0 Then
                        sNumAtCard = oDtGroup.Rows(i).Item(0).ToString.Trim()

                        Dim sDate As String = oDv(0)(1).ToString.Trim()
                        Dim dtDocDate As Date
                        dtDocDate = CDate(oDv(0)(1))
                        Dim dAmount As Double
                        Dim sDbCode As String = oDv(0)(21).ToString.Trim()
                        Dim dTotal As Double = 0.0

                        For k As Integer = 0 To oDv.Count - 1
                            Try
                                dAmount = CDbl(oDv(k)(8).ToString.Trim())
                            Catch ex As Exception
                                dAmount = 0
                            End Try
                            dTotal = dTotal + dAmount
                        Next

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Num At Card " & sNumAtCard, sFuncName)

                        Dim sDocEntry As String = String.Empty
                        sSQL = "SELECT ""DocEntry"",""CardCode"" FROM ""OINV"" WHERE ""NumAtCard"" = '" & sNumAtCard & "' AND ""DocStatus"" = 'O' AND ""CANCELED"" = 'N'"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                        oRS.DoQuery(sSQL)
                        If Not (oRS.BoF And oRS.EoF) Then
                            oRS.MoveFirst()

                            sCardCode = oRS.Fields.Item("CardCode").Value

                            oPayment.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
                            oPayment.CardCode = sCardCode
                            oPayment.DocDate = dtDocDate
                            oPayment.Remarks = oDv(0)(4).ToString.Trim()
                            oPayment.UserFields.Fields.Item("U_Source").Value = sFileName

                            Do Until oRS.EoF
                                sDocEntry = oRS.Fields.Item("DocEntry").Value

                                If sDocEntry <> "" Then
                                    oPayment.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
                                    oPayment.Invoices.DocEntry = sDocEntry
                                    oPayment.Invoices.SumApplied = dTotal

                                    Dim dTrnsfrAmt, dCashAmt, dCheckAmt As Double
                                    Dim sPayRef As String = String.Empty
                                    For k As Integer = 0 To oDv.Count - 1
                                        Dim iCheckNum As Integer = 0
                                        Dim sPayType As String = oDv(k)(6).ToString.Trim()
                                        Dim sCheckNum As String = String.Empty
                                        sCheckNum = oDv(k)(7).ToString.Trim()
                                        Dim sPayAmount As Double = 0.0
                                        Try
                                            sPayAmount = CDbl(oDv(k)(8))
                                        Catch ex As Exception
                                            sPayAmount = 0.0
                                        End Try
                                        Dim sType As String = String.Empty
                                        Dim sAccount As String = String.Empty
                                        Dim sCountryCode As String = String.Empty
                                        Dim sBankCode As String = String.Empty
                                        Dim sChkAcctNo As String = String.Empty

                                        sSQL = "SELECT ""U_Mode"",""U_GL_Account"",""U_CheckCountryCode"",""U_CheckBankCode"",""U_CheckAccountNo"" FROM ""@AE_PAYMENT"" " & _
                                               " WHERE UPPER(""U_Method"") = '" & sPayType.ToUpper() & "' AND UPPER(""U_Branch"") = '" & oCompany.CompanyDB.ToString.ToUpper & "'"
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                                        oDs = ExecuteSQLQuery(sSQL, p_oCompDef.sSAPDBName)
                                        If oDs.Tables(0).Rows.Count > 0 Then
                                            sType = oDs.Tables(0).Rows(0).Item(0).ToString
                                            sAccount = oDs.Tables(0).Rows(0).Item(1).ToString
                                            sCountryCode = oDs.Tables(0).Rows(0).Item(2).ToString
                                            sBankCode = oDs.Tables(0).Rows(0).Item(3).ToString
                                            sChkAcctNo = oDs.Tables(0).Rows(0).Item(4).ToString
                                        End If

                                        If sType.ToUpper() = "BANK TRANSFER" Then
                                            If sPayRef = String.Empty Then
                                                sPayRef = oDv(k)(7).ToString.Trim()
                                            Else
                                                sPayRef = sPayRef + "," + oDv(k)(7).ToString.Trim()
                                            End If
                                            dTrnsfrAmt = dTrnsfrAmt + sPayAmount
                                            oPayment.TransferAccount = sAccount
                                            oPayment.TransferSum = dTrnsfrAmt
                                            oPayment.TransferReference = sPayRef
                                        ElseIf sType.ToUpper() = "CASH" Then
                                            dCashAmt = dCashAmt + sPayAmount
                                            oPayment.CashAccount = sAccount
                                            oPayment.CashSum = dCashAmt
                                        ElseIf sType.ToUpper() = "CHEQUE" Then
                                            Dim sChq As String = String.Empty
                                            For Each c As Char In sCheckNum
                                                If IsNumeric(c) Then
                                                    If sChq = "" Then
                                                        sChq = c
                                                    Else
                                                        sChq = sChq & c
                                                    End If
                                                End If
                                            Next
                                            Try
                                                iCheckNum = CInt(sChq)
                                            Catch ex As Exception
                                                iCheckNum = 0
                                            End Try

                                            dCheckAmt = dCheckAmt + sPayAmount
                                            oPayment.Checks.DueDate = Date.Now.ToString()
                                            oPayment.Checks.CountryCode = sCountryCode
                                            oPayment.Checks.BankCode = sBankCode
                                            oPayment.Checks.CheckSum = sPayAmount
                                            oPayment.Checks.CheckNumber = iCheckNum
                                            oPayment.CheckAccount = sAccount
                                            oPayment.Checks.Add()

                                        End If
                                    Next

                                    bIsLineAdded = True
                                    oPayment.Invoices.Add()
                                End If

                                oRS.MoveNext()
                            Loop
                        Else
                            sErrDesc = "Invoice not found for the Reference number " & sNumAtCard
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        End If
                    End If
                End If
            Next

            If bIsLineAdded = True Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Document", sFuncName)

                lRetCode = oPayment.Add()
                If lRetCode <> 0 Then
                    oCompany.GetLastError(lErrCode, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                Else
                    Dim iDocNo As Integer
                    oCompany.GetNewObjectCode(iDocNo)
                    Console.WriteLine("Document Created Successfully :: " & iDocNo)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oPayment)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fucntion Completed Succesfully.", sFuncName)
            CreateIncomingPayment = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
            CreateIncomingPayment = RTN_ERROR
        End Try
    End Function

    Public Function GetSingleStringValue(ByVal sSql As String) As String
        Dim sFuncName As String = "GetSingleStringValue"
        Dim oDs As DataSet
        Dim sValue As String = String.Empty

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)

        oDs = ExecuteSQLQuery(sSql)

        If oDs.Tables(0).Rows.Count > 0 Then
            sValue = oDs.Tables(0).Rows(0).Item(0).ToString
        End If

        Return sValue
    End Function

    Public Function GetIntergerValue(ByVal sSql As String) As Integer
        Dim sFuncName As String = "GetSingleStringValue"
        Dim oDs As DataSet
        Dim sValue As Integer

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)

        oDs = ExecuteSQLQuery(sSql)

        If oDs.Tables(0).Rows.Count > 0 Then
            sValue = oDs.Tables(0).Rows(0).Item(0).ToString
        End If

        Return sValue
    End Function

#End Region

End Module

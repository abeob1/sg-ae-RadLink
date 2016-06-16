Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Data.Common
Imports System.Threading

Module modCommon

#Region "Connection Object [Connect to DI Company]"

    Public Function ConnectToCompany(ByRef oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String, Optional ByVal sDBName As String = "") As Long
        ' **********************************************************************************
        '   Function    :   ConnectToCompany()
        '   Purpose     :   This function will be providing to proceed the connectivity of 
        '                   using SAP DIAPI function
        '               
        '   Parameters  :   ByRef oCompany As SAPbobsCOM.Company
        '                       oCompany =  set the SAP DI Company Object
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   SRI
        '   Date        :   October 2013
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim iRetValue As Integer = -1
        Dim iErrCode As Integer = -1
        Try
            sFuncName = "ConnectToCompany()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing the Company Object", sFuncName)
            oCompany = New SAPbobsCOM.Company

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning the representing database name", sFuncName)

            oCompany.Server = p_oCompDef.sServer
            ' oCompany.Server = "192.168.11.35"

            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB

            oCompany.CompanyDB = p_oCompDef.sSAPDBName
            oCompany.UserName = p_oCompDef.sSAPUser
            oCompany.Password = p_oCompDef.sSAPPwd

            oCompany.LicenseServer = p_oCompDef.sLicenceServer

            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English

            oCompany.UseTrusted = False

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the Company Database.", sFuncName)
            iRetValue = oCompany.Connect()

            If iRetValue <> 0 Then
                oCompany.GetLastError(iErrCode, sErrDesc)

                sErrDesc = String.Format("Connection to Database ({0}) {1} {2} {3}", _
                    oCompany.CompanyDB, System.Environment.NewLine, _
                                vbTab, sErrDesc)

                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ConnectToCompany = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ConnectToCompany = RTN_ERROR
        End Try
    End Function

    Public Function GetSystemIntializeInfo(ByRef oCompDef As CompanyDefault, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   GetSystemIntializeInfo()
        '   Purpose     :   This function will be providing to proceed the initializing 
        '                   variable control during the system start-up
        '               
        '   Parameters  :   ByRef oCompDef As CompanyDefault
        '                       oCompDef =  set the Company Default structure
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   SRI
        '   Date        :   October 2013
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim sConnection As String = String.Empty
        Try

            sFuncName = "GetSystemIntializeInfo()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oCompDef.sDBName = String.Empty
            oCompDef.sServer = String.Empty
            oCompDef.iServerLanguage = 3
            oCompDef.iServerType = 9
            oCompDef.sSAPUser = String.Empty
            oCompDef.sSAPPwd = String.Empty
            oCompDef.sSAPDBName = String.Empty

            oCompDef.sInboxDir = String.Empty
            oCompDef.sSuccessDir = String.Empty
            oCompDef.sFailDir = String.Empty
            oCompDef.sLogPath = String.Empty
            oCompDef.sDebug = String.Empty
            oCompDef.sCustomerCode = String.Empty


            oCompDef.p_sSMSUserName = String.Empty
            oCompDef.p_sSMSPassword = String.Empty
            oCompDef.p_sSMSFrom = String.Empty
            oCompDef.p_sGIROSMS = String.Empty
            oCompDef.p_sCheckSMS = String.Empty

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Server")) Then
                oCompDef.sServer = ConfigurationManager.AppSettings("Server")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LicenceServer")) Then
                oCompDef.sLicenceServer = ConfigurationManager.AppSettings("LicenceServer")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPDBName")) Then
                oCompDef.sSAPDBName = ConfigurationManager.AppSettings("SAPDBName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPUserName")) Then
                oCompDef.sSAPUser = ConfigurationManager.AppSettings("SAPUserName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPPassword")) Then
                oCompDef.sSAPPwd = ConfigurationManager.AppSettings("SAPPassword")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBUser")) Then
                oCompDef.sDBUser = ConfigurationManager.AppSettings("DBUser")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBPwd")) Then
                oCompDef.sDBPwd = ConfigurationManager.AppSettings("DBPwd")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DSN")) Then
                oCompDef.sDSN = ConfigurationManager.AppSettings("DSN")
            End If

            ' folder
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("InboxDir")) Then
                oCompDef.sInboxDir = ConfigurationManager.AppSettings("InboxDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SuccessDir")) Then
                oCompDef.sSuccessDir = ConfigurationManager.AppSettings("SuccessDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("FailDir")) Then
                oCompDef.sFailDir = ConfigurationManager.AppSettings("FailDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LogPath")) Then
                oCompDef.sLogPath = ConfigurationManager.AppSettings("LogPath")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CustomerCode")) Then
                oCompDef.sCustomerCode = ConfigurationManager.AppSettings("CustomerCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Debug")) Then
                oCompDef.sDebug = ConfigurationManager.AppSettings("Debug")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailFrom")) Then
                oCompDef.sEmailFrom = ConfigurationManager.AppSettings("EmailFrom")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailTo")) Then
                oCompDef.sEmailTo = ConfigurationManager.AppSettings("EmailTo")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailSubject")) Then
                oCompDef.sEmailSubject = ConfigurationManager.AppSettings("EmailSubject")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPServer")) Then
                oCompDef.sSMTPServer = ConfigurationManager.AppSettings("SMTPServer")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPPort")) Then
                oCompDef.sSMTPPort = ConfigurationManager.AppSettings("SMTPPort")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPUser")) Then
                oCompDef.sSMTPUser = ConfigurationManager.AppSettings("SMTPUser")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPPassword")) Then
                oCompDef.sSMTPPassword = ConfigurationManager.AppSettings("SMTPPassword")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CreditNoteGL")) Then
                oCompDef.sCreditNoteGL = ConfigurationManager.AppSettings("CreditNoteGL")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("NonStockItem")) Then
                oCompDef.sNonStockItem = ConfigurationManager.AppSettings("NonStockItem")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("FFSItemCode")) Then
                oCompDef.sFFSItemCode = ConfigurationManager.AppSettings("FFSItemCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CAPItemCode")) Then
                oCompDef.sCAPItemCode = ConfigurationManager.AppSettings("CAPItemCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("TPAItemCode")) Then
                oCompDef.sTPAItemCode = ConfigurationManager.AppSettings("TPAItemCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("FFSGLCode")) Then
                oCompDef.sFFSGLCode = ConfigurationManager.AppSettings("FFSGLCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CAPGLCode")) Then
                oCompDef.sCAPGLCode = ConfigurationManager.AppSettings("CAPGLCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CAPGLCode")) Then
                oCompDef.sCAPGLCode = ConfigurationManager.AppSettings("CAPGLCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DefaultCostCenter")) Then
                oCompDef.sDefaultCostCenter = ConfigurationManager.AppSettings("DefaultCostCenter")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("ServiceFee")) Then
                oCompDef.dServiceFee = CDbl(ConfigurationManager.AppSettings("ServiceFee"))
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CustBPSeriesName")) Then
                oCompDef.sCustBPSeriesName = ConfigurationManager.AppSettings("CustBPSeriesName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("VenBPSeriesName")) Then
                oCompDef.sVenBPSeriesName = ConfigurationManager.AppSettings("VenBPSeriesName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("ReportDSN")) Then
                oCompDef.sReportDSN = ConfigurationManager.AppSettings("ReportDSN")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("ReportPDFPath")) Then
                oCompDef.sReportPDFPath = ConfigurationManager.AppSettings("ReportPDFPath")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("ReportsPath")) Then
                oCompDef.sReportsPath = ConfigurationManager.AppSettings("ReportsPath")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CheckGLAccount")) Then
                oCompDef.sCheckGLAccount = ConfigurationManager.AppSettings("CheckGLAccount")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("GIROGLAccount")) Then
                oCompDef.sGIROGLAccount = ConfigurationManager.AppSettings("GIROGLAccount")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CheckBankAccount")) Then
                oCompDef.sCheckBankAccount = ConfigurationManager.AppSettings("CheckBankAccount")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CheckBankCode")) Then
                oCompDef.sCheckBankCode = ConfigurationManager.AppSettings("CheckBankCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("GIROGLAccountAIA")) Then
                oCompDef.sGIROGLAccountAIA = ConfigurationManager.AppSettings("GIROGLAccountAIA")
            End If

            'GJ Start
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("GJ_CheckGLAccount")) Then
                oCompDef.sGJ_CheckGLAccount = ConfigurationManager.AppSettings("GJ_CheckGLAccount")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("GJ_GIROGLAccount")) Then
                oCompDef.sGJ_GIROGLAccount = ConfigurationManager.AppSettings("GJ_GIROGLAccount")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("GJ_CheckBankAccount")) Then
                oCompDef.sGJ_CheckBankAccount = ConfigurationManager.AppSettings("GJ_CheckBankAccount")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("GJ_CheckBankCode")) Then
                oCompDef.sGJ_CheckBankCode = ConfigurationManager.AppSettings("GJ_CheckBankCode")
            End If


            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("GJ_FFSGLCode")) Then
                oCompDef.sGJ_FFSGLCode = ConfigurationManager.AppSettings("GJ_FFSGLCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("GJ_CAPGLCode")) Then
                oCompDef.sGJ_CAPGLCode = ConfigurationManager.AppSettings("GJ_CAPGLCode")
            End If

            'GJ End

            'SMS Credential
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMSUserName")) Then
                oCompDef.p_sSMSUserName = ConfigurationManager.AppSettings("SMSUserName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMSPassword")) Then
                oCompDef.p_sSMSPassword = ConfigurationManager.AppSettings("SMSPassword")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMSFrom")) Then
                oCompDef.p_sSMSFrom = ConfigurationManager.AppSettings("SMSFrom")
            End If


            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("GIROSMS")) Then
                oCompDef.p_sGIROSMS = ConfigurationManager.AppSettings("GIROSMS")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CHECKSMS")) Then
                oCompDef.p_sCheckSMS = ConfigurationManager.AppSettings("CHECKSMS")
            End If

            'dbs

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBS_CheckGLAccount")) Then
                oCompDef.sDBS_CheckGLAccount = ConfigurationManager.AppSettings("DBS_CheckGLAccount")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBS_CheckBankAccount")) Then
                oCompDef.sDBS_CheckBankAccount = ConfigurationManager.AppSettings("DBS_CheckBankAccount")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBS_CheckBankCode")) Then
                oCompDef.sDBS_CheckBankCode = ConfigurationManager.AppSettings("DBS_CheckBankCode")
            End If


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            GetSystemIntializeInfo = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            GetSystemIntializeInfo = RTN_ERROR
        End Try
    End Function

#End Region

    Public Function ExecuteSQLQuery(ByVal sQuery As String) As DataSet

        '**************************************************************
        ' Function      : ExecuteQuery
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : Sri
        ' Date          : Nov 2013
        ' Change        :
        '**************************************************************

        Dim sFuncName As String = String.Empty
        'Dim sConstr As String = "DRIVER={HDBODBC32};SERVERNODE={" & p_oCompDef.sServer & "};DSN=" & p_oCompDef.sDSN & ";UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";"

        Dim oCmd As New Odbc.OdbcCommand
        Dim oDs As New DataSet
        Dim oDbProviderFactoryObject As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.Odbc")
        Dim oCon As DbConnection = oDbProviderFactoryObject.CreateConnection()

        Try
            sFuncName = "ExecuteQuery()"
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Fucntion...", sFuncName)
            'oCon.ConnectionString = "DRIVER={HDBODBC};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & " ;SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName & ""
            oCon.ConnectionString = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName

            oCon.Open()
            oCmd.CommandType = CommandType.Text
            oCmd.CommandText = sQuery
            oCmd.Connection = oCon
            oCmd.CommandTimeout = 0
            Dim da As New Odbc.OdbcDataAdapter(oCmd)
            da.Fill(oDs)
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Successfully.", sFuncName)

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            oCon.Dispose()
        End Try
        Return oDs
    End Function

    Public Function ExecuteSQLQuery(ByVal sQuery As String, ByVal sSAPDBName As String) As DataSet

        '**************************************************************
        ' Function      : ExecuteQuery
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : Sri
        ' Date          : Nov 2013
        ' Change        :
        '**************************************************************

        Dim sFuncName As String = String.Empty
        'Dim sConstr As String = "DRIVER={HDBODBC32};SERVERNODE={" & p_oCompDef.sServer & "};DSN=" & p_oCompDef.sDSN & ";UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";"

        Dim oCmd As New Odbc.OdbcCommand
        Dim oDs As New DataSet
        Dim oDbProviderFactoryObject As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.Odbc")
        Dim oCon As DbConnection = oDbProviderFactoryObject.CreateConnection()

        Try
            sFuncName = "ExecuteQuery()"
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Fucntion...", sFuncName)
            ' oCon.ConnectionString = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & " ;SERVERNODE=" & p_oCompDef.sServer & ";CS=" & sSAPDBName & ""
            oCon.ConnectionString = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & sSAPDBName

            oCon.Open()
            oCmd.CommandType = CommandType.Text
            oCmd.CommandText = sQuery
            oCmd.Connection = oCon
            oCmd.CommandTimeout = 0
            Dim da As New Odbc.OdbcDataAdapter(oCmd)
            da.Fill(oDs)
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Successfully.", sFuncName)

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            oCon.Dispose()
        End Try
        Return oDs
    End Function


    Public Function ExecuteProcedureForDataSet(ByVal spName As String, ByVal ParamArray parameters() As SqlParameter) As DataSet

        '**************************************************************
        ' Function      : ExecuteProcedureForDataSet
        ' Purpose       : Execute Procedure and returns Dataset
        ' Parameters    : ByVal spName - string command Text
        '               : Byval parameters ParamArray 
        ' Author        : Sri
        ' Date          : March 07 2008
        ' Change        :
        '**************************************************************

        'p_oCompDef

        'Dim sConstr As String = "Data Source=" & ConfigurationManager.AppSettings("Server") & ";Initial Catalog=" & ConfigurationManager.AppSettings("DBName") & ";User ID=" & ConfigurationManager.AppSettings("SQLUser") & "; Password=" & ConfigurationManager.AppSettings("SQLPwd")

        Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd

        Dim oCon As New SqlConnection(sConstr)
        Dim oCmd As New SqlCommand
        Dim oDs As New DataSet
        Dim sFuncName As String = String.Empty
        Try
            sFuncName = "ExecuteProcedureForDataSet()"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Fucntion...", sFuncName)

            oCmd.CommandType = CommandType.StoredProcedure
            oCmd.CommandText = spName
            oCmd.Connection = oCon
            If parameters.Length > 0 Then
                Dim p As SqlParameter
                For Each p In parameters
                    If Not p Is Nothing Then
                        oCmd.Parameters.Add(p)
                    End If
                Next
            End If
            oCmd.CommandTimeout = 0
            Dim da As New SqlDataAdapter(oCmd)
            da.Fill(oDs)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Successfully.", sFuncName)

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            Throw New Exception(ex.Message)
        Finally
            If Not oCon Is Nothing Then
                oCon.Close()
                oCon.Dispose()
            End If
        End Try
        Return oDs

    End Function

    Public Function ExecuteProcedureForNonQuery(ByVal spName As String, ByVal ParamArray parameters() As SqlParameter) As Integer

        Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd

        Dim oCon As New SqlConnection(sConstr)
        Dim oCmd As New SqlCommand

        Try
            oCmd.CommandType = CommandType.StoredProcedure
            oCmd.CommandText = spName
            oCmd.Connection = oCon
            If oCon.State = ConnectionState.Closed Then
                oCon.Open()
            End If
            If parameters.Length > 0 Then
                Dim p As SqlParameter
                For Each p In parameters
                    If Not p Is Nothing Then
                        oCmd.Parameters.Add(p)
                    End If
                Next
            End If
            oCmd.CommandTimeout = 0
            oCmd.ExecuteNonQuery()

        Catch ex As Exception
            Throw ex
        Finally
            If Not oCon Is Nothing Then
                oCon.Close()
                oCon.Dispose()
            End If
        End Try
    End Function

    Public Function ExecuteSQLNonQuery(ByVal sQuery As String) As Integer

        '**************************************************************
        ' Function      : ExecuteQuery
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : Sri
        ' Date          : Nov 2013
        ' Change        :
        '**************************************************************

        Dim sFuncName As String = String.Empty
        ' Dim sConstr As String = "SERVERNODE={" & p_oCompDef.sServer & "};DSN=" & p_oCompDef.sDSN & ";UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";"
        Dim sConstr As String = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName
        Dim oCon As New Odbc.OdbcConnection(sConstr)
        Dim oCmd As New Odbc.OdbcCommand
        Dim oDs As New DataSet
        Try
            sFuncName = "ExecuteSQLNonQuery()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Fucntion...", sFuncName)

            oCmd.CommandType = CommandType.Text
            oCmd.CommandText = sQuery
            oCmd.Connection = oCon
            If oCon.State = ConnectionState.Closed Then
                oCon.Open()
            End If
            oCmd.CommandTimeout = 0
            oCmd.ExecuteNonQuery()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Successfully.", sFuncName)

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            If Not oCon Is Nothing Then
                oCon.Close()
                oCon.Dispose()
            End If
        End Try

    End Function

    Public Function CreateDataTable(ByVal ParamArray oColumnName() As String) As DataTable
        Dim oDataTable As DataTable = New DataTable()

        Dim oDataColumn As DataColumn

        For i As Integer = LBound(oColumnName) To UBound(oColumnName)
            oDataColumn = New DataColumn()
            oDataColumn.DataType = Type.GetType("System.String")
            oDataColumn.ColumnName = oColumnName(i).ToString
            oDataTable.Columns.Add(oDataColumn)
        Next

        Return oDataTable

    End Function

    Public Sub AddDataToTable(ByVal oDt As DataTable, ByVal ParamArray sColumnValue() As String)
        Dim oRow As DataRow = Nothing
        oRow = oDt.NewRow()
        For i As Integer = LBound(sColumnValue) To UBound(sColumnValue)
            oRow(i) = sColumnValue(i).ToString
        Next
        oDt.Rows.Add(oRow)
    End Sub


    Public Function StartTransaction(ByRef sErrDesc As String) As Long
        ' ***********************************************************************************
        '   Function   :    StartTransaction()
        '   Purpose    :    Start DI Company Transaction
        '
        '   Parameters :    ByRef sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :   Sri
        '   Date       :   29 April 2013
        '   Change     :
        ' ***********************************************************************************
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "StartTransaction()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_oCompany.InTransaction Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Found hanging transaction.Rolling it back.", sFuncName)
                p_oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            p_oCompany.StartTransaction()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            StartTransaction = RTN_SUCCESS
        Catch exc As Exception
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            StartTransaction = RTN_ERROR
        End Try

    End Function

    Public Function RollBackTransaction(ByRef sErrDesc As String) As Long
        ' ***********************************************************************************
        '   Function   :    RollBackTransaction()
        '   Purpose    :    Roll Back DI Company Transaction
        '
        '   Parameters :    ByRef sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    Sri
        '   Date       :    29 April 2013
        '   Change     :
        ' ***********************************************************************************
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "RollBackTransaction()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_oCompany.InTransaction Then
                p_oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No active transaction found for rollback", sFuncName)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            RollBackTransaction = RTN_SUCCESS
        Catch exc As Exception
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            RollBackTransaction = RTN_ERROR
        Finally
            GC.Collect()
        End Try

    End Function

    Public Function CommitTransaction(ByRef sErrDesc As String) As Long
        ' ***********************************************************************************
        '   Function   :    CommitTransaction()
        '   Purpose    :    Commit DI Company Transaction
        '
        '   Parameters :    ByRef sErrDesc As String
        '                       sErrDesc=Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    Sri
        '   Date       :    29 April 2013
        '   Change     :
        ' ***********************************************************************************
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "CommitTransaction()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_oCompany.InTransaction Then
                p_oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No active transaction found for commit", sFuncName)
            End If

            CommitTransaction = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            CommitTransaction = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try

    End Function

    Public Function GetDataViewFromExcel(ByVal CurrFileToUpload As String, ByVal sSheet As String) As DataView

        'Event      :   GetDataViewFromExcel
        'Purpose    :   For reading of CSV file
        'Author     :   Sri 
        'Date       :   22 Nov 2013 

        'Dim sConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CurrFileToUpload & ";Extended Properties=""Excel 8.0;HDR=NO;IMEX=1"""

        Dim sConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CurrFileToUpload & ";Extended Properties=""Excel 12.0;HDR=NO;IMEX=1"""


        Dim objConn As New System.Data.OleDb.OleDbConnection(sConnectionString)
        Dim da As OleDb.OleDbDataAdapter
        Dim dt As DataTable
        Dim dv As DataView
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "GetDataViewFromExcel"
            'Open Data Adapter to Read from Text file
            da = New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [" & sSheet & "$]", objConn)
            dt = New DataTable("ExcelDataTable")

            'Fill dataset using dataadapter
            da.Fill(dt)
            dv = New DataView(dt)
            Return dv

        Catch ex As Exception
            Return Nothing
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try

    End Function

    Public Sub GetMaxCode(ByRef iCode As Integer)
        Dim sSQL As String
        Dim oDS As New DataSet
        sSQL = "select MAX(TO_INT(""Code"")) as code from ""@AI_TB02_PROVIDERS"""
        oDS = ExecuteSQLQuery(sSQL)

        If oDS.Tables(0).Rows.Count > 0 Then
            If Not IsDBNull(oDS.Tables(0).Rows(0).Item(0)) = True Then
                iCode = oDS.Tables(0).Rows(0).Item(0)
            Else
                iCode = 0
            End If
        End If
    End Sub

    Public Function ConnectToTargetCompany(ByRef oCompany As SAPbobsCOM.Company, _
                                             ByVal sDBCode As String, _
                                             ByRef sCostCenter As String, _
                                             ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   ConnectToTargetCompany()
        '   Purpose     :   This function will be providing to proceed the connectivity of 
        '                   using SAP DIAPI function
        '               
        '   Parameters  :   ByRef oCompany As SAPbobsCOM.Company
        '                       oCompany =  set the SAP DI Company Object
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   SRI
        '   Date        :   October 2013
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim iRetValue As Integer = -1
        Dim iErrCode As Integer = -1
        Dim sSQL As String = String.Empty
        Dim oDs As New DataSet
        Dim sSAPUser As String = String.Empty
        Dim sSAPPWd As String = String.Empty
        Dim sDBName As String = String.Empty


        Try
            sFuncName = "ConnectToTargetCompany()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            sSQL = "SELECT * FROM ""@AE_BRANCH""  WHERE ""U_Prefix"" ='" & sDBCode & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSQL, sFuncName)

            oDs = ExecuteSQLQuery(sSQL)

            If oDs.Tables(0).Rows.Count > 0 Then

                sDBName = oDs.Tables(0).Rows(0).Item("U_Database").ToString
                sSAPUser = oDs.Tables(0).Rows(0).Item("U_SAPUSER").ToString
                sSAPPWd = oDs.Tables(0).Rows(0).Item("U_SAPPASSWORD").ToString
                sCostCenter = oDs.Tables(0).Rows(0).Item("U_OcrCode2").ToString

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing the Company Object", sFuncName)
                oCompany = New SAPbobsCOM.Company

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning the representing database name", sFuncName)
                oCompany.Server = p_oCompDef.sServer

                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB

                oCompany.LicenseServer = p_oCompDef.sLicenceServer
                oCompany.CompanyDB = sDBName
                oCompany.UserName = sSAPUser
                oCompany.Password = sSAPPWd

                oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English

                oCompany.UseTrusted = False

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the Company Database.", sFuncName)
                Console.WriteLine("Connecting to Company DB :: " & sDBName)
                iRetValue = oCompany.Connect()

                If iRetValue <> 0 Then
                    oCompany.GetLastError(iErrCode, sErrDesc)

                    sErrDesc = String.Format("Connection to Database ({0}) {1} {2} {3}", _
                        oCompany.CompanyDB, System.Environment.NewLine, _
                                    vbTab, sErrDesc)

                    Throw New ArgumentException(sErrDesc)
                End If
            Else
                sErrDesc = "No Database login information found in COMPANYDATA Table. Please check"
                Throw New ArgumentException(sErrDesc)
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ConnectToTargetCompany = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ConnectToTargetCompany = RTN_ERROR
        End Try
    End Function

    Public Function GetDateTimeValue(ByVal DateString As String) As DateTime
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_StringToDate(DateString).Fields.Item(0).Value
    End Function

End Module

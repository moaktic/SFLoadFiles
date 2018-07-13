#Region "Imports section"
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.ComponentModel
Imports System.Net
Imports System.Collections.Specialized
Imports System.Net.Mail
Imports System.IO
Imports System.Text.RegularExpressions
Imports MySql.Data.MySqlClient

#End Region

Public Class Form1

#Region "Class level Variable declarations"
    Private dbClient As DBHelperClient
    Private dbMain As DBHelper
    Private dsClientMaster As New DataSet
    Private SavedFilePath As String
    Private sqlConMain As MySqlConnection


#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'check if the errorlog file exists...if yes...send email and abort!
        Try
            Dim strAppPath As String = ConfigurationManager.AppSettings("logfilepath")
            If File.Exists(strAppPath & "\" & Format(Date.Now, "MM-dd-yyyy") & ".txt") Then
                Dim strErrorDetails As String = File.ReadAllText(strAppPath & "\" & Format(Date.Now, "MM-dd-yyyy") & ".txt")
                Dim MailObj As MailMessage = New MailMessage(ConfigurationManager.AppSettings("SUOOPRTEMAIL1"), ConfigurationManager.AppSettings("SUOOPRTEMAIL1"), ConfigurationManager.AppSettings("ALERTHOSTNAME") + " ALERT: VCT2 Fax service aborted2", strErrorDetails)
                MailObj.CC.Add(ConfigurationManager.AppSettings("SUOOPRTEMAIL2"))
                MailObj.IsBodyHtml = True
                'Dim smtp As SmtpClient
                'smtp = New SmtpClient("adobe")
                'Dim smtp As New SmtpClient("smtp.mandrillapp.com")
                'smtp.Port = "587"
                'smtp.Credentials = New System.Net.NetworkCredential("kandersson@tic-us.com", "zzzK--drNM-IkyJo78iylg")
                Dim smtp As New SmtpClient("smtp.sparkpostmail.com")
                smtp.Port = "587"
                smtp.Credentials = New System.Net.NetworkCredential("SMTP_Injection", "221cef49b936ec9b3f9596fb295ba05b9da64ef2")
                smtp.Send(MailObj)
                End
            End If
            
            END

        WriteToErrorLog("Started: " & DateTime.Now.ToString, "Info", "", "StartProcess", 1, "")
        'comment the next line if you want to run the job manually.
        sqlConMain = New MySqlConnection(ConfigurationManager.AppSettings("DATA.CONNECTIONSTRINGM"))
        sqlConMain.Open()

        Dim da As New MySqlDataAdapter("SELECT * FROM client_master where active = 1", sqlConMain)
        Dim clientdbold As String = "client_9999"
        Dim clientdb As String = ""
        da.Fill(dsClientMaster)
        For index As Integer = 0 To dsClientMaster.Tables(0).Rows.Count - 1
            clientdb = dsClientMaster.Tables(0).Rows(index)("database_name")
            ConfigurationManager.AppSettings("DATA.CONNECTIONSTRINGC") = ConfigurationManager.AppSettings("DATA.CONNECTIONSTRINGC").Replace(clientdbold, dsClientMaster.Tables(0).Rows(index)("database_name"))
            StartProcess(sender, e)
            clientdbold = clientdb
        Next

        End

        Catch ex As Exception
            Dim trace = New System.Diagnostics.StackTrace(ex, True)
            WriteToErrorLog(ex.Message, "Error", ex.StackTrace, "StartProcess", 1, trace.GetFrame(0).GetFileLineNumber().ToString)
            End
        End Try

    End Sub

    Private Sub StartProcess(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProcessAll.Click
        Try

            'LoadVendor("1","C:\inetpub\wwwroot\SFGit\UploadedFiles\1\1_Vendor File.csv")
            'End

            Dim strFileID As String = ""
            Dim sqlstr As String = ""

            'SavedFilePath = ConfigurationManager.AppSettings("folderpath") & Session("client_id")
            SavedFilePath = ConfigurationManager.AppSettings("folderpath") & "1"

            dbClient = New DBHelperClient
            sqlstr = "SELECT FileID FROM tbluploadfile where FileStatus = 'ReadyToImport'"
            strFileID = dbClient.ExecuteScalar(CommandType.Text, sqlstr)
            If Not strFileID Is Nothing Then
                sqlstr = " Update tbluploadfile set FileStatus = 'ImportingToDB' where FileId = " & strFileID
                dbClient.ExecuteNonQuery(CommandType.Text, sqlstr)

                'loop for files!
                Dim filecount As Integer = Directory.GetFiles(SavedFilePath, strFileID & "_" & "*.csv").Count
                Dim strFileNames() As String
                ReDim strFileNames(filecount - 1)
                strFileNames = Directory.GetFiles(SavedFilePath, strFileID & "_" & "*.csv")
                For index = 0 To filecount - 1
                    If strFileNames(index).ToLower.Contains("vendorfile") Or strFileNames(index).ToLower.Contains("vendor file") Then
                        LoadVendor(strFileID, strFileNames(index))
                    End If
                    If strFileNames(index).ToLower.Contains("invoice to po file") Or strFileNames(index).ToLower.Contains("invoicetopofile") Then
                        LoadInvoiceToPO(strFileID, strFileNames(index))
                    End If
                    If strFileNames(index).ToLower.Contains("pofile") Or strFileNames(index).ToLower.Contains("po file") Then
                        LoadPOFile(strFileID, strFileNames(index))
                    End If
                    If strFileNames(index).ToLower.Contains("invoicedetail") Or strFileNames(index).ToLower.Contains("invoice detail") Then
                        LoadInvoiceDetail(strFileID, strFileNames(index))
                    End If
                    If strFileNames(index).ToLower.Contains("invoiceheaderfile") Or strFileNames(index).ToLower.Contains("invoice header file") Then
                        LoadInvoiceHeader(strFileID, strFileNames(index))
                    End If
                Next

                sqlstr = " Update tbluploadfile set FileStatus = 'FilesUploaded' where FileId = " & strFileID
                dbClient.ExecuteNonQuery(CommandType.Text, sqlstr)

                Dim parmsm(0) As DBHelperClient.Parameters
                parmsm(0) = New DBHelperClient.Parameters("iFileId", strFileID)
                dbClient.ExecuteNonQuery(CommandType.StoredProcedure, "SetFileDashboard", parmsm)

                Dim Errcnt As Integer
                sqlstr = " SELECT COUNT(*) FROM tblfileerrordetails WHERE fileId = " & strFileID & " and ErrType = 'Format Incorrect' "
                Errcnt = dbClient.ExecuteScalar(CommandType.Text, sqlstr)

                If Errcnt > 0 Then
                    sqlstr = " Update tbluploadfile set FileStatus = 'File Errors.' where FileId = " & strFileID
                    dbClient.ExecuteNonQuery(CommandType.Text, sqlstr)
                Else
                    sqlstr = " Update tbluploadfile set FileStatus = 'UploadComplete' where FileId = " & strFileID
                    dbClient.ExecuteNonQuery(CommandType.Text, sqlstr)
                End If

                WriteToErrorLog("Ended: " & DateTime.Now.ToString, "Info", "", "StartProcess", 1, "")
            End If

                       

        Catch ex As Exception
            Dim trace = New System.Diagnostics.StackTrace(ex, True)
            WriteToErrorLog(ex.Message, "Error", ex.StackTrace, "StartProcess", 1, trace.GetFrame(0).GetFileLineNumber().ToString)
            End
        End Try
    End Sub
    Protected Sub LoadVendor(ByVal strFileID As String, ByVal strFileName As String)
        Dim parms(1) As DBHelperClient.Parameters
        Dim sqlstr As String = ""
        Dim MinRowId As Integer = 0
        Try
            sqlstr = "Load Data local Infile '" & (strFileName).Replace("\", "\\") & "' "
            sqlstr = sqlstr + "  into Table tbluploadvendor CHARACTER SET binary Fields Terminated by ',' ENCLOSED BY '""' Lines Terminated by '\n' ignore 4 lines  "
            sqlstr = sqlstr + " (VendorNum, AddressID, Commodity, CommodityDesc, CreatedOn, CreatedBy, "
            sqlstr = sqlstr + " VendorName, AddressLine1, AddressLine2, AddressLine3, AddressLine4, City, State, POBox, "
            sqlstr = sqlstr + " Zip, Country, Active, Telephone, Fax, TaxId, Email, Flex1, Flex2, Flex3, Flex4) "
            dbClient.ExecuteNonQuery(CommandType.Text, sqlstr)


            sqlstr = " Update tbluploadvendor set FileId = " & strFileID & " where FileId IS NULL "
            dbClient.ExecuteNonQuery(CommandType.Text, sqlstr)
            sqlstr = " Select IFNULL(MIN(VMrowid), 0) from tbluploadvendor where FileId = " & strFileID
            MinRowId = dbClient.ExecuteScalar(CommandType.Text, sqlstr)
            sqlstr = " Update tbluploadvendor set LineNum = (VMrowid - " & MinRowId & "  + 1) where FileId = " & strFileID
            dbClient.ExecuteNonQuery(CommandType.Text, sqlstr)
            parms(0) = New DBHelperClient.Parameters("iFileId", strFileID)
            parms(1) = New DBHelperClient.Parameters("stablename", "tbluploadvendor")
            dbClient.ExecuteNonQuery(CommandType.StoredProcedure, "ValidateUpload", parms)
        Catch ex As Exception
            Dim trace = New System.Diagnostics.StackTrace(ex, True)
            WriteToErrorLog(ex.Message, "Error", ex.StackTrace, "LoadVendor", 1, trace.GetFrame(0).GetFileLineNumber().ToString)
            End
        End Try
    End Sub

    Protected Sub LoadInvoiceToPO(ByVal strFileID As String, ByVal strFileName As String)
        Dim parms(1) As DBHelperClient.Parameters
        Dim sqlstr As String = ""
        Dim MinRowId As Integer = 0
        Try
            sqlstr = "Load Data local Infile '" & (strFileName).Replace("\", "\\") & "' "
            sqlstr = sqlstr + "  into Table tbluploadpoinv CHARACTER SET binary Fields Terminated by ',' ENCLOSED BY '""' Lines Terminated by '\n' ignore 1 lines  "
            sqlstr = sqlstr + " (PONumber, POLine, InvoiceId, InvoiceQty, TranCurAmount, Flex1, Flex2, Flex3, Flex4) "
            dbClient.ExecuteNonQuery(CommandType.Text, sqlstr)
            sqlstr = " Update tbluploadpoinv set FileId = " & strFileID & " where FileId IS NULL "
            dbClient.ExecuteNonQuery(CommandType.Text, sqlstr)
            sqlstr = " Select IFNULL(Min(invporowid), 0) from tbluploadpoinv where FileId = " & strFileID
            MinRowId = dbClient.ExecuteScalar(CommandType.Text, sqlstr)
            sqlstr = " Update tbluploadpoinv set LineNum = (invporowid - " & MinRowId & "  + 1) where FileId = " & strFileID
            dbClient.ExecuteNonQuery(CommandType.Text, sqlstr)
            parms(0) = New DBHelperClient.Parameters("iFileId", strFileID)
            parms(1) = New DBHelperClient.Parameters("stablename", "tbluploadpoinv")
            dbClient.ExecuteNonQuery(CommandType.StoredProcedure, "ValidateUpload", parms)

        Catch ex As Exception
            Dim trace = New System.Diagnostics.StackTrace(ex, True)
            WriteToErrorLog(ex.Message, "Error", ex.StackTrace, "LoadInvoiceToPO", 1, trace.GetFrame(0).GetFileLineNumber().ToString)
            End
        End Try
    End Sub

    Protected Sub LoadPOFile(ByVal strFileID As String, ByVal strFileName As String)
        Dim parms(1) As DBHelperClient.Parameters
        Dim sqlstr As String = ""
        Dim MinRowId As Integer = 0
        Try
            sqlstr = "Load Data local Infile '" & (strFileName).Replace("\", "\\") & "' "
            sqlstr = sqlstr + "  into Table tbluploadpo CHARACTER SET binary Fields Terminated by ',' ENCLOSED BY '""' Lines Terminated by '\n' ignore 1 lines  "
            sqlstr = sqlstr + " (PONumber, POLine, POText, PartNum, POPartCommodity, POPartcommodityDesc, "
            sqlstr = sqlstr + " UOM, PricePerUnit, POQty, POAmount, VendorNum, Status, CreatedOn, CreatedBy, Company, "
            sqlstr = sqlstr + " Plant, PlantDesc, PaymentTerm, Flex1, Flex2, Flex3, Flex4) "
            dbClient.ExecuteNonQuery(CommandType.Text, sqlstr)
            sqlstr = " Update tbluploadpo set FileId = " & strFileID & " where FileId IS NULL "
            dbClient.ExecuteNonQuery(CommandType.Text, sqlstr)
            sqlstr = " Select IFNULL(Min(POrowid), 0) from tbluploadpo where FileId = " & strFileID
            MinRowId = dbClient.ExecuteScalar(CommandType.Text, sqlstr)
            sqlstr = " Update tbluploadpo set LineNum = (POrowid - " & MinRowId & "  + 1) where FileId = " & strFileID
            dbClient.ExecuteNonQuery(CommandType.Text, sqlstr)
            parms(0) = New DBHelperClient.Parameters("iFileId", strFileID)
            parms(1) = New DBHelperClient.Parameters("stablename", "tbluploadpo")
            dbClient.ExecuteNonQuery(CommandType.StoredProcedure, "ValidateUpload", parms)
        Catch ex As Exception
            Dim trace = New System.Diagnostics.StackTrace(ex, True)
            WriteToErrorLog(ex.Message, "Error", ex.StackTrace, "LoadPOFile", 1, trace.GetFrame(0).GetFileLineNumber().ToString)
            End
        End Try
    End Sub

    Protected Sub LoadInvoiceDetail(ByVal strFileID As String, ByVal strFileName As String)
        Dim parms(1) As DBHelperClient.Parameters
        Dim sqlstr As String = ""
        Dim MinRowId As Integer = 0
        Try
            sqlstr = "Load Data local Infile '" & (strFileName).Replace("\", "\\") & "' "
            sqlstr = sqlstr + "  into Table tbluploadinvdetail CHARACTER SET binary Fields Terminated by ',' ENCLOSED BY '""' Lines Terminated by '\n' ignore 1 lines  "
            sqlstr = sqlstr + " (InvoiceId, LineNumber, GLNum, GLDesc, CostCenter, "
            sqlstr = sqlstr + " CostCenterDesc, TranCurAmount, InvoiceLineText, Flex1, Flex2, Flex3, Flex4) "
            dbClient.ExecuteNonQuery(CommandType.Text, sqlstr)
            sqlstr = " Update tbluploadinvdetail set FileId = " & strFileID & " where FileId IS NULL "
            dbClient.ExecuteNonQuery(CommandType.Text, sqlstr)
            sqlstr = " Select IFNULL(Min(IDrowid), 0) from tbluploadinvdetail where FileId = " & strFileID
            MinRowId = dbClient.ExecuteScalar(CommandType.Text, sqlstr)
            sqlstr = " Update tbluploadinvdetail set LineNum = (IDrowid - " & MinRowId & "  + 1) where FileId = " & strFileID
            dbClient.ExecuteNonQuery(CommandType.Text, sqlstr)

            parms(0) = New DBHelperClient.Parameters("iFileId", strFileID)
            parms(1) = New DBHelperClient.Parameters("stablename", "tbluploadinvdetail")
            dbClient.ExecuteNonQuery(CommandType.StoredProcedure, "ValidateUpload", parms)
        Catch ex As Exception
            Dim trace = New System.Diagnostics.StackTrace(ex, True)
            WriteToErrorLog(ex.Message, "Error", ex.StackTrace, "LoadVendor", 1, trace.GetFrame(0).GetFileLineNumber().ToString)
            End
        End Try
    End Sub

    Protected Sub LoadInvoiceHeader(ByVal strFileID As String, ByVal strFileName As String)
        Dim parms(1) As DBHelperClient.Parameters
        Dim sqlstr As String = ""
        Dim MinRowId As Integer = 0
        Try
            sqlstr = "Load Data local Infile '" & (strFileName).Replace("\", "\\") & "' "
            sqlstr = sqlstr + "  into Table tbluploadinvheader CHARACTER SET binary Fields Terminated by ',' ENCLOSED BY '""' Lines Terminated by '\n' ignore 1 lines  "
            sqlstr = sqlstr + " (InvoiceId, Company, InvoiceNum, InvoiceDate, InvoiceText, InvoicePaid, "
            sqlstr = sqlstr + " InvoiceVoid, TranCurAmount, Currency, CreatedBy, CreatedOn, VendorNum, VendorAddressId, "
            sqlstr = sqlstr + " PaymentTerm, Flex1, Flex2, Flex3, Flex4) "
            dbClient.ExecuteNonQuery(CommandType.Text, sqlstr)
            sqlstr = " Update tbluploadinvheader set FileId = " & strFileID & " where FileId IS NULL "
            dbClient.ExecuteNonQuery(CommandType.Text, sqlstr)
            sqlstr = " Select IFNULL(Min(IHrowid), 0) from tbluploadinvheader where FileId = " & strFileID
            MinRowId = dbClient.ExecuteScalar(CommandType.Text, sqlstr)
            sqlstr = " Update tbluploadinvheader set LineNum = (IHrowid - " & MinRowId & "  + 1) where FileId = " & strFileID
            dbClient.ExecuteNonQuery(CommandType.Text, sqlstr)
            parms(0) = New DBHelperClient.Parameters("iFileId", strFileID)
            parms(1) = New DBHelperClient.Parameters("stablename", "tbluploadinvheader")
            dbClient.ExecuteNonQuery(CommandType.StoredProcedure, "ValidateUpload", parms)
        Catch ex As Exception
            Dim trace = New System.Diagnostics.StackTrace(ex, True)
            WriteToErrorLog(ex.Message, "Error", ex.StackTrace, "LoadInvoiceHeader", 1, trace.GetFrame(0).GetFileLineNumber().ToString)
            End
        End Try
    End Sub


    Public Sub WriteToErrorLog(ByVal msg As String, ByVal logtype As String, ByVal stkTrace As String, ByVal source As String, ByVal loglevel As Integer, ByVal moreinfo As String)

        Dim fs1 As FileStream

        Dim strAppPath As String = ConfigurationManager.AppSettings("logfilepath")

        If loglevel >= ConfigurationManager.AppSettings("loglevel") Then
            'check and make the directory if necessary; this is set to look in the application
            'folder, you may wish to place the error log in another location depending upon the
            'the user's role and write access to different areas of the file system
            If Not System.IO.Directory.Exists(strAppPath) Then
                System.IO.Directory.CreateDirectory(strAppPath)
            End If

            'check the file size....if bigger than 50 MB then stop logging!
            If File.Exists(strAppPath & "\" & Format(Date.Now, "MM-dd-yyyy") & ".txt") Then
                Dim logfile As New FileInfo(strAppPath & "\" & Format(Date.Now, "MM-dd-yyyy") & ".txt")
                If logfile.Length > CType(ConfigurationManager.AppSettings("MAXLOGFILELENGTHINBYTES"), Long) Then
                    Exit Sub
                End If
            End If
            fs1 = New FileStream(strAppPath & "\" & Format(Date.Now, "MM-dd-yyyy") & ".txt", FileMode.Append, FileAccess.Write)

            'log it
            Dim s1 As StreamWriter = New StreamWriter(fs1)
            s1.Write("Source: " & source & vbCrLf)
            s1.Write("Type: " & logtype & vbCrLf)
            s1.Write("Message: " & msg & vbCrLf)
            s1.Write("StackTrace: " & stkTrace & vbCrLf)
            s1.Write("More info: " & moreinfo & vbCrLf)
            s1.Write("Date/Time: " & DateTime.Now.ToString() & vbCrLf)
            s1.Write("===========================================================================================" & vbCrLf)
            s1.Close()
            fs1.Close()
        End If
        'End If

    End Sub


End Class

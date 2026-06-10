
Imports SAPbobsCOM

Public Class DatabaseCreation

#Region "Declaration"
    Private objUtilities As Utilities
    Dim DBCode As String = "v0.135"
    Dim DBName As String = "v0.135"
    Dim Version As String = "v0.370"
#End Region

#Region "DB Creation Main"
    Public Sub New()
        objUtilities = New Utilities
    End Sub
    Public Function CreateTables() As Boolean
        Try
            objUtilities.CreateTable("TNX_DB", "DBCONFIG(EInvoice)TABLE", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            objUtilities.AddAlphaField("@TNX_DB", "VERSION", "VERSION", 100)
            Dim oRs As SAPbobsCOM.Recordset
            oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery("SELECT * FROM ""@TNX_DB"" where ""U_VERSION"" = '" & Version & "'")
            Dim iDBConfigRecordCount As Integer = oRs.RecordCount
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
            If iDBConfigRecordCount = 0 Then

                Me.LicenceNew()
                objMain.CreateLicenceNewUDO()
                Me.CorpTax()
                objMain.CorpTaxUDO()
                Me.FtaVat()
                objMain.FtaVatUDO()

                Me.OnboardingProcessTable()
                objMain.CreateONBPUDO()
                Me.InvoicePosting()
                objMain.CreateIPUDO()
                Me.InvoicePostingTable()
                objMain.CreateINVFUDO()
                PayPosting()
                objMain.PAYUDO()
                VatReport()
                objMain.VATREPORTUDO()
                Me.CreateAPPROVALSTAGES()
                objMain.CreateAPPROVALSTAGESUDO()
                PayLoadData()
                objMain.PAYLOADUDO()

                Me.CreateDDTable()
                objMain.DDUDO()
                Me.CreateAPPROVALTemplates()
                objMain.CreateApprovalTemplatesUDO()
                'Me.CreateAPPROVALSTAGES()
                'Me.CreateDDTable()
                ' Me.CreateDraftTable()
                'objMain.DraftUDO()
                'Me.CreateMenuTable()
                'objMain.CreateMenuUDO()
                Me.FtaVatcal1()
                objMain.FtaVatcalUDO1()

                Me.CorpTax1()
                objMain.CorpTaxUDO1()
                Me.FtaVat1()
                objMain.FtaVatUDO1()
                Me.LkMster()
                objMain.LkMsterUDO()
                Me.CorporateTaxConfiguration()
                objMain.CTAXConifgUDO()
                Me.CorporateTaxCalculation()
                objMain.CorporateTaxCalculationUDO()
                Me.Corp()
                objMain.CorpUDO()
                Me.Errorlogs()
                objMain.objUtilities.AddDateField("OINV", "EIDate", "EInvoice Date", SAPbobsCOM.BoFldSubTypes.st_None)
                objMain.objUtilities.AddDateField("OINV", "EITime", "EInvoice Time", SAPbobsCOM.BoFldSubTypes.st_Time)
                objMain.objUtilities.AddAlphaField("OINV", "ERRORMSG", "Error messages", 254)
                objMain.objUtilities.AddAlphaField("OINV", "XMLPath", "XML Path", 254)
                objMain.objUtilities.AddAlphaField("OINV", "INVHASH", "ZATCA Invoice Hash", 254)
                objMain.objUtilities.AddAlphaField("OINV", "UUID", "ZATCA UUID", 254)
                objMain.objUtilities.AddAlphaField("OINV", "UUID1", "ZATCA UUID1", 254)
                objMain.objUtilities.AddAlphaField("OINV", "MSG", "ZATCA Message", 254)
                objMain.objUtilities.AddAlphaField("OINV", "STATUS1", "ZATCA Status", 254)
                objMain.objUtilities.AddAlphaMemoField("OINV", "XML", "ZATCA XML ", 254)
                objMain.objUtilities.AddAlphaMemoField("OINV", "QRCODE", "QRCODE", 254)
                objMain.objUtilities.AddAlphaMemoField("OINV", "RESPONSE", "API Response", 254)
                objMain.objUtilities.AddAlphaMemoField("OINV", "ARQT", "API Request", 254)


                objMain.objUtilities.addField("OINV", "PAYMEANS", "Payment Means", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "10,30,42,48,1", "In cash,Credit transfer,Payment to a bank account,Bank card, The Instrument is not defined", "")

                objMain.objUtilities.AddAlphaField("OINV", "STATUS", "ZATCA Status", 254)

                objMain.objUtilities.addField("OACT", "CTAX", "Corporate Tax Include", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "Yes")

                objMain.objUtilities.addField("OALT", "WAAlert", "WhatsApp Alert", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
                objMain.objUtilities.addField("AOB1", "WASent", "Whatsapp Sent", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")

                objMain.objUtilities.AddAlphaField("ORIN", "REASON", "Reason", 254)
                objMain.objUtilities.AddAlphaField("ORIN", "INVNO", "Invoice No", 30)
                objMain.objUtilities.AddAlphaField("ORIN", "UUID", "ZATCA UUID", 254)
                objMain.objUtilities.AddAlphaField("ORIN", "UUID1", "ZATCA UUID1", 254)
                objMain.objUtilities.AddAlphaField("ORIN", "MSG", "ZATCA Message", 254)
                objMain.objUtilities.AddAlphaField("ORIN", "STATUS1", "ZATCA Status", 254)
                objMain.objUtilities.AddAlphaMemoField("ORIN", "XML", "ZATCA XML ", 254)
                objMain.objUtilities.AddAlphaMemoField("ORIN", "QRCODE", "QRCODE", 254)
                objMain.objUtilities.AddAlphaMemoField("ORIN", "RESPONSE", "API Response", 254)
                objMain.objUtilities.addField("ORIN", "PAYMEANS", "Payment Means", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "10,30,42,48,1", "In cash,Credit transfer,Payment to a bank account,Bank card, The Instrument is not defined", "")

                objMain.objUtilities.AddAlphaField("CRD1", "StreetAr", "Street Arabic", 254)
                objMain.objUtilities.AddAlphaField("CRD1", "SubDivAr", "Sub Division Arabic", 254)
                objMain.objUtilities.AddAlphaField("CRD1", "CityAr", "City Arabic", 254)

                ' objMain.objUtilities.addField("OALT", "WAAlert", "WhatsApp Alert", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "NO")

                ' objMain.objUtilities.addField("AOB1", "WASent", "Whatsapp Sent", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "NO")
                objMain.objUtilities.addField("OCRD", "InvType", "Invoice Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "B2C,B2B", "B2C,B2B", "B2C")

                RunSQLScripts()
                UpdateTransactionNotification()
                'UpdateTransactionNotification()
                Me.AddUserFormAuthorizations()

                Me.AddAlerts()
                Me.QueryManager()
                Dim err As String

                err = objUtilities.AddDataToNoObjectTable("TNX_DB", Version, Version, "U_VERSION", Version)

                If err <> "" Then
                    MsgBox(err)
                End If
                objMain.objApplication.StatusBar.SetText("Your Database has now been upgraded to Version " + Version + ".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        End Try
        Return True
    End Function

#End Region

#Region "Create Tables"

    Sub LicenceNew()
        objMain.objUtilities.CreateTable("TNX_LICENCE", "Licence Administration New one", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "Company", "Company Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "DocNum", "Document Number", 100)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "License", "License", 254)

        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "EIVC", "EInvoice", 1)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "CTTX", "Corporate Tax", 1)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "VRPT", "Vat Report", 1)


        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "DB", "Database Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "HKey", "Hardware Key", 100)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "Total", "Total Licenses", 10)
        objMain.objUtilities.AddDateField("@TNX_LICENCE", "SDate", "Start Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_LICENCE", "EDate", "Expiry Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "Addon", "Add-on Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "NOUSR", "No of Users", 254)

        objMain.objUtilities.CreateTable("TNX_LICENCE_C0", "License Administration Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE_C0", "Code", "Device Code", 254)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE_C0", "USERC", "User Code", 254)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE_C0", "Name", "User Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE_C0", "Sts", "Status", 1)

        'objMain.objUtilities.CreateTable("SBO_ADDON", "License Addon", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        'objMain.objUtilities.AddAlphaField("@SBO_ADDON", "ADCODE", "Addon Code", 254)
        'objMain.objUtilities.AddAlphaField("@SBO_ADDON", "ADNAM", "Addon Name", 254)
        'objMain.objUtilities.AddAlphaField("@SBO_ADDON", "USER", "Used", 1)
        'objMain.objUtilities.AddAlphaField("@SBO_ADDON", "AVBLE", "Available", 254)

    End Sub
    'Public Sub CreateMenuTable()

    '    objMain.objUtilities.CreateTable("TENXMENU", "Menu Settings", SAPbobsCOM.BoUTBTableType.bott_Document)


    '    objMain.objUtilities.CreateTable("TENXMENUC0", "Menu Settings Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
    '    objMain.objUtilities.AddAlphaField("@TENXMENUC0", "PID", "Parent Id", 80)
    '    objMain.objUtilities.AddAlphaField("@TENXMENUC0", "PMN", "Parent Menu Name", 100)
    '    objMain.objUtilities.AddAlphaField("@TENXMENUC0", "MID", "Menu Id", 80)
    '    objMain.objUtilities.AddAlphaField("@TENXMENUC0", "MNM", "Menu Name", 100)
    '    objMain.objUtilities.AddAlphaField("@TENXMENUC0", "VSBL", "Visible", 4)


    '    objMain.objUtilities.CreateTable("TENXMENUC1", "Menu Settings Child 1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
    '    objMain.objUtilities.AddAlphaField("@TENXMENUC1", "USRID", "User Id", 80)
    '    objMain.objUtilities.AddAlphaField("@TENXMENUC1", "USRNM", "User Name", 150)
    '    objMain.objUtilities.AddAlphaField("@TENXMENUC1", "MENUID", "Menu Id", 80)
    '    objMain.objUtilities.AddAlphaField("@TENXMENUC1", "MENUNM", "Menu Name", 150)
    '    objMain.objUtilities.AddAlphaField("@TENXMENUC1", "DSBL", "Disable", 4)


    'End Sub

    Sub CreateDDTable()
        objMain.objUtilities.CreateTable("SBO_DD", "DDTable", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.CreateTable("SBO_DD1", "DD1Table", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@SBO_DD", "AppID", "Appln ID Field", 20)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "AppStat", "Appln Status Field", 20)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "Table", "Table", 40)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "Doc", "Document", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "statusN", "Status New", 20)

        objMain.objUtilities.AddAlphaField("@SBO_DD", "docnum", "Document Number", 20)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "tempid", "Template ID", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "status", "Status", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "userid", "UserID", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "objtype", "objtype", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "Isdraft", "Isdraft", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "Remarks", "Remarks", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "Appcnt", "Application", 100)
        objMain.objUtilities.AddDateField("@SBO_DD", "docdate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "Stage", "Stage", 100)

        objMain.objUtilities.AddAlphaField("@SBO_DD1", "EmpId", "Employee Id", 50)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "Stage", "Stage", 50)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "statusN", "Status New", 20)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "TempId", "Template ID", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "Userid", "UserID", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "status", "Status", 100)
        objMain.objUtilities.AddDateField("@SBO_DD1", "cdate", "Create Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "ctime", "Create Time", 20)
        objMain.objUtilities.AddDateField("@SBO_DD1", "Udate", "Update Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "Utime", "Update Time", 20)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "Remarks", "Remarks", 200)

    End Sub
    Sub LkMster()
        objMain.objUtilities.CreateTable("TNX_LKMTR", "Link Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNX_LKMTR", "CTAX", "Corpurate Tax", 150)
        objMain.objUtilities.AddAlphaField("@TNX_LKMTR", "FVRT", "FTA VAT Report Tax", 150)
    End Sub
    Sub FtaVatcal1()
        objMain.objUtilities.CreateTable("TNX_FTAVATcal", " FTA VAT Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNX_FTAVATcal", "CMR", "From Period", 2)
        objMain.objUtilities.AddAlphaField("@TNX_FTAVATcal", "SO", "To Period", 2)
        objMain.objUtilities.AddInteger("@TNX_FTAVATcal", "GRPO", "Alert Before", SAPbobsCOM.BoFldSubTypes.st_None, 5)
        objMain.objUtilities.AddDateField("@TNX_FTAVATcal", "DCNF", "Submission Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_FTAVATcal", "CMN", "Status", 1)


    End Sub
    Sub FtaVat1()
        objMain.objUtilities.CreateTable("TNX_FTAVAT", " FTA VAT Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNX_FTAVAT", "CMR", "From Period", 2)
        objMain.objUtilities.AddAlphaField("@TNX_FTAVAT", "SO", "To Period", 2)
        objMain.objUtilities.AddInteger("@TNX_FTAVAT", "GRPO", "Alert Before", SAPbobsCOM.BoFldSubTypes.st_None, 5)
        objMain.objUtilities.AddDateField("@TNX_FTAVAT", "DCNF", "Submission Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_FTAVAT", "CMN", "Status", 1)


    End Sub
    Sub CorpTax1()
        objMain.objUtilities.CreateTable("TNX_CORPTAX", "Corporate Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNX_CORPTAX", "CMR", "From Period", 2)
        objMain.objUtilities.AddAlphaField("@TNX_CORPTAX", "SO", "To Period", 2)
        objMain.objUtilities.AddInteger("@TNX_CORPTAX", "GRPO", "Alert Before", SAPbobsCOM.BoFldSubTypes.st_None, 5)
        objMain.objUtilities.AddDateField("@TNX_CORPTAX", "DCNF", "Submission Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_CORPTAX", "CMN", "Status", 1)

    End Sub
    Sub Corp()
        objMain.objUtilities.CreateTable("TNX_CORP", "Corporate Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNX_CORP", "CMR", "From Period", 2)
        objMain.objUtilities.AddAlphaField("@TNX_CORP", "SO", "To Period", 2)
        objMain.objUtilities.AddInteger("@TNX_CORP", "GRPO", "Alert Before", SAPbobsCOM.BoFldSubTypes.st_None, 5)
        objMain.objUtilities.AddDateField("@TNX_CORP", "DCNF", "Submission Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_CORP", "CMN", "Status", 1)

    End Sub
    Sub FtaVat()
        objMain.objUtilities.CreateTable("FTAVAT", "User Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

        objMain.objUtilities.AddDateField("@FTAVAT", "CMR", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@FTAVAT", "SO", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddInteger("@FTAVAT", "GRPO", "Alert Before", SAPbobsCOM.BoFldSubTypes.st_None, 5)
        objMain.objUtilities.AddDateField("@FTAVAT", "DCNF", "Submission Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@FTAVAT", "CMN", "Status", 1)
        objMain.objUtilities.AddAlphaField("@FTAVAT", "CODE", "Code", 20)
        objMain.objUtilities.AddAlphaField("@FTAVAT", "NAME", "Name", 100)

    End Sub
    Sub CorpTax()
        objMain.objUtilities.CreateTable("CORPTAX", "User Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

        objMain.objUtilities.AddDateField("@CORPTAX", "CMR", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@CORPTAX", "SO", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddInteger("@CORPTAX", "GRPO", "Alert Before", SAPbobsCOM.BoFldSubTypes.st_None, 5)
        objMain.objUtilities.AddDateField("@CORPTAX", "DCNF", "Submission Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@CORPTAX", "CMN", "Status", 1)
        objMain.objUtilities.AddAlphaField("@CORPTAX", "CODE", "Code", 20)
        objMain.objUtilities.AddAlphaField("@CORPTAX", "NAME", "Name", 100)

    End Sub

    Sub InvoicePosting()
        objMain.objUtilities.CreateTable("TNX_IP", "Invoice Posting Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddDateField("@TNX_IP", "Fdate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_IP", "Tdate", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_IP", "Status", "Status", 50)
        objMain.objUtilities.AddAlphaField("@TNX_IP", "DocType", "DocumentType", 50)


        objMain.objUtilities.CreateTable("TNX_IP_C0", "Invoice Posting Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_IP_C0", "Select", "Select", 5, "Y")
        objMain.objUtilities.AddAlphaField("@TNX_IP_C0", "DNum", "Document Number", 100)
        objMain.objUtilities.AddAlphaField("@TNX_IP_C0", "DocEntry", "Document Entry", 100)
        objMain.objUtilities.AddAlphaField("@TNX_IP_C0", "CCode", "Customer Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_IP_C0", "CName", "Customer Name", 100)
        objMain.objUtilities.AddFloatField("@TNX_IP_C0", "Tbdisc", "Total Before Discount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_IP_C0", "Disc", "Discount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_IP_C0", "Tax", "Tax", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_IP_C0", "Total", "Total", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@TNX_IP_C0", "MSG", "ZATCA Message", 254)
        objMain.objUtilities.AddAlphaField("@TNX_IP_C0", "STATUS", "ZATCA Status", 254)


    End Sub
    Sub CreateAPPROVALSTAGES()
        objMain.objUtilities.CreateTable("SBO_AST", "Approval Stages", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@SBO_AST", "SName", "Stage Name", 40)
        objMain.objUtilities.AddAlphaField("@SBO_AST", "SDesc", "Stage Description", 40)
        objMain.objUtilities.AddAlphaField("@SBO_AST", "NAP", "No. of Approvals Required", 40)
        objMain.objUtilities.AddAlphaField("@SBO_AST", "NRJ", "No. of Rejections Required", 40)
        objMain.objUtilities.AddAlphaField("@SBO_AST", "NAME", "NAME", 40)


        objMain.objUtilities.CreateTable("SBO_AST_C0", "Approval Stages", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@SBO_AST_C0", "AUTH", "AUTH", 40)
        objMain.objUtilities.AddAlphaField("@SBO_AST_C0", "EmpId", "EmpId", 50)
        objMain.objUtilities.AddAlphaField("@SBO_AST_C0", "Dept", "Dept", 30)
        objMain.objUtilities.AddAlphaField("@SBO_AST_C0", "UKey", "UKey", 30)
    End Sub
    Sub CreateAPPROVALTemplates()
        objMain.objUtilities.CreateTable("SBO_APPHDR", "Temp", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "T1", "Stage Name", 40)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "T2", "Stage Description", 40)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "NAP", "No. of Approvals Required", 40)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "C1", "Active", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "PCR", "Price Change Request", 254)



        objMain.objUtilities.CreateTable("SBO_APPREQ", "Originator", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@SBO_APPREQ", "M1_1", "Name", 40)
        objMain.objUtilities.AddAlphaField("@SBO_APPREQ", "EmpId", "EmpId", 50)
        objMain.objUtilities.AddAlphaField("@SBO_APPREQ", "M1_2", "Department", 30)

        objMain.objUtilities.CreateTable("SBO_APPAUT", "Stages", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@SBO_APPAUT", "M3_1", "M3_1", 40)
        objMain.objUtilities.AddAlphaField("@SBO_APPAUT", "M3_2", "M3_2", 50)


        objMain.objUtilities.AddAlphaField("@SBO_APPAUT", "NAMES", "Names", 30)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "Name", "Name", 50)

        objMain.objUtilities.CreateTable("SBO_APPHDR", "Conditions", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "Conds", "Always", 50)


        objMain.objUtilities.AddAlphaField("@SBO_APPREQ", "Name", "EmpId", 50)

        objMain.objUtilities.AddAlphaField("@SBO_APPREQ", "Dept", "Department", 50)

        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "Descrip", "Desciption", 100)

        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "Acive", "Active", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "Active", "Active", 1)

        'objMain.objUtilities.CreateTable("SBO_APPHDR", "Conditions", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        'objMain.objUtilities.AddButton("@SBO_APPHDR", "M3_1", "M3_1", 40)
        'objMain.objUtilities.AddButton(objform.UniqueID, "Btn_Reset", objform.Items.Item("71").Top + 230,_  objform.Items.Item("73").Left, 80, objform.Items.Item("71").Height + 10, "71", "&Reset", objform.Items.Item("71").FromPane, objform.Items.Item("71").ToPane)

        objMain.objUtilities.CreateTable("SBO_APPDOC", "Documents", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "RcReq", "Recruitment Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "RmSm", "Resume Submission", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "OL", "Offer Letter", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "VRQ", "Advertisement Request", 1) 'C'
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Price", "Price Change Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "SuppReb", "SupplierRebate", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "IMP", "ItemMasterUpload", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "ITEML", "ItemListing", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "ITEML", "ItemListing", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "BP", "BusinessPartner", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "CPA", "Corporate Tax", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "VPA", "VAT Report", 1)


        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "EmpMstr", "Employee Master Data", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "EmpSal", "Employee Salary Setup", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "ConNew", "Employee Contract Renewal", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "TS", "TimeSheet", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "LvRj", "Leave Rejoin Application", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "LvApp", "Leave Application", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "CompOff", "Compensatory off", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "AttSum", "Attendance Summary", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Shift", "Shift Schedule", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "AReg", "Attendance Regularisation", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "LER", "Leave Encashment Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "LPC", "Leave Period End Closing", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "LnApp", "Loans and Advances Application", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "LnPr", "Loan Pre Closure Application", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PD", "Payments and Deductions", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PrPre", "Payroll Pre-Process", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Claim", "Claim Application", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Review", "Employee Review / Appraisal", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PerEval", "Performance Evaluation", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Encash", "Leave Encashment", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "EOS", "End of Service Process", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Resign", "Employee Resignation Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Clear", "Exit Clearance", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Air", "Airticket Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Doc", "Document Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PTC", "Petty Cash Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PTV", "Petty Cash Voucher", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "MTR", "Material Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Trans", "Branch Transfer", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "GRPMTNG", "Group Meeting", 1)


    End Sub
    Sub CorporateTaxConfiguration()
        objMain.objUtilities.CreateTable("TNX_CTAXCNF", "Corporate Tax Configuration", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddFloatField("@TNX_CTAXCNF", "MnProfit", "Minimum Profit", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_CTAXCNF", "MxProfit", "Maximum Profit", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_CTAXCNF", "TaxPrc", "Tax Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddAlphaField("@TNX_CTAXCNF", "LAccount", "Liability Account", 100)
        objMain.objUtilities.AddAlphaField("@TNX_CTAXCNF", "EAccount", "Expenditure Account", 100)
        objMain.objUtilities.AddAlphaField("@TNX_CTAXCNF", "FINA", "Finanical Year", 10)
    End Sub

    Sub CorporateTaxCalculation()
        objMain.objUtilities.CreateTable("TNX_CTAXCAL", "Corporate Tax Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddDateField("@TNX_CTAXCAL", "FDate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_CTAXCAL", "TDate", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_CTAXCAL", "Branch", "Branch", 80)
        objMain.objUtilities.AddFloatField("@TNX_CTAXCAL", "PPeriod", "Profit Period", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_CTAXCAL", "CTax", "Corporate Tax %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_CTAXCAL", "CTaxVal", "Corporate Tax Value", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddDateField("@TNX_CTAXCAL", "JEPOSTD", "JE Posting Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_CTAXCAL", "JENo", "JE No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_CTAXCAL", "Status", "Status", 15)
        objMain.objUtilities.AddAlphaField("@TNX_CTAXCAL", "RMC", "Remarks", 254)
        objMain.objUtilities.AddAlphaField("@TNX_CTAXCAL", "DST", "Document Status", 1)
        objMain.objUtilities.AddAlphaField("@TNX_CTAXCAL", "CURR", "currency", 30)
        objMain.objUtilities.AddAlphaField("@TNX_CTAXCAL", "YEAR", "YEAR", 30)
        objMain.objUtilities.AddAlphaField("@TNX_CTAXCAL", "AIPD", "Application ID", 50)
        objMain.objUtilities.AddDateField("@TNX_CTAXCAL", "DAT", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.CreateTable("TNX_CTAXCAL_C2", "Corporate Tax Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        '.objUtilities.AddAlphaField("@TNX_ATTACH_C3", "TPA", "Target Path", 254)
        objMain.objUtilities.AddLinkField("@TNX_CTAXCAL_C2", "TPH", "Target Path", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_CTAXCAL_C2", "FNM", "File Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_CTAXCAL_C2", "FTR", "Free Text", 254)
        objMain.objUtilities.AddDateField("@TNX_CTAXCAL_C2 ", "ATCD", "Attachment Date", SAPbobsCOM.BoFldSubTypes.st_None)


    End Sub
    Sub VatReport()
        objMain.objUtilities.CreateTable("TNX_VATRP", "VAT Report Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddDateField("@TNX_VATRP", "SD", "Submission Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_VATRP", "TRNM", "TRNL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_VATRP", "APPI", "Application Id", 254)
        objMain.objUtilities.AddAlphaField("@TNX_VATRP", "TPNE", "Taxable Person Name(English)", 254)
        objMain.objUtilities.AddAlphaField("@TNX_VATRP", "TPNA", "Taxable Person Name(Arabic)", 254)
        objMain.objUtilities.AddAlphaField("@TNX_VATRP", "TPA", "Taxable Person Address", 254)
        objMain.objUtilities.AddAlphaField("@TNX_VATRP", "FTY", "Document Status", 1)
        objMain.objUtilities.AddDateField("@TNX_VATRP", "VRPY", "VAT Return Period", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_VATRP", "VATTO", "VAT Return Period TO", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.CreateTable("TNX_VATCTM_C1", "VAT Report Child1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_VATCTM_C1", "TRN", "Description", 254)
        objMain.objUtilities.AddAlphaField("@TNX_VATCTM_C1", "ACN", "Account Number", 254)
        ' objMain.objUtilities.AddAlphaField("@TNX_VATCTM_C1", "TLP", "Tool Tip", 254)
        objMain.objUtilities.AddFloatField("@TNX_VATCTM_C1", "AMT", "Amount(AED)", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_VATCTM_C1", "VATA", "VAT Amount(AED)", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_VATCTM_C1", "AST", "Adjustment(AED)", SAPbobsCOM.BoFldSubTypes.st_Sum)

        objMain.objUtilities.CreateTable("TNX_ATTACH_C3", "VAT Report Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        '.objUtilities.AddAlphaField("@TNX_ATTACH_C3", "TPA", "Target Path", 254)
        objMain.objUtilities.AddLinkField("@TNX_ATTACH_C3", "TPA", "Target Path", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_ATTACH_C3", "FN", "File Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_ATTACH_C3", "FTT", "Free Text", 254)
        objMain.objUtilities.AddDateField("@TNX_ATTACH_C3", "ATD", "Attachment Date", SAPbobsCOM.BoFldSubTypes.st_None)
        'objMain.objUtilities.AddFloatField("@TNX_ATTACH_C3", "VATA", "VAT Amount(AED)", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.CreateTable("TNX_VATRP_C0", "VAT Report Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_VATRP_C0", "VATE", "Description", 254)
        objMain.objUtilities.AddAlphaField("@TNX_VATRP_C0", "ARN", "Account Number", 254)
        ' objMain.objUtilities.AddAlphaField("@TNX_VATRP_C0", "TVAT", "Tool Tip", 254)
        objMain.objUtilities.AddFloatField("@TNX_VATRP_C0", "AUT", "Amount(AED)", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_VATRP_C0", "RVAT", "Recoverable VAT Amount(AED)", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_VATRP_C0", "AVAT", "Adjustment(AED)", SAPbobsCOM.BoFldSubTypes.st_Sum)
        'objMain.objUtilities.AddAlphaField("@TNX_VATRP_C0", "ASP", "Are you using the profit margin scheme ", 1)
        'objMain.objUtilities.AddAlphaField("@TNX_VATRP_C0", "VATGCC", "GCC VAT", 254)
        'objMain.objUtilities.AddAlphaField("@TNX_VATRP_C0", "TTVAT", "Tool Tip", 254)
        'objMain.objUtilities.AddFloatField("@TNX_VATRP_C0", "DVAT", "Amount(AED)", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        'objMain.objUtilities.AddFloatField("@TNX_VATRP_C0", "EVATA", "VAT Amount(AED)", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        'objMain.objUtilities.AddFloatField("@TNX_VATRP_C0", "AAVAT", "Adjustment(AED)", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNX_VATRP", "TVD", "Total Vat", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_VATRP", "TVRP", "Total Value of recoverable tax for the period ", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_VATRP", "NVAT", "Total Value of recoverable", SAPbobsCOM.BoFldSubTypes.st_Sum)
        'objMain.objUtilities.AddAlphaField("@TNX_VATRP", "TAVMN", "Total vat refund", 1)




    End Sub

    Sub OnboardingProcessTable()
        objMain.objUtilities.CreateTable("TNX_ONBP", "Onboarding Process Table", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "CName", "Company Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "Email", "Email", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "City", "City", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "PCode", "Postal Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "Sheet", "Additional Sheet", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "Address", "Short Address", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "ANum", "Additional Number", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "OTP", "One Time Password", 30)
        objMain.objUtilities.AddDateField("@TNX_ONBP", "DocDate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "Country", "Country", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "Type", "Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "IdNum", "ID Number", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "District", "District", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "SName", "Street Name", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "BNum", "Building Number", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "BCat", "Business Category", 50)
        objMain.objUtilities.AddAlphaMemoField("@TNX_ONBP", "CSR", "CSR", 5000)
        objMain.objUtilities.AddAlphaMemoField("@TNX_ONBP", "PCSID", "PCSID", 5000)
        objMain.objUtilities.AddAlphaMemoField("@TNX_ONBP", "CSID", "CSID", 5000)
        objMain.objUtilities.AddAlphaMemoField("@TNX_ONBP", "PBKey", "Public Key", 5000)
        objMain.objUtilities.AddAlphaMemoField("@TNX_ONBP", "PRKey", "Private Key", 5000)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "SName", "Street Name", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "PType", "Posting Type", 50)


        objMain.objUtilities.CreateTable("TNX_ONBP_C0", "Onboarding Process Child Table", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP_C0", "CCode", "Customer Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP_C0", "PMode", "Posting Mode", 30)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP_C0", "IActive", "In Active", 5)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP_C0", "ILog", "Is Logged in", 10)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP_C0", "Options", "Options", 10)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP_C0", "Remarks", "Remarks", 254)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP_C0", "URL", "URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP_C0", "UUID", "UUID", 254)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP_C0", "Key", "Key", 254)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP_C0", "PType", "PType", 254)

    End Sub
    Sub PayPosting()
        objMain.objUtilities.CreateTable("TNX_PAY", "Invoice Posting Table", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        '  objMain.objUtilities.AddAlphaField("@TNX_INVF", "IPURL", "Invoice Posting URL", 254)
        'objMain.objUtilities.AddAlphaField("@TNX_INVF", "DBNE", "DB Name", 254)
        'objMain.objUtilities.AddAlphaField("@TNX_INVF", "CTY", "Country", 60)
        'objMain.objUtilities.AddAlphaField("@TNX_INVF", "IPPRL", "Invoice Possting URL", 254)
    End Sub
    Sub PayLoadData()
        objMain.objUtilities.CreateTable("TNX_PAYLD", "Invoice Posting Table", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        '  objMain.objUtilities.AddAlphaField("@TNX_INVF", "IPURL", "Invoice Posting URL", 254)
        'objMain.objUtilities.AddAlphaField("@TNX_INVF", "DBNE", "DB Name", 254)
        'objMain.objUtilities.AddAlphaField("@TNX_INVF", "CTY", "Country", 60)
        'objMain.objUtilities.AddAlphaField("@TNX_INVF", "IPPRL", "Invoice Possting URL", 254)
    End Sub
    Sub InvoicePostingTable()
        objMain.objUtilities.CreateTable("TNX_INVF", "Invoice Posting Table", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        '  objMain.objUtilities.AddAlphaField("@TNX_INVF", "IPURL", "Invoice Posting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "DBNE", "DB Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "CTY", "Country", 60)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "IPPRL", "Invoice Possting URL", 254)


        objMain.objUtilities.AddAlphaField("@TNX_INVF", "IQRCG", "Invoice QR Code Generation URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "IUUID", "Invoice UUIDs URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "CMPRL", "Credit Memo Posting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "CMQRC", "Credit Memo QR Code Generation URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "CMURL", "Credit Memo UUID URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "BPID", "Business Profile ID", 100)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "APIK", "API Key", 60)
        'objMain.objUtilities.AddDateField("@TNX_ONBP", "DocDate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        'objMain.objUtilities.AddAlphaField("@TNX_ONBP", "Country", "Country", 50)

        'B2B
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "IPURLB", "Invoice Possting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "IQRCGB", "Invoice QR Code Generation URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "UUIDB", "Invoice UUIDs URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "CMPRLB", "Credit Memo Posting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "CMQRB", "Credit Memo QR Code Generation URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "CMURLB", "Credit Memo UUID URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "BPIDB", "Business Profile ID", 100)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "APIKB", "API Key", 60)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TYPE", "Posting Type", 60)

        'LIVE

        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TIPPRL", "Invoice Possting URL", 254)


        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TIQRCG", "Invoice QR Code Generation URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TIUUID", "Invoice UUIDs URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TCMPRL", "Credit Memo Posting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TCMQRC", "Credit Memo QR Code Generation URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TCMURL", "Credit Memo UUID URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TBPID", "Business Profile ID", 100)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TAPIK", "API Key", 60)
        'objMain.objUtilities.AddDateField("@TNX_ONBP", "DocDate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        'objMain.objUtilities.AddAlphaField("@TNX_ONBP", "Country", "Country", 50)

        'B2B
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TIPURLB", "Invoice Possting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TIQRCGB", "Invoice QR Code Generation URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TUUIDB", "Invoice UUIDs URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TCMPRLB", "Credit Memo Posting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TCMQRB", "Credit Memo QR Code Generation URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TCMURLB", "Credit Memo UUID URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TBPIDB", "Business Profile ID", 100)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TAPIKB", "API Key", 60)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LTYPE", "Posting Type", 60)

        'UAE test
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "AIPPRL", "Access Token URL", 254)


        objMain.objUtilities.AddAlphaField("@TNX_INVF", "AIUUID", "Get UUIDs URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "ACMQRC", "Password", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "ABPID", "Business Profile ID", 100)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "ALIQRCG", "Invoice Posting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "ALCMPRL", "User Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "ALCMURL", "Credit Memo UUID URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "ALAPIK", "API Key", 60)

        'live
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LIPPRL", "Access Token URL", 254)


        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LIUUID", "Get UUIDs URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LCMQRC", "Password", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LBPID", "Business Profile ID", 100)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LIQRCG", "Invoice Posting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LCMPRL", "User Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LCMURL", "Credit Memo UUID URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LAPIK", "API Key", 60)

        objMain.objUtilities.AddAlphaField("@TNX_INVF", "UTA", "Access Token URL", 254)


        objMain.objUtilities.AddAlphaField("@TNX_INVF", "GURL", "Get UUIDs URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "BPI", "Password", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "IPU", "Business Profile ID", 100)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "CPL", "Invoice Posting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "USM", "User Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "CMRL", "Credit Memo UUID URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "APID", "API Key", 60)

        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LUTA", "Access Token URL", 254)


        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LGURL", "Get UUIDs URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LBPI", "Password", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LIPU", "Business Profile ID", 100)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LCPL", "Invoice Posting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LUSM", "User Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LCMRL", "Credit Memo UUID URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LAPID", "API Key", 60)
    End Sub

    Sub Errorlogs()

        objMain.objUtilities.CreateTable("TNX_ERRORLOGS", "Error Logs", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        objMain.objUtilities.AddDateField("@TNX_ERRORLOGS", "DATE", "Error Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_ERRORLOGS", "TIME", "Error Time", 10)
        objMain.objUtilities.AddAlphaField("@TNX_ERRORLOGS", "ERROR_MSG", "Error Message", 254)

    End Sub





#End Region

#Region "SQL Scripts, Queries, Authorizations"
    Sub RunSQLScripts()
        Try

            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_ZATCA_OINV.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_ZATCA_ORIN.sql"))

            'objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("SQLQuery.PayLoad.txt"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_EINVOICE_TN.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("Corporate.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("VatReport.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("Tnx_Branch.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("EInvoice.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNXPLByBranch.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("C_CORP_EXCEPTIONS.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("C_CORP_COMPUTATION.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("C_CORP_ADJUSTMENTBREAKDOWN.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("C_CORP_LIABILITYTREND.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("C_CORP_MONTHLYTAXABLEINCOME.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("C_CORP_LOADKPIS.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("C_VAT_LOADKPI.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("C_VAT_MONTHLYVATCHART.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("C_VAT_BUSINESSUNIT.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("C_VAT_EXCEPTIONSUMMARY.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("C_VAT_FILLINGREADINESS.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("C_VAT_EXCEPTIONDETAILS.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("C_DASHBOARD_EINVOICEKPI.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("C_DASHBOARD_TREND_CHART.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("C_DASHBOARD_STATUSMIX.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("C_DASHBOARD_ERROR_CHART.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("C_DASHBOARD_RECENTFAILURES.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("C_DASHBOARD_ERRORCHART.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("C_DASHBOARD_ERRORCHART.sql"))
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub QueryManager()
        Try

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Sales Orders Posted", 10, objUtilities.LoadEmbeddedSQL("SalesOrdersPosted.sql"), "Alert")

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Corporate Tax - Open Status", 12, objUtilities.LoadEmbeddedSQL("CorporateTaxOpenStatus.sql"), "Alert")

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Corporate Tax - Approved", 14, objUtilities.LoadEmbeddedSQL("CorporateTaxApproved.sql"), "Alert")

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Corporate Tax JE post", 16, objUtilities.LoadEmbeddedSQL("CorporateTaxJEPost.sql"), "Alert")

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "FTA Vat Report", 18, objUtilities.LoadEmbeddedSQL("FTAVatReport.sql"), "Alert")

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "FTA Vat Report- Approved", 20, objUtilities.LoadEmbeddedSQL("FTAVatReportApproved.sql"), "Alert")

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "E Invoice - status", 22, objUtilities.LoadEmbeddedSQL("EInvoiceStatus.sql"), "Alert")

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "E Invoice AR credit", 24, objUtilities.LoadEmbeddedSQL("EInvoiceARCredit.sql"), "Alert")

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "E Invoice - pending", 26, objUtilities.LoadEmbeddedSQL("EInvoicePending.sql"), "Alert")

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "AP Credit Memo Alert", 28, objUtilities.LoadEmbeddedSQL("APCreditMemoAlert.sql"), "Alert")

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "CorporatetaxFiling", 30, objUtilities.LoadEmbeddedSQL("CorporatetaxFiling.sql"), "Alert")

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "VatReportFiling", 32, objUtilities.LoadEmbeddedSQL("VatReportFiling.sql"), "Alert")
        Catch ex As Exception

        End Try
    End Sub
    Sub AddAlerts()
        Try


            objUtilities.CreateCustomAlert("AP Credit Memo Alert", 8, 1, 1, AlertManagementFrequencyType.atfi_Days)


            objUtilities.CreateCustomAlert("Sales Orders Posted", 10, 1, 1, AlertManagementFrequencyType.atfi_Days)


            objUtilities.CreateCustomAlert("Corporate Tax - Open Status", 12, 1, 1, AlertManagementFrequencyType.atfi_Days)


            objUtilities.CreateCustomAlert("Corporate Tax - Approved", 14, 1, 1, AlertManagementFrequencyType.atfi_Days)


            objUtilities.CreateCustomAlert("Corporate Tax JE post", 16, 1, 1, AlertManagementFrequencyType.atfi_Days)

            objUtilities.CreateCustomAlert("FTA Vat Report", 18, 1, 1, AlertManagementFrequencyType.atfi_Days)


            objUtilities.CreateCustomAlert("FTA Vat Report- Approved", 20, 1, 1, AlertManagementFrequencyType.atfi_Days)


            objUtilities.CreateCustomAlert("E Invoice - status", 22, 1, 1, AlertManagementFrequencyType.atfi_Days)


            objUtilities.CreateCustomAlert("E Invoice AR credit", 24, 1, 1, AlertManagementFrequencyType.atfi_Days)


            objUtilities.CreateCustomAlert("E Invoice - pending", 26, 1, 1, AlertManagementFrequencyType.atfi_Days)


            objUtilities.CreateCustomAlert("AP Credit Memo Alert", 28, 1, 1, AlertManagementFrequencyType.atfi_Days)
            objUtilities.CreateCustomAlert("CorporatetaxFiling", 30, 1, 1, AlertManagementFrequencyType.atfi_Days)
            objUtilities.CreateCustomAlert("VatReportFiling", 32, 1, 1, AlertManagementFrequencyType.atfi_Days)

            objMain.objApplication.StatusBar.SetText("All Alerts created successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Sub UpdateTransactionNotification()
        Try

            Dim oRecordset As Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim procedureName As String = "SBO_SP_TRANSACTIONNOTIFICATION"

            ' 1. Get existing procedure source
            Dim DBName As String = objMain.objCompany.CompanyDB
            Dim Qry As String = "SELECT ""DEFINITION"" FROM ""SYS"".""PROCEDURES"" WHERE ""SCHEMA_NAME"" = '" & DBName & "' AND ""PROCEDURE_NAME"" = 'SBO_SP_TRANSACTIONNOTIFICATION'"
            oRecordset.DoQuery(Qry)

            If oRecordset.EoF Then
                Throw New Exception("Procedure not found.")
            End If

            Dim originalCode As String = oRecordset.Fields.Item(0).Value.ToString()
            If originalCode.Contains("TNX_EINVOICE_TN") Then
                Exit Sub
            End If
            ' 2. Your code to append
            Dim codeToInsert As String = "
    
    if :error = 0 then
        CALL TNX_EINVOICE_TN(:object_type, :transaction_type, :num_of_cols_in_key, :list_of_key_cols_tab_del, :list_of_cols_val_tab_del, :error, :error_message);
    end if;
select :error, :error_message FROM dummy;
    "

            ' 3. Insert before the final "select :error, :error_message FROM dummy;"
            Dim finalCode As String = ""
            If originalCode.Contains("select :error, :error_message FROM dummy;") Then
                finalCode = originalCode.Replace("select :error, :error_message FROM dummy;", codeToInsert)
            Else
                Throw New Exception("""select :error, :error_message FROM dummy;"" not found in procedure definition.")
            End If
            finalCode = finalCode.Replace("CREATE", "ALTER")

            oRecordset.DoQuery(finalCode)
            objMain.objApplication.StatusBar.SetText("Transaction Notification Procedure updated successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub AddUserFormAuthorizations()
        Try

            objUtilities.AddAuthorization("VATR", "FTA VAT REPORT", "VATR", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("CTAXCAL", "Corporate Tax Calculation", "CTAXCAL", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("CTAXC", "Corporate Tax Configuration", "CTAXC", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("CONFT", "E Invoicing Configuration", "CONFT", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("COTAX", "Corporate Tax", "COTAX", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("frm_FTAVM", "FTA VAT", "frm_FTAVM", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("frm_LKMTR", "LINK MASTER", "frm_LKMTR", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region


End Class







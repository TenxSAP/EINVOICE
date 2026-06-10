Option Strict Off
Option Explicit On
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports System.Xml

Public Class MainCls

#Region "Declaration"
    Public WithEvents objApplication As SAPbouiCOM.Application
    Public objCompany As SAPbobsCOM.Company
    Public ExePath As String
    Public objUtilities As Utilities
    Public objDatabaseCreation As DatabaseCreation
    Public Shared ohtLookUpForm As Hashtable = New Hashtable
    'Addon Files
    'Public ObjFORCAST As clsFORCAST
    'Public ObjSUBTYPE As clssubmenutype
    'Public oClsParameterSelection As ClsParameterSelection
    'Public oApprovedForecating As ApprovedForecating
    'Public oSubParameterSelection As SubParameterSelection
    Public ObjCorporateTaxConfiguration As clsCorporateTexConfig
    Public ObjCorporateTaxCalculation As ClsCorporateTaxCalcu

    'vsm
    Public ObjClsCorpTax As ClsCorpTax
    Public ObjclsFtaVat As clsFtaVat
    Public ObjClsLkMstr As ClsLkMstr
    Public objMenusettings As ClsMenuSettings

    Public objLicenceNew As Cfrm_LicenceAdministrationNew
    Public objDevice As DeviceMaster
    Public objLicenceAdministration As cfrm_LicenseAdministration


    'vsm



    Public ObjVatReport As ClsVatReports
    Public ObjPayloadD As ClsPayloadD
    ' Public ObjClsCorpTax As ClsCorpTax
    Public ObjClsCorpTaxMstr As ClsCropTaxMstr
    Public ObjclsFtaVatMstr As ClsFtaVatMstr


    Public objEstimation As clsEstimation
    Public objARInvoice As clsARInvoice
    Public objARCreditMemo As clsARCreditMemo
    Public objARDownPayment As clsARDownPayment
    Public objPayLoad As ClsPayLoad

    Public ObjApproval As clsApproval
    Public ObjGRID As clsGRID
    Public ObjADR As clsApprovalDecision
    Public ObjAStg As ClsAPPROVALSTAGES
    Public ObjAPTEMP As clsAPPROVALTEMP
    Public ObjGRIDES As ClsGRIDES
    Public ObjDraftProcedure As ClsDraftProcedureStages
    Public ObjVIEW As ClsVIEWFORAPP
    Public objSAPAlertWindow As clsSAPAlertWindow


    Public objAREinvoice As CLSEinvoiceButton

    Public objInvPost As ClsInvPost
    Public objOnboarding As ClsOnboarding
    'Vamshi Sai
    Public objInvoicePosting As ClsInvoicePsoting

    Public oGeneralService As SAPbobsCOM.GeneralService
    Public oGeneralData As SAPbobsCOM.GeneralData
    Public oSons As SAPbobsCOM.GeneralDataCollection
    Public oSon As SAPbobsCOM.GeneralData
    Public oChildren As SAPbobsCOM.GeneralDataCollection
    Public oChild As SAPbobsCOM.GeneralData
    Public sCmp As SAPbobsCOM.CompanyService
    Public oGeneralParams As SAPbobsCOM.GeneralDataParams
    Public oGeneralService1 As SAPbobsCOM.GeneralService
    Public oGeneralData1 As SAPbobsCOM.GeneralData
    Public oChildren1 As SAPbobsCOM.GeneralDataCollection
    Public oChild1 As SAPbobsCOM.GeneralData
    Public sCmp1 As SAPbobsCOM.CompanyService
    Public oGeneralParams1 As SAPbobsCOM.GeneralDataParams
    Public oitem As SAPbouiCOM.Item
    Dim SOSeries As String = ""
    Dim SODocNum As String = ""
    Dim PaymentType As String = ""
#End Region
    Public Sub New()
        objUtilities = New Utilities
        objDatabaseCreation = New DatabaseCreation
    End Sub

#Region "Initialilse"
    Public Function Initialise() As Boolean
        objApplication = objUtilities.GetApplication()
        If objApplication Is Nothing Then Return False
        objCompany = objUtilities.GetCompany(objApplication)
        If objCompany Is Nothing Then : Return False : Exit Function : End If
        'If Me.Addonlicence("10XComplianceOS", objCompany.CompanyDB, objCompany.UserName) = False Then

        '    MessageBox.Show("Invalid 10XComplianceOS licence, please contact 10X Team")

        '    objMain.objCompany.Disconnect()

        '    Exit Function

        'End If
        If Not objDatabaseCreation.CreateTables() Then Return False
        Me.LoadFromXML("Menu.xml")
        If objMain.objCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            IsNull = "IFNULL"
            GetDate = "NOW"
            HanaLen = "Length"
            HanaInt = "Integer"
            Hana = "Hana"
            Concate = "||"
            SelectCase = ""
        Else
            IsNull = "IsNull"
            GetDate = "GetDate"
            HanaLen = "Len"
            HanaInt = "Int"
            Sql = "Sql"
            Concate = "+"
            SelectCase = "Select"
        End If
        CreateObjects()
        CheckLicenceMenuEnableDisable()

        'Try
        '    If objMain.objApplication.Menus.Exists("ME_TenXFrameWork") = True Then
        '        objMain.objApplication.Menus.Item("ME_TenXFrameWork").Image = System.Windows.Forms.Application.StartupPath & "/TenXFrameWork.JPG"
        '    End If
        'Catch ex As Exception
        'End Try
        'objApplication.StatusBar.SetText(Me.GetUserName() + " ! :-)  TenXFrameWork  Addon has been Connected.......You can continue your work ..........", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        objApplication.StatusBar.SetText("E Invoice Addon is connected....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Return True
    End Function
#End Region

    'Public Function GetUserName1() As String
    '    Dim oRsGetUserNames As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    Try
    '        Dim mHour As Integer
    '        Dim Greeting As String
    '        mHour = Hour(Now)
    '        If mHour < 12 Then
    '            Greeting = "Good Morning"
    '        ElseIf mHour < 16 Then
    '            Greeting = "Good Afternoon"
    '        Else
    '            Greeting = "Good Evening"
    '        End If

    '        Dim GetUserNames As String = "Select ""U_NAME"" From OUSR Where ""USER_CODE"" = '" & objMain.objCompany.UserName & "'"
    '        oRsGetUserNames.DoQuery(GetUserNames)
    '        Return Greeting + "...." + oRsGetUserNames.Fields.Item(0).Value.ToString
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        oRsGetUserNames = Nothing
    '    End Try
    'End Function

#Region "Create Object"
    Private Sub CreateObjects()
        'ObjFORCAST = New clsFORCAST
        'oClsParameterSelection = New ClsParameterSelection
        'oSubParameterSelection = New SubParameterSelection
        'oApprovedForecating = New ApprovedForecating
        'ObjSUBTYPE = New clssubmenutype
        ObjClsCorpTax = New ClsCorpTax
        ObjclsFtaVat = New clsFtaVat
        objMenusettings = New ClsMenuSettings
        ObjApproval = New clsApproval
        ObjGRID = New clsGRID
        ObjADR = New clsApprovalDecision
        ObjAStg = New ClsAPPROVALSTAGES
        ObjAPTEMP = New clsAPPROVALTEMP
        ObjGRIDES = New ClsGRIDES
        ObjDraftProcedure = New ClsDraftProcedureStages
        ObjVIEW = New ClsVIEWFORAPP
        objSAPAlertWindow = New clsSAPAlertWindow
        ObjclsFtaVatMstr = New ClsFtaVatMstr
        ObjClsCorpTaxMstr = New ClsCropTaxMstr
        objEstimation = New clsEstimation
        objARInvoice = New clsARInvoice
        objARCreditMemo = New clsARCreditMemo
        objARDownPayment = New clsARDownPayment
        objAREinvoice = New CLSEinvoiceButton
        objPayLoad = New ClsPayLoad
        ObjCorporateTaxConfiguration = New clsCorporateTexConfig
        ObjCorporateTaxCalculation = New ClsCorporateTaxCalcu
        ObjVatReport = New ClsVatReports
        objInvPost = New ClsInvPost
        objOnboarding = New ClsOnboarding
        ObjPayloadD = New ClsPayloadD
        'Vamshi Sai
        objInvoicePosting = New ClsInvoicePsoting

        'vsm
        ObjClsCorpTax = New ClsCorpTax
        ObjclsFtaVat = New clsFtaVat
        ObjClsLkMstr = New ClsLkMstr

        objLicenceNew = New Cfrm_LicenceAdministrationNew
        objDevice = New DeviceMaster
        objLicenceAdministration = New cfrm_LicenseAdministration

        'vsm


    End Sub
#End Region

#Region "Create UDO"
    Public Sub CreateApprovalTemplatesUDO()
        If Not Me.UDOExists("SBO_APP") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("SBO_APP", "SBO_APP", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "SBO_APPHDR", "SBO_APPREQ", "SBO_APPDOC", "SBO_APPAUT")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub CreateMenuUDO()
        If Not Me.UDOExists("OTENXMENU") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("OTENXMENU", "Menu Settings UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TENXMENU", "TENXMENUC0", "TENXMENUC1")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub LkMsterUDO()
        If Not Me.UDOExists("TNX_LKMTR") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("TNX_LKMTR", "TNX_LKMTR", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "TNX_LKMTR")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub FtaVatcalUDO1()
        If Not Me.UDOExists("TNX_FTAVATcal") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDONoLog("TNX_FTAVATcal", "TNX_FTAVATcal", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "TNX_FTAVATcal")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub FtaVatUDO1()
        If Not Me.UDOExists("TNX_FTAVAT") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDONoLog("TNX_FTAVAT", "TNX_FTAVAT", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "TNX_FTAVAT")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub CorpTaxUDO1()
        If Not Me.UDOExists("TNX_CORPTAX") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDONoLog("TNX_CORPTAX", "TNX_CORPTAX", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "TNX_CORPTAX")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub CorpUDO()
        If Not Me.UDOExists("TNX_CORP") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDONoLog("TNX_CORP", "TNX_CORP", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "TNX_CORP")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub CreateAPPROVALSTAGESUDO()
        If Not Me.UDOExists("SBO_ASTGUDO") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            Me.registerUDONoLog("SBO_ASTGUDO", "SBO_ASTGUDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "SBO_AST", "SBO_AST_C0")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub FtaVatUDO()
        If Not Me.UDOExists("UDO_FTAVAT") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("UDO_FTAVAT", "UDO_FTAVAT", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "FTAVAT")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub CorpTaxUDO()
        If Not Me.UDOExists("UDO_CORPTAX") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("UDO_CORPTAX", "UDO_CORPTAX", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "CORPTAX")
            findaliasdescription = Nothing
        End If
    End Sub

    Public Sub DDUDO()
        If Not Me.UDOExists("SBO_DDUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("SBO_DDUDO", "SBO_DDUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "SBO_DD", "SBO_DD1")
            findAliasNDescription = Nothing
        End If

    End Sub
    Public Sub CreateONBPUDO()
        If Not Me.UDOExists("TNX_ONBP_UDO") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("TNX_ONBP_UDO", "Onboarding Process UDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "TNX_ONBP", "TNX_ONBP_C0")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub CreateIPUDO()
        If Not Me.UDOExists("TNX_IPUDO") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_IPUDO", "Invoice Posting UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_IP", "TNX_IP_C0")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub VATREPORTUDO()
        If Not Me.UDOExists("TNXVATRUDO") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNXVATRUDO", "TNXVATRUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_VATRP", "TNX_VATRP_C0", "TNX_VATCTM_C1", "TNX_ATTACH_C3")
            findaliasdescription = Nothing
        End If
    End Sub

    Public Sub CorporateTaxCalculationUDO()
        If Not Me.UDOExists("TNX_CTCAUDO") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_CTCAUDO", "TNX_CTCAUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_CTAXCAL", "TNX_CTAXCAL_C2")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub CTAXConifgUDO()
        If Not Me.UDOExists("TNX_CTXUDO") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDONoLog("TNX_CTXUDO", "TNX_CTXUDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "TNX_CTAXCNF")
            findaliasdescription = Nothing
        End If
    End Sub


    Public Sub CreateDatMTUDO()
        If Not Me.UDOExists("TNX_DBM_UDO") Then
            Dim findaliasndescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("TNX_DBM_UDO", "TNX_DBM_UDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasndescription, "TNX_DM")
            findaliasndescription = Nothing
        End If
    End Sub

    'vamshi sai
    Public Sub PAYUDO()
        If Not Me.UDOExists("TNXPAYUDO") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("TNXPAYUDO", "TNXPAYUDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "TNX_PAY")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub PAYLOADUDO()
        If Not Me.UDOExists("TNXPAYLDUDO") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("TNXPAYLDUDO", "TNXPAYLDUDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "TNX_PAYLD")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub CreateINVFUDO()
        If Not Me.UDOExists("TNX_INVF_UDO") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("TNX_INVF_UDO", "Invoicing Configuration UDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "TNX_INVF")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub CreateLicenceNewUDO()
        If Not Me.UDOExists("TNX_LICENSE_UDO") Then
            Dim findaliasndescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_LICENSE_UDO", "Licence  Administration New1 UDO ", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasndescription, "TNX_LICENCE", "TNX_LICENCE_C0")
            findaliasndescription = Nothing
        End If
    End Sub

#End Region

#Region "UDO Exists"
    Public Function UDOExists(ByVal code As String) As Boolean
        GC.Collect()
        Dim v_UDOMD As SAPbobsCOM.UserObjectsMD
        Dim v_ReturnCode As Boolean
        v_UDOMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
        v_ReturnCode = v_UDOMD.GetByKey(code)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UDOMD)
        v_UDOMD = Nothing
        Return v_ReturnCode
    End Function
#End Region

#Region "Register UDO"

    Function registerUDO(ByVal UDOCode As String, ByVal UDOName As String, ByVal UDOType As SAPbobsCOM.BoUDOObjType, ByVal findAliasNDescription As String(,), ByVal parentTableName As String, Optional ByVal childTable1 As String = "", Optional ByVal childTable2 As String = "", Optional ByVal childTable3 As String = "", Optional ByVal childTable4 As String = "", Optional ByVal childTable5 As String = "", Optional ByVal childTable6 As String = "", Optional ByVal LogOption As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO) As Boolean
        Dim actionSuccess As Boolean = False
        Try
            registerUDO = False
            Dim v_udoMD As SAPbobsCOM.UserObjectsMD
            v_udoMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            v_udoMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanLog = LogOption
            v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.Code = UDOCode
            v_udoMD.Name = UDOName
            v_udoMD.TableName = parentTableName
            If LogOption = SAPbobsCOM.BoYesNoEnum.tYES Then
                v_udoMD.LogTableName = "L" & parentTableName
            End If
            v_udoMD.ObjectType = UDOType
            For i As Int16 = 0 To findAliasNDescription.GetLength(0) - 1
                If i > 0 Then v_udoMD.FindColumns.Add()
                v_udoMD.FindColumns.ColumnAlias = findAliasNDescription(i, 0)
                v_udoMD.FindColumns.ColumnDescription = findAliasNDescription(i, 1)
            Next
            If childTable1 <> "" Then
                v_udoMD.ChildTables.TableName = childTable1
                v_udoMD.ChildTables.Add()
            End If
            If childTable2 <> "" Then
                v_udoMD.ChildTables.TableName = childTable2
                v_udoMD.ChildTables.Add()
            End If
            If childTable3 <> "" Then
                v_udoMD.ChildTables.TableName = childTable3
                v_udoMD.ChildTables.Add()
            End If
            If childTable4 <> "" Then
                v_udoMD.ChildTables.TableName = childTable4
                v_udoMD.ChildTables.Add()
            End If
            If childTable5 <> "" Then
                v_udoMD.ChildTables.TableName = childTable5
                v_udoMD.ChildTables.Add()
            End If
            If childTable6 <> "" Then
                v_udoMD.ChildTables.TableName = childTable6
                v_udoMD.ChildTables.Add()
            End If

            If v_udoMD.Add() = 0 Then
                registerUDO = True
                objMain.objApplication.StatusBar.SetText("Successfully Registered UDO >" & UDOCode & ">" & UDOName & " >" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objMain.objApplication.StatusBar.SetText("Failed to Register UDO >" & UDOCode & ">" & UDOName & " >" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                registerUDO = False
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_udoMD)
            v_udoMD = Nothing
            GC.Collect()
        Catch ex As Exception
            objMain.objApplication.SetStatusBarMessage(ex.Message)
        End Try
    End Function

    Function registerUDONoLog(ByVal UDOCode As String, ByVal UDOName As String, ByVal UDOType As SAPbobsCOM.BoUDOObjType, ByVal findAliasNDescription As String(,), ByVal parentTableName As String, Optional ByVal childTable1 As String = "", Optional ByVal childTable2 As String = "", Optional ByVal childTable3 As String = "", Optional ByVal childTable4 As String = "", Optional ByVal childTable5 As String = "", Optional ByVal LogOption As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO) As Boolean
        Dim actionSuccess As Boolean = False
        Try
            registerUDONoLog = False
            Dim v_udoMD As SAPbobsCOM.UserObjectsMD
            v_udoMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            v_udoMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanLog = LogOption
            v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.Code = UDOCode
            v_udoMD.Name = UDOName
            v_udoMD.TableName = parentTableName
            If LogOption = SAPbobsCOM.BoYesNoEnum.tYES Then
                v_udoMD.LogTableName = "A" & parentTableName
            End If
            v_udoMD.ObjectType = UDOType
            For i As Int16 = 0 To findAliasNDescription.GetLength(0) - 1
                If i > 0 Then v_udoMD.FindColumns.Add()
                v_udoMD.FindColumns.ColumnAlias = findAliasNDescription(i, 0)
                v_udoMD.FindColumns.ColumnDescription = findAliasNDescription(i, 1)
            Next
            If childTable1 <> "" Then
                v_udoMD.ChildTables.TableName = childTable1
                v_udoMD.ChildTables.Add()
            End If
            If childTable2 <> "" Then
                v_udoMD.ChildTables.TableName = childTable2
                v_udoMD.ChildTables.Add()
            End If
            If childTable3 <> "" Then
                v_udoMD.ChildTables.TableName = childTable3
                v_udoMD.ChildTables.Add()
            End If
            If childTable4 <> "" Then
                v_udoMD.ChildTables.TableName = childTable4
                v_udoMD.ChildTables.Add()
            End If
            If childTable5 <> "" Then
                v_udoMD.ChildTables.TableName = childTable5
                v_udoMD.ChildTables.Add()
            End If

            If v_udoMD.Add() = 0 Then
                registerUDONoLog = True
                objMain.objApplication.StatusBar.SetText("Successfully Registered UDO >" & UDOCode & ">" & UDOName & " >" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objMain.objApplication.StatusBar.SetText("Failed to Register UDO >" & UDOCode & ">" & UDOName & " >" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                registerUDONoLog = False
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_udoMD)
            v_udoMD = Nothing
            GC.Collect()
        Catch ex As Exception
            objMain.objApplication.SetStatusBarMessage(ex.Message)
        End Try
    End Function

#End Region

#Region "Add Menu's With XML"


    Private Sub LoadFromXML(ByRef FileName As String)
        'Dim oXmlDoc As Xml.XmlDocument
        'oXmlDoc = New Xml.XmlDocument
        Dim oXmlDoc As XmlDocument
        oXmlDoc = New XmlDocument()
        '// load the content of the XML File
        Dim sPath As String
        sPath = IO.Directory.GetParent(Application.ExecutablePath).ToString
        ExePath = sPath
        oXmlDoc.Load(sPath & "\" & FileName)
        '// load the form to the SBO application in one batch
        objApplication.LoadBatchActions(oXmlDoc.InnerXml)
        sPath = objApplication.GetLastBatchResults()
    End Sub
#End Region


#Region "Item Event"
    Private Sub objApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles objApplication.ItemEvent
        Try
            '------------------------------------------------------------------------
            Try
                If TenXFrameWork.MainCls.ohtLookUpForm.ContainsValue(FormUID) = True Then
                    Dim keys As ICollection = TenXFrameWork.MainCls.ohtLookUpForm.Keys
                    Dim keysArray(TenXFrameWork.MainCls.ohtLookUpForm.Count - 1) As String
                    keys.CopyTo(keysArray, 0)
                    For Each key As String In keysArray
                        If FormUID = TenXFrameWork.MainCls.ohtLookUpForm(key) Then
                            While TenXFrameWork.MainCls.ohtLookUpForm.ContainsValue(key) = True
                                For Each dKey As String In keysArray
                                    If key = TenXFrameWork.MainCls.ohtLookUpForm(dKey) Then
                                        key = dKey
                                        Exit For
                                    End If
                                Next
                            End While
                            objMain.objApplication.Forms.Item(key).Select()
                            BubbleEvent = False
                            Exit Sub
                        End If
                    Next
                End If
            Catch ex As Exception
            End Try
            Select Case pVal.FormTypeEx
                'Addon Files
                'Case "FORCAST"
                '    ObjFORCAST.ItemEvent(FormUID, pVal, BubbleEvent)
                'Case "TNX_FSLC"
                '    oClsParameterSelection.ItemEvent(FormUID, pVal, BubbleEvent)
                'Case "TNX_USR"
                '    oSubParameterSelection.ItemEvent(FormUID, pVal, BubbleEvent)
                'Case "TNX_OAFC"
                '    oApprovedForecating.ItemEvent(FormUID, pVal, BubbleEvent
                Case "133"
                    objARInvoice.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "179"
                    objARCreditMemo.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "65300"
                    objARDownPayment.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frm_Approve"
                    ObjVIEW.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "198"
                    If objSAPAlertWindow IsNot Nothing Then
                        objSAPAlertWindow.ItemEvent(FormUID, pVal, BubbleEvent)
                    End If
                'Case "TENXMENU_Form"
                '    objMenusettings.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "ME_ASR"
                    ObjApproval.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GRID"
                    ObjGRID.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_ADR"
                    ObjADR.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "SBO_AST"
                    ObjAStg.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GRIDES"
                    ObjGRIDES.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "SBO_Draft"
                    ObjDraftProcedure.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "Temp"
                    ObjAPTEMP.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "IK_ESTMT"
                    objEstimation.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "INV_P"
                    objInvPost.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "ONBP"
                    objOnboarding.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "PAYR"
                    objPayLoad.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "133"
                    objAREinvoice.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "CTAXC"
                    ObjCorporateTaxConfiguration.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "CTAXCAL"
                    ObjCorporateTaxCalculation.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "VATR"
                    ObjVatReport.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "PAYLD"
                    ObjPayloadD.ItemEvent(FormUID, pVal, BubbleEvent)
                    'vsm
                Case "COTAX"
                    ObjClsCorpTax.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frm_FTAVM"
                    ObjclsFtaVat.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frm_LKMTR"
                    ObjClsLkMstr.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "License"
                    objLicenceNew.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "DEVICE"
                    objDevice.ItemEvent(FormUID, pVal, BubbleEvent)
                    'Case "License"
                    '    objLicenceAdministration.ItemEvent(FormUID, pVal, BubbleEvent)


            End Select
        Catch ex As Exception
            objApplication.MessageBox(ex.Message)
        End Try
    End Sub
#End Region

#Region "Menu Events"
    Private Sub objApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles objApplication.MenuEvent
        'Try
        Dim objform As SAPbouiCOM.Form
        ' Catch ex As Exception

        ' End Try

        Try
            '  objform = objMain.objApplication.Forms.ActiveForm
            Select Case pVal.MenuUID

                'Case "FORCAST"
                '    ObjFORCAST.MenuEvent(pVal, BubbleEvent)
                'Case "TNX_FSLC"
                '    oClsParameterSelection.MenuEvent(pVal, BubbleEvent)
                'Case "TNX_USR"
                '    oSubParameterSelection.MenuEvent(pVal, BubbleEvent)
                'Case "TNX_OAFC"
                '    oApprovedForecating.MenuEvent(pVal, BubbleEvent)
                'Case "SUBTYPE"
                '    ObjSUBTYPE.MenuEvent(pVal, BubbleEvent)
                '    'Find

                Case "IK_ESTMT"
                    objEstimation.MenuEvent(pVal, BubbleEvent)

                Case "Inv_Posting"
                    objInvPost.MenuEvent(pVal, BubbleEvent)

                'Case "TENXMENU"
                '    objMenusettings.MenuEvent(pVal, BubbleEvent)

                Case "On_Process"
                    objOnboarding.MenuEvent(pVal, BubbleEvent)
                Case "PAYR"
                    objPayLoad.MenuEvent(pVal, BubbleEvent)
                Case "CTAXC"
                    ObjCorporateTaxConfiguration.MenuEvent(pVal, BubbleEvent)
                Case "CTAXCAL"
                    ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)
                Case "VATR"
                    ObjVatReport.MenuEvent(pVal, BubbleEvent)
                Case "PAYLD"
                    ObjPayloadD.MenuEvent(pVal, BubbleEvent)
                Case "COTX"
                    ObjClsCorpTax.MenuEvent(pVal, BubbleEvent)
                Case "FTAV"
                    ObjclsFtaVat.MenuEvent(pVal, BubbleEvent)
                       'Approval
                Case "ME_ASR"
                    ObjApproval.MenuEvent(pVal, BubbleEvent)
                Case "GRID"
                    'ObjGRID.MenuEvent(pVal, BubbleEvent)
                Case "ME_ADR"
                    ObjADR.MenuEvent(pVal, BubbleEvent)
                Case "SBO_AST"
                    ObjAStg.MenuEvent(pVal, BubbleEvent)
                Case "Temp"
                    ObjAPTEMP.MenuEvent(pVal, BubbleEvent)
                Case "GRIDES"
                    ObjGRIDES.MenuEvent(pVal, BubbleEvent)
                    'vsm
                Case "COTX"
                    ObjClsCorpTax.MenuEvent(pVal, BubbleEvent)
                Case "FTAV"
                    ObjclsFtaVat.MenuEvent(pVal, BubbleEvent)
                Case "LKMT"
                    ObjClsLkMstr.MenuEvent(pVal, BubbleEvent)
                Case "License"
                    objLicenceNew.MenuEvent(pVal, BubbleEvent)
                     'Case "License"
                '    objLicenceAdministration.MenuEvent(pVal, BubbleEvent)
                    'Vamshi Sai
                Case "Invoice_Posting"
                    objInvoicePosting.MenuEvent(pVal, BubbleEvent)
                Case "DEVICE"
                    objDevice.MenuEvent(pVal, BubbleEvent)

                Case "1282"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "133" Then
                        objARInvoice.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VATR" Then
                        ObjVatReport.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "CTAXCAL" Then
                        ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "COTX" Then
                        ObjClsCorpTax.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "FTAV" Then
                        ObjclsFtaVat.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "LKMT" Then
                        ObjClsLkMstr.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "License" Then
                        objLicenceNew.MenuEvent(pVal, BubbleEvent)
                    End If


                Case "1281"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "CTAXC" Then
                        ObjCorporateTaxConfiguration.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "CTAXCAL" Then
                        ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)

                    End If
                    'Navigations
                    'Case "1288"
                    '    objform = objMain.objApplication.Forms.ActiveForm
                    '    If objform.TypeEx = "TNX_USR" Then
                    '        oSubParameterSelection.MenuEvent(pVal, BubbleEvent)
                    '    End If
                    '    If objform.TypeEx = "TNX_OAFC" Then
                    '        oApprovedForecating.MenuEvent(pVal, BubbleEvent)
                    '    End If
                    'Case "1289"
                    '    objform = objMain.objApplication.Forms.ActiveForm
                    '    If objform.TypeEx = "TNX_USR" Then
                    '        oSubParameterSelection.MenuEvent(pVal, BubbleEvent)
                    '    End If
                    '    If objform.TypeEx = "TNX_OAFC" Then
                    '        oApprovedForecating.MenuEvent(pVal, BubbleEvent)
                    '    End If
                    'Case "1290"
                    '    objform = objMain.objApplication.Forms.ActiveForm
                    '    If objform.TypeEx = "TNX_USR" Then
                    '        oSubParameterSelection.MenuEvent(pVal, BubbleEvent)
                    '    End If
                    '    If objform.TypeEx = "TNX_OAFC" Then
                    '        oApprovedForecating.MenuEvent(pVal, BubbleEvent)
                    '    End If
                Case "1291"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "CTAXC" Then
                        ObjCorporateTaxConfiguration.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VATR" Then
                        ObjVatReport.MenuEvent(pVal, BubbleEvent)
                        'ElseIf objform.TypeEx = "CTAXCAL" Then
                        '    ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)
                    End If
                    'If objform.TypeEx = "TNX_OAFC" Then
                    '    oApprovedForecating.MenuEvent(pVal, BubbleEvent)
                    'End If
                    '    'ADD ROW
                Case "1293"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "CTAXC" Then
                        ObjCorporateTaxConfiguration.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VATR" Then
                        ObjVatReport.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "CTAXCAL" Then
                        ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "COTX" Then
                        ObjClsCorpTax.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "FTAV" Then
                        ObjclsFtaVat.MenuEvent(pVal, BubbleEvent)
                    End If

                Case "774"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "CTAXCAL" Then
                        ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)
                    End If
                    '    'ADD ROW
                Case "1282"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "CTAXC" Then
                        ObjCorporateTaxConfiguration.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "VATR" Then
                        ObjVatReport.MenuEvent(pVal, BubbleEvent)
                        'ElseIf objform.TypeEx = "CTAXCAL" Then
                        '    ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)
                    End If

                    'Case "Add Row"
                    '    objform = objMain.objApplication.Forms.ActiveForm
                    '    If objform.TypeEx = "FORCAST" Then
                    '        ObjFORCAST.MenuEvent(pVal, BubbleEvent)
                    '    End If
                Case "519"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "CTAXC" Then
                        ObjCorporateTaxConfiguration.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VATR" Then
                        ObjVatReport.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "CTAXCAL" Then
                        ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "COTX" Then
                        ObjClsCorpTax.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "FTAV" Then
                        ObjclsFtaVat.MenuEvent(pVal, BubbleEvent)
                    End If

                Case "520"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "CTAXC" Then
                        ObjCorporateTaxConfiguration.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VATR" Then
                        ObjVatReport.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "CTAXCAL" Then
                        ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "COTX" Then
                        ObjClsCorpTax.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "FTAV" Then
                        ObjclsFtaVat.MenuEvent(pVal, BubbleEvent)
                    End If
                Case "1284"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "CTAXCAL" Then
                        ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)
                    End If

                Case "1292"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "CTAXC" Then
                        ObjCorporateTaxConfiguration.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VATR" Then
                        ObjVatReport.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "CTAXCAL" Then
                        ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "COTAX" Then
                        ObjClsCorpTax.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_FTAVM" Then
                        ObjclsFtaVat.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "Temp" Then
                        ObjAPTEMP.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "DEVICE" Then
                        objDevice.MenuEvent(pVal, BubbleEvent)

                        'ElseIf objform.TypeEx = "TNX_USR" Then
                        '    oSubParameterSelection.MenuEvent(pVal, BubbleEvent)
                    End If
                    'Case "Delete Row"
                    '    objform = objMain.objApplication.Forms.ActiveForm
                    '    If objform.TypeEx = "FORCAST" Then
                    '        ObjFORCAST.MenuEvent(pVal, BubbleEvent)
                    '    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

#Region "Form Data Event"
    Private Sub objApplication_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles objApplication.FormDataEvent
        If BusinessObjectInfo.BeforeAction = False Then
            Select Case BusinessObjectInfo.FormTypeEx
                Case "TNX_OAFC"
                    'oApprovedForecating.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "133"
                    objARInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                    objARInvoice.FormDataEvent1(BusinessObjectInfo, BubbleEvent)
                Case "179"
                    objARCreditMemo.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "CTAXCAL"
                    'ObjCorporateTaxCalculation.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "65300"
                    objARDownPayment.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "VATR"
                    ObjVatReport.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            End Select
        End If
    End Sub
#End Region
    Public Function Addonlicence(ByVal Addonname As String, ByVal DBName As String, ByVal User As String) As Boolean
        Dim rsConfig1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            If objMain.objCompany.Connected Then

                Dim query As String = "Select T0.""U_License"" as ""licence"", ""U_HKey"" as ""Hardwarekey"",T0.""U_Company"", T0.""U_DB"" As ""Dbname"",T0.""U_Addon"" as ""Addon"", T0.""U_Total"" as ""Total""," &
                                      "T1.""U_Code"" as ""User"", T1.""U_Sts"" as ""Status"", " &
                                      "T0.""U_SDate"" as ""fromdate"", T0.""U_EDate"" as ""Todate"" " &
                                      "from ""@TNX_LICENCE"" T0 " &
                                      "INNER JOIN ""@TNX_LICENCE_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
                                      "WHERE T1.""U_Sts"" = 'Y' AND T1.""U_Code"" ='" & User & "' AND T0.""U_Addon"" ='" & Addonname & "' " &
                                      "AND T0.""U_DB"" ='" & DBName & "' and CURRENT_DATE BETWEEN T0.""U_SDate"" AND T0.""U_EDate"""
                rsConfig1.DoQuery(query)

                If rsConfig1.RecordCount = 0 Then
                    objMain.objApplication.StatusBar.SetText("Invalid 10X license file, Please Contact 10X Team", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                Dim licence As String = rsConfig1.Fields.Item("licence").Value
                If String.IsNullOrEmpty(licence) Then
                    objMain.objApplication.StatusBar.SetText("Invalid 10X license file, Please Contact 10X Team", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                If rsConfig1.Fields.Item("Status").Value <> "Y" Then
                    objMain.objApplication.StatusBar.SetText("Invalid 10X license file, Please Contact 10X Team", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                Dim Company As String = rsConfig1.Fields.Item("U_Company").Value
                Dim Database As String = rsConfig1.Fields.Item("Dbname").Value
                Dim HKey As String = rsConfig1.Fields.Item("Hardwarekey").Value
                Dim TotalLicense As String = rsConfig1.Fields.Item("Total").Value
                Dim Addon As String = rsConfig1.Fields.Item("Addon").Value

                Dim StartDate As Date = CDate(rsConfig1.Fields.Item("fromdate").Value)
                Dim ExpDate As Date = CDate(rsConfig1.Fields.Item("Todate").Value)

                Dim SDate As String = StartDate.ToString("yyyy/MM/dd")
                Dim EDate As String = ExpDate.ToString("yyyy/MM/dd")

                Dim uniKey As String = "10X_SAP_B!_"
                Dim strKey As String = StartDate.ToString("dd-MMM-yyyy") & "|" & ExpDate.ToString("dd-MMM-yyyy") & "|" & Company & "|" & TotalLicense & "|" & Database & "|" & HKey & "|" & Addonname
                Dim generatedLicense As String = GetSHA1HashData(strKey + GetSHA1HashData(uniKey))

                If licence <> generatedLicense Then
                    objMain.objApplication.StatusBar.SetText("Invalid license file, Please Contact 10X Team", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                Return True
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("10X License check error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            If rsConfig1 IsNot Nothing Then
                Marshal.ReleaseComObject(rsConfig1)
                rsConfig1 = Nothing
            End If
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try

    End Function
    Private Function GetSHA1HashData(ByVal data As String) As String
        Dim sha2 As SHA1 = sha2.Create()
        Dim hashData As Byte() = sha2.ComputeHash(Encoding.[Default].GetBytes(data))
        Dim returnValue As System.Text.StringBuilder = New StringBuilder()
        For i As Integer = 0 To hashData.Length - 1
            returnValue.Append(hashData(i).ToString())
        Next
        Return returnValue.ToString()
    End Function
    Sub CheckLicenceMenuEnableDisable()

        Try

            'Reload Menu XML first
            Me.LoadFromXML("Menu.xml")

            Dim rs As SAPbobsCOM.Recordset
            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim Qry As String =
        "SELECT TOP 1 " &
        "IFNULL(""U_EIVC"",'N') AS ""EINV"", " &
        "IFNULL(""U_CTTX"",'N') AS ""CTAX"", " &
        "IFNULL(""U_VRPT"",'N') AS ""VATR"" " &
        "FROM ""@TNX_LICENCE"" " &
        "WHERE ""U_DB"" = '" & objMain.objCompany.CompanyDB & "' " &
        "ORDER BY ""DocEntry"" DESC"

            rs.DoQuery(Qry)

            If rs.RecordCount = 0 Then

                DisableEInvoiceMenus()
                DisableCorporateTaxMenus()
                DisableVatReportMenus()

                Exit Sub

            End If

            Dim IsEInvoice As String =
        rs.Fields.Item("EINV").Value.ToString.Trim

            Dim IsCorpTax As String =
        rs.Fields.Item("CTAX").Value.ToString.Trim

            Dim IsVatReport As String =
        rs.Fields.Item("VATR").Value.ToString.Trim

            If IsEInvoice <> "Y" Then
                DisableEInvoiceMenus()
            End If

            If IsCorpTax = "Y" Then

                EnableCorporateTaxMenus()

            Else

                DisableCorporateTaxMenus()

            End If

            If IsVatReport <> "Y" Then
                DisableVatReportMenus()
            End If

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
        "Licence menu check error : " & ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
    Private Sub EnableCorporateTaxMenus()

        Try

            Dim oMenus As SAPbouiCOM.Menus
            Dim oMenuItem As SAPbouiCOM.MenuItem
            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams

            oMenus = objMain.objApplication.Menus
            oMenuItem = oMenus.Item("15616")

            'CTAXC
            If Not oMenus.Exists("CTAXC") Then

                oCreationPackage =
            objMain.objApplication.CreateObject(
            SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "CTAXC"
                oCreationPackage.String = "Corporate Tax"
                oCreationPackage.Enabled = True
                oCreationPackage.Position = 5

                oMenuItem.SubMenus.AddEx(oCreationPackage)

            End If

            'COTX
            If Not oMenus.Exists("COTX") Then

                oCreationPackage =
            objMain.objApplication.CreateObject(
            SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "COTX"
                oCreationPackage.String = "Corporate Tax Calendar"
                oCreationPackage.Enabled = True
                oCreationPackage.Position = 6

                oMenuItem.SubMenus.AddEx(oCreationPackage)

            End If

            'FTAV
            If Not oMenus.Exists("FTAV") Then

                oCreationPackage =
            objMain.objApplication.CreateObject(
            SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "FTAV"
                oCreationPackage.String = "FTA VAT Calendar"
                oCreationPackage.Enabled = True
                oCreationPackage.Position = 7

                oMenuItem.SubMenus.AddEx(oCreationPackage)

            End If

            'LKMT
            If Not oMenus.Exists("LKMT") Then

                oCreationPackage =
            objMain.objApplication.CreateObject(
            SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "LKMT"
                oCreationPackage.String = "Link Master"
                oCreationPackage.Enabled = True
                oCreationPackage.Position = 8

                oMenuItem.SubMenus.AddEx(oCreationPackage)

            End If

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub
    'Sub CheckLicenceMenuEnableDisable()

    '    Try
    '        Dim rs As SAPbobsCOM.Recordset
    '        rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        Dim Qry As String =
    '        "SELECT TOP 1 " &
    '        "IFNULL(""U_EIVC"",'N') AS ""EINV"", " &
    '        "IFNULL(""U_CTTX"",'N') AS ""CTAX"", " &
    '        "IFNULL(""U_VRPT"",'N') AS ""VATR"" " &
    '        "FROM ""@TNX_LICENCE"" " &
    '        "WHERE ""U_DB"" = '" & objMain.objCompany.CompanyDB & "' " &
    '        "ORDER BY ""DocEntry"" DESC"

    '        rs.DoQuery(Qry)

    '        If rs.RecordCount = 0 Then
    '            DisableEInvoiceMenus()
    '            DisableCorporateTaxMenus()
    '            DisableVatReportMenus()
    '            Exit Sub
    '        End If

    '        Dim IsEInvoice As String = rs.Fields.Item("EINV").Value.ToString().Trim()
    '        Dim IsCorpTax As String = rs.Fields.Item("CTAX").Value.ToString().Trim()
    '        Dim IsVatReport As String = rs.Fields.Item("VATR").Value.ToString().Trim()

    '        If IsEInvoice <> "Y" Then DisableEInvoiceMenus()
    '        If IsCorpTax <> "Y" Then DisableCorporateTaxMenus()
    '        If IsVatReport <> "Y" Then DisableVatReportMenus()

    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText("Licence menu check error : " & ex.Message)
    '    End Try

    'End Sub
    Private Sub DisableEInvoiceMenus()
        DisableMenu("Invoice_Posting")
        DisableMenu("PAYLD")
    End Sub

    Private Sub DisableCorporateTaxMenus()
        DisableMenu("CTAXC")
        DisableMenu("COTX")
        DisableMenu("FTAV")
        DisableMenu("LKMT")
        DisableMenu("CTAXCAL")
        DisableMenu("Load_List")
    End Sub

    Private Sub DisableVatReportMenus()
        DisableMenu("VATR")
    End Sub

    Private Sub DisableMenu(ByVal MenuUID As String)
        Try
            If objMain.objApplication.Menus.Exists(MenuUID) Then
                objMain.objApplication.Menus.RemoveEx(MenuUID)
            End If
        Catch ex As Exception
        End Try
    End Sub

#Region "Application Event"
    Private Sub oApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles objApplication.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged, SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                objCompany.Disconnect()
                End
        End Select
    End Sub
#End Region

#Region "Right Click Event"
    Private Sub objApplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles objApplication.RightClickEvent
        Dim objForm As SAPbouiCOM.Form
        objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)
        'If objForm.TypeEx = "FORCAST" Then
        '    'ObjFORCAST.RightClickEvent(eventInfo, BubbleEvent)
        'End If
    End Sub
#End Region

End Class
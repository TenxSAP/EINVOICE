Imports System.ServiceModel
Imports SAPbobsCOM

Public Class ClsVIEWFORAPP

#Region "       Declaration             "
    Public objForm, objSTRForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim str, str1, str2, str3, Query, str4, Query2 As String
    Public rs, RsNum As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False
    Dim objutilities As Utilities
    Public rsGetLeaveDetails, rsUpdateEmpLeaveQuota As SAPbobsCOM.Recordset
    Dim oDt As SAPbouiCOM.DataTable
    Dim oGrid, oSubGrid As SAPbouiCOM.Grid
    Dim rs1, rs2, rs3, rs4 As SAPbobsCOM.Recordset
    Dim column, column1 As SAPbouiCOM.EditTextColumn
    Public AppIDV, StageV, qrystring, DocumentType As String
    Public TemplateID, Table, AppIDField, AppStatField, OriginatorV As String
    Public oCmpSrv As SAPbobsCOM.CompanyService
    Public oMessageService As SAPbobsCOM.MessagesService

#End Region
    Sub CreateForm(ByVal Code As String, ByVal Stage As String)

        objMain.objUtilities.LoadForm("VIEWW.xml", "frm_Approve", ResourceType.Embeded)
        objForm = objMain.objApplication.Forms.GetForm("frm_Approve", objMain.objApplication.Forms.ActiveForm.TypeCount)
        AppIDV = Code
        StageV = Stage
        'objForm.Items.Item("E2").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        'objForm.Items.Item("E2").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        objForm.Items.Item("docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

        Me.createViewForApproval(Stage, objForm)
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Select Case pVal.EventType
            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                objForm = objMain.objApplication.Forms.Item(FormUID)

                'link button  for approval screen to supplier rebate and Price change Request PCR
                Dim DocType As String = objForm.Items.Item("Doc").Specific.Value



                If pVal.ItemUID = "LB_PI" And pVal.BeforeAction = False And DocType = "Corporate Tax" Then

                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    Dim DocNum As String = objForm.Items.Item("docnum").Specific.Value
                    'objMain.objSupplierRebate.CreateForm(FormUID, "", "", "", DocNum, "", "")

                    'objMain.ObjPriceChangeRequest.CreateForm(DocNum)


                    objMain.objApplication.ActivateMenuItem("CTAXCAL")


                    Dim objNewForm As SAPbouiCOM.Form
                    objNewForm = objMain.objApplication.Forms.ActiveForm
                    objNewForm.Freeze(True)
                    objNewForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    objNewForm.Items.Item("AIPD").Specific.Value = DocNum
                    objNewForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    objNewForm.Freeze(False)

                End If


                If pVal.ItemUID = "LB_PI" And pVal.BeforeAction = False And DocType = "VAT Report" Then

                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    Dim DocNum As String = objForm.Items.Item("docnum").Specific.Value
                    'objMain.objSupplierRebate.CreateForm(FormUID, "", "", "", DocNum, "", "")

                    'objMain.ObjPriceChangeRequest.CreateForm(DocNum)


                    objMain.objApplication.ActivateMenuItem("VATR")


                    Dim objNewForm As SAPbouiCOM.Form
                    objNewForm = objMain.objApplication.Forms.ActiveForm
                    objNewForm.Freeze(True)
                    objNewForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    objNewForm.Items.Item("APPI").Specific.Value = DocNum
                    objNewForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    objNewForm.Freeze(False)

                End If

                If pVal.BeforeAction = False And pVal.ItemUID = "LB_PI" And DocType = "PriceChangeRequest" Then
                    Try
                        Dim DocNumberPCR As String = objForm.Items.Item("docnum").Specific.Value
                        'objMain.ObjPriceChangeRequest.CreateForm_Find(DocNumberPCR)
                    Catch ex As Exception
                    End Try
                End If


                If pVal.BeforeAction = False And pVal.ItemUID = "LB_PI" And DocType = "" Then


                    Dim objGRForm As SAPbouiCOM.Form

                    Try

                        Dim DocNumberItemMaster As String = objForm.Items.Item("docnum").Specific.Value

                        objMain.objApplication.ActivateMenuItem("CTAXCAL")

                        objGRForm = objMain.objApplication.Forms.ActiveForm

                        objGRForm.Freeze(True)

                        objGRForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE

                        objGRForm.Items.Item("AIPD").Specific.Value = DocNumberItemMaster

                        objGRForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)



                        'oDS.SetValue("U_PRSTAT", oDS.Offset, "")

                        objGRForm.Freeze(False)

                    Catch ex As Exception

                        objGRForm.Freeze(False)

                    End Try


                End If

                'Item Listing
                If pVal.BeforeAction = False And pVal.ItemUID = "LB_PI" And DocType = "ItemListing" Then


                    Dim objGRForm As SAPbouiCOM.Form

                    Try
                        Dim DocNumberItemListing As String = objForm.Items.Item("docnum").Specific.Value
                        'objMain.objItemList.CreateForm_Find(DocNumberItemListing)
                    Catch ex As Exception
                    End Try

                    Dim DocNumberItemMaster As String = objForm.Items.Item("docnum").Specific.Value
                    Dim APPID As String = objForm.Items.Item("docnum").Specific.Value
                    'Try

                    'objMain.objApplication.ActivateMenuItem("ME_ITMLIST")
                    '    Dim DocNumberItemMaster As String = objForm.Items.Item("docnum").Specific.Value

                    'objGRForm = objMain.objApplication.Forms.ActiveForm
                    '    objMain.objApplication.ActivateMenuItem("ME_ITMLIST")

                    'objGRForm.Freeze(True)

                    'objGRForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE

                    'objGRForm.Items.Item("AppID").Specific.Value = DocNumberItemMaster

                    'objGRForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                    'Dim oDS As SAPbouiCOM.DBDataSource
                    '    objGRForm = objMain.objApplication.Forms.ActiveForm

                    'oDS = objGRForm.DataSources.DBDataSources.Item("@SBO_LIST")
                    '    objGRForm.Freeze(True)

                    ''oDS.SetValue("U_PRSTAT", oDS.Offset, "")
                    '    objGRForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE

                    '    objGRForm.Items.Item("AppID").Specific.Value = DocNumberItemMaster

                    'objGRForm.Freeze(False)
                    If APPID <> "" Then
                        '    objGRForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        objMain.objApplication.StatusBar.SetText("Opening Item list data , please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        '    Dim oDS As SAPbouiCOM.DBDataSource

                        '  objMain.objItemList.CreateForm(objForm.UniqueID, "", "", "Y", APPID)
                        '    oDS = objGRForm.DataSources.DBDataSources.Item("@SBO_LIST")

                        '    'oDS.SetValue("U_PRSTAT", oDS.Offset, "")

                    End If
                    '    objGRForm.Freeze(False)

                    'Catch ex As Exception

                    '    objGRForm.Freeze(False)

                    'End Try


                End If

                'Business partner
                If pVal.BeforeAction = False AndAlso pVal.ItemUID = "LB_PI" AndAlso DocType = "BusinessPartner" Then

                    Dim objBPForm As SAPbouiCOM.Form = Nothing

                    Try
                        Dim CardCode As String = objForm.Items.Item("docnum").Specific.Value
                        objMain.objApplication.ActivateMenuItem("2561")
                        objBPForm = objMain.objApplication.Forms.ActiveForm
                        objBPForm.Freeze(True)
                        objBPForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objBPForm.Items.Item("5").Specific.Value = CardCode
                        objBPForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objBPForm.Freeze(False)
                    Catch ex As Exception
                        If objBPForm IsNot Nothing Then
                            objBPForm.Freeze(False)
                        End If
                        objMain.objApplication.StatusBar.SetText(ex.Message)
                    End Try

                End If



                If pVal.ItemUID = "5" And pVal.ColUID = "0" And pVal.BeforeAction = False Then
                    ' Me.CreateForm(FormUID)
                End If
                'If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                '    If Me.Validate(FormUID) = False Then

                '        BubbleEvent = False
                '    End If
                'End If
                If pVal.ItemUID = "1" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And pVal.BeforeAction = True Then

                    If Me.Validate(FormUID) = False Then
                        BubbleEvent = False
                        Exit Sub   ' <<< THIS will stop entire execution
                    End If

                    BubbleEvent = False
                    Dim objAppStatusCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("statusN").Specific
                    Dim Status As String = objAppStatusCombo.Selected.Value
                    OriginatorV = objForm.Items.Item("userid").Specific.Value.ToString().Trim()
                    Dim Remarks As String = objForm.Items.Item("Remarks").Specific.Value.ToString().Trim()

                    If Status = "A" Or Status = "N" Then
                        If updateDraftTable(Status, TemplateID, Remarks, objMain.objCompany.UserName, AppIDV) = True Then
                            objMain.objApplication.StatusBar.SetText("Operation Completed Sucessfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                        End If
                    Else
                        objMain.objApplication.StatusBar.SetText("Operation Completed Sucessfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                    End If
                End If


        End Select
    End Sub
    Public Function Validate(ByVal FormUID As String) As Boolean
        ' Dim CustomerCode As SAPbouiCOM.Matrix
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim APPID As String = objForm.Items.Item("docnum").Specific.Value
            Dim TEMID As String = objForm.Items.Item("tempid").Specific.Value

            rs4 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim QRYS As String = "Select ""U_FOUT"" from ""@TNX_PCRT"" where ""U_AppID"" = '" & APPID & "'"
            rs4.DoQuery(QRYS)

            If TEMID = "FLIER" Then
                If rs4.Fields.Item("U_FOUT").Value.ToString() = "No" Then

                    objMain.objApplication.SetStatusBarMessage("Please change the Flyer-Out status to Yes before approving the document.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False

                End If
            End If

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message & "Errors In Validation Function", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function
    Public Sub createViewForApproval(ByVal Stage As String, ByVal objForm As SAPbouiCOM.Form)

        Try
            '--- Load Document in Find Mode ---
            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            objForm.Items.Item("docnum").Specific.Value = AppIDV
            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            '--- Get Header + Current Stage Status ---
            rs1 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim Qry1 As String =
            "SELECT T0.""U_statusN"" as ""Status"", " &
            "       T1.""U_statusN"" as ""StageStatus"", " &
            "       T0.""U_tempid"", T0.""U_Table"", T0.""U_AppID"", T0.""U_AppStat"" " &
            "FROM ""@SBO_DD"" T0 " &
            "INNER JOIN ""@SBO_DD1"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
            "WHERE T0.""U_docnum"" = '" & AppIDV & "' " &
            "AND T1.""U_Userid"" = '" & objMain.objCompany.UserName & "' " &
            "AND T1.""U_Stage"" = '" & Stage & "'"

            rs1.DoQuery(Qry1)

            If rs1.RecordCount = 0 Then
                objMain.objApplication.StatusBar.SetText(
                "No Approval Data Found",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            rs1.MoveFirst()

            '--- Assign Values ---
            TemplateID = rs1.Fields.Item("U_tempid").Value
            Table = rs1.Fields.Item("U_Table").Value
            AppIDField = rs1.Fields.Item("U_AppID").Value
            AppStatField = rs1.Fields.Item("U_AppStat").Value

            Dim StageStatus As String = rs1.Fields.Item("StageStatus").Value.ToString()
            Dim Status As String = rs1.Fields.Item("Status").Value.ToString()

            objForm.Items.Item("Stage").Specific.Value = Stage

            '--- Load Approval History Grid ---
            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim QueryString As String =
            "SELECT T0.""U_Stage"" As ""Approval Stage"", " &
            "       T0.""U_Userid"" As ""UserID"", " &
            "CASE " &
            "   WHEN T0.""U_statusN""='S' THEN 'Submitted' " &
            "   WHEN T0.""U_statusN""='A' THEN 'Approved' " &
            "   WHEN T0.""U_statusN""='N' THEN 'Cancelled' " &
            "   WHEN T0.""U_statusN""='O' THEN 'Open' " &
            "END As ""Status"", " &
            "T0.""U_cdate"" As ""CreateDate"", " &
            "T0.""U_ctime"" As ""CreateTime"", " &
            "T0.""U_Udate"" As ""UpdateDate"", " &
            "T0.""U_Utime"" As ""UpdateTime"", " &
            "T0.""U_Remarks"" As ""Remarks"" " &
            "FROM ""@SBO_DD1"" T0 " &
            "WHERE T0.""DocEntry"" = (SELECT ""DocEntry"" FROM ""@SBO_DD"" WHERE ""U_docnum"" = '" & AppIDV & "') " &
            "ORDER BY T0.""LineId"" ASC"

            '--- Bind Grid ---
            If objForm.DataSources.DataTables.Count = 0 Then
                objForm.DataSources.DataTables.Add("DT_0")
            End If

            objForm.DataSources.DataTables.Item("DT_0").ExecuteQuery(QueryString)

            Dim objGrid As SAPbouiCOM.Grid = objForm.Items.Item("Grid").Specific
            objGrid.DataTable = objForm.DataSources.DataTables.Item("DT_0")
            objGrid.CollapseLevel = 1

            For i As Integer = 0 To objGrid.DataTable.Columns.Count - 1
                objGrid.Columns.Item(i).Editable = False
            Next

            objForm.Visible = True

            '--- Handle UI Based on Status ---
            Dim objStatusCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("E1").Specific

            If StageStatus = "S" And Status <> "A" Then
                'User can approve/reject
                objForm.Items.Item("1").Enabled = True

            ElseIf StageStatus = "R" Then
                'Rejected
                objStatusCombo.Select(2, SAPbouiCOM.BoSearchKey.psk_Index)
                objForm.Items.Item("1").Enabled = False
                objForm.Items.Item("Remarks").Enabled = False
            ElseIf StageStatus = "N" Then
                'cancel
                objStatusCombo.Select(4, SAPbouiCOM.BoSearchKey.psk_Index)
                objForm.Items.Item("1").Enabled = False
                objForm.Items.Item("Remarks").Enabled = False
            Else
                'Approved or already processed
                objStatusCombo.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
                objForm.Items.Item("1").Enabled = False
                objForm.Items.Item("Remarks").Enabled = False
            End If

        Catch ex As Exception
            objMain.objApplication.SetStatusBarMessage(
            ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            True)
        End Try

    End Sub
    'Public Sub createViewForApproval(ByVal Stage As String, ByVal objForm As SAPbouiCOM.Form)
    '    Try
    '        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
    '        objForm.Items.Item("docnum").Specific.Value = AppIDV
    '        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
    '        rs1 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        'Dim Qry1 As String = "Select * ,T0.""U_statusN""As ""Status"", T1.""U_statusN""As""StageStatus"" From ""@SBO_OWDD"" T0 INNER JOIN ""@SBO_WDD1"" T1 On T0.""DocEntry""=T1.""DocEntry"" Where T0.""U_docnum""='" & AppIDV & "' and T1.""U_Userid""='" & objMain.objCompany.UserName & "' and T1.""U_Stage""='" & Stage & "' "
    '        Dim Qry1 As String = "Select * ,T0.""U_statusN""as ""Status"", T1.""U_statusN""as""StageStatus"" From ""@SBO_DD"" T0 INNER JOIN ""@SBO_DD1"" T1 ON T0.""DocEntry""=T1.""DocEntry"" Where T0.""U_docnum""='" & AppIDV & "' and T1.""U_Userid""='" & objMain.objCompany.UserName & "' and T1.""U_Stage""='" & Stage & "' "
    '        rs1.DoQuery(Qry1)

    '        TemplateID = rs1.Fields.Item("U_tempid").Value
    '        Table = rs1.Fields.Item("U_Table").Value
    '        AppIDField = rs1.Fields.Item("U_AppID").Value
    '        AppStatField = rs1.Fields.Item("U_AppStat").Value
    '        Dim StageStatus As String = rs1.Fields.Item("StageStatus").Value
    '        Dim Status As String = rs1.Fields.Item("Status").Value
    '        objForm.Items.Item("Stage").Specific.Value = Stage

    '        'Me.FormItem(enControlName.sta_AnswerDetails).TextStyle = 5
    '        'Me.FormItem(enControlName.sta_DocumentDetails).TextStyle = 5
    '        'Me.FormItem("27").TextStyle = 5
    '        'Me.FormItem("txt5").TextStyle = 5


    '        rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        'Dim QueryString As String = "SELECT T0.""U_Stage"" As ""Approval Stage"", T0.""U_Userid"" As ""UserID"" " &
    '        '    " , Case When T0.""U_statusN""='S' Then 'Submitted' When T0.""U_statusN""='A' Then 'Approved' When T0.""U_statusN""='N' Then 'Rejected'When T0.""U_statusN""='O' Then 'Open' End As ""Status"" " &
    '        '    " , T0.""U_cdate"" As ""CreateDate"", T0.""U_ctime"" As ""CreateTime"", T0.""U_Udate"" As ""UpdateDate"", T0.""U_Utime"" As ""UpdateTime"", T0.""U_Remarks"" As ""Remarks"" FROM ""@SBO_WDD1""  T0 WHERE T0.""DocEntry"" =(Select ""DocEntry"" From ""@SBO_OWDD"" Where ""U_docnum""='" & AppIDV & "') Order by T0.""LineId"" asc"

    '        Dim QueryString As String = "SELECT T0.""U_Stage"" As ""Approval Stage"", T0.""U_Userid"" As ""UserID"" " &
    '            " , Case When T0.""U_statusN""='S' Then 'Submitted' When T0.""U_statusN""='A' Then 'Approved' When T0.""U_statusN""='N' Then 'Rejected'When T0.""U_statusN""='O' Then 'Open' End As ""Status"" " &
    '            " , T0.""U_cdate"" As ""CreateDate"", T0.""U_ctime"" As ""CreateTime"", T0.""U_Udate"" As ""UpdateDate"", T0.""U_Utime"" As ""UpdateTime"", T0.""U_Remarks"" As ""Remarks"" FROM ""@SBO_DD1""  T0 WHERE T0.""DocEntry"" =(Select ""DocEntry"" From ""@SBO_DD"" Where ""U_docnum""='" & AppIDV & "') Order by T0.""LineId"" asc"

    '        If objForm.DataSources.DataTables.Count() = 0 Then
    '            objForm.DataSources.DataTables.Add("DT_0")
    '        End If
    '        objForm.DataSources.DataTables.Item(0).ExecuteQuery("" & QueryString & "")
    '        Dim objGrid As SAPbouiCOM.Grid = objForm.Items.Item("Grid").Specific
    '        objGrid.DataTable = objForm.DataSources.DataTables.Item("DT_0")
    '        objGrid.CollapseLevel = 1
    '        For i As Integer = 0 To objGrid.DataTable.Columns.Count - 1
    '            objGrid.Columns.Item(i).Editable = False
    '        Next
    '        Dim objStatusCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("E1").Specific
    '        objForm.Visible = True
    '        If StageStatus = "S" And Status <> "A" Then
    '            objForm.Items.Item("1").Enabled = True
    '        ElseIf StageStatus = "N" Then
    '            objStatusCombo.Select(2, SAPbouiCOM.BoSearchKey.psk_Index)
    '            objForm.Items.Item("1").Enabled = False
    '            objForm.Items.Item("Remarks").Enabled = False
    '        Else
    '            objStatusCombo.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
    '            objForm.Items.Item("1").Enabled = False
    '            objForm.Items.Item("Remarks").Enabled = False
    '        End If

    '    Catch ex As Exception
    '        objMain.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
    '    Finally

    '    End Try
    'End Sub
    Public Function updateDraftTable(ByVal Status As String, ByVal tempid As String, ByVal Remarks As String, ByVal authorizer As String, ByVal APPIDV As String)
        Try
            Dim oRsDraftDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'Dim DraftDetQry As String = "Select * From ""@SBO_OWDD"" T0 INNER JOIN ""@SBO_WDD1"" T1 ON T0.""DocEntry""=T1.""DocEntry"" " &
            '"Where T0.""U_docnum""='" & APPIDV & "' and T1.""U_Userid""='" & authorizer & "' and T1.""U_Stage""='" & StageV & "' "
            'oRsDraftDetails.DoQuery(DraftDetQry)

            Dim DraftDetQry As String = "Select * From ""@SBO_DD"" T0 INNER JOIN ""@SBO_DD1"" T1 ON T0.""DocEntry""=T1.""DocEntry"" " &
            "Where T0.""U_docnum""='" & APPIDV & "' and T1.""U_Userid""='" & authorizer & "' and T1.""U_Stage""='" & StageV & "' "
            oRsDraftDetails.DoQuery(DraftDetQry)

            Table = oRsDraftDetails.Fields.Item("U_Table").Value
            AppIDField = oRsDraftDetails.Fields.Item("U_AppID").Value
            AppStatField = oRsDraftDetails.Fields.Item("U_AppStat").Value
            OriginatorV = oRsDraftDetails.Fields.Item("U_userid").Value
            DocumentType = oRsDraftDetails.Fields.Item("U_objtype").Value

            Dim rsup1, rsup2, rsup3, oRsDraftMaster, oRsDocUpdate As SAPbobsCOM.Recordset
            Dim st, strUpdate As String
            Dim appcnt, totcnt As Integer
            rsup1 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsup2 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsup3 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsDraftMaster = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsDocUpdate = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '--------Current Stage counts
            Dim oRsCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim CountQry As String = "SELECT IFNULL(T0.""U_NAP"",0) As ""NAP"", IFNULL(T0.""U_NRJ"",0) As ""NRJ"" FROM ""@SBO_AST""  T0 WHERE T0.""Code"" ='" & StageV & "'"
            oRsCount.DoQuery(CountQry)
            Dim NoOfApprovals As Double = oRsCount.Fields.Item("NAP").Value
            Dim NoOfRejects As Double = oRsCount.Fields.Item("NRJ").Value

            'If Table = "@SBO_RESUMESUBM" Then
            '    Dim strStatus As String = " Select IFNULL(R2.""U_Apprve"",'') ""ApproveStatus"" from ""@SBO_RESUMESUBM"" R1 INNER JOIN ""@SBO_RESUMESUBM_C0"" R2 " &
            '    " ON R1.""DocEntry"" = R2.""DocEntry"" where R2.""U_CCode"" <> '' and R1.""U_RSNO"" = '" & AppIDV & "' and R2.""U_Apprve"" = '' "
            '    Dim rsStatus As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '    rsStatus.DoQuery(strStatus)
            '    If rsStatus.RecordCount > 0 And rsStatus.Fields.Item("ApproveStatus").Value = "" Then
            '        objMain.objApplication.StatusBar.SetText("Please Select Candidate status before Approve/Reject", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '        'Me.FormItem("RSRcNo").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            '        Return False
            '        Exit Function
            '    End If
            'End If

            If Status = "A" Then
                'strUpdate = "UPDATE ""@SBO_WDD1"" SET ""U_statusN""='A',""U_Remarks""='" & Remarks & "',""U_Udate""='" & DateTime.Now.ToString("yyyyMMdd") & "',""U_Utime""='" & DateTime.Now.ToString("HHmm") & "' " &
                '    " where ""U_Userid""='" & objMain.objCompany.UserName & "' and ""U_Stage""='" & StageV & "' and ""DocEntry""=(select ""DocEntry"" from ""@SBO_OWDD"" where ""U_docnum""='" & APPIDV & "')"
                'rsup1.DoQuery(strUpdate)

                strUpdate = "UPDATE ""@SBO_DD1"" SET ""U_statusN""='A',""U_Remarks""='" & Remarks & "',""U_Udate""='" & DateTime.Now.ToString("yyyyMMdd") & "',""U_Utime""='" & DateTime.Now.ToString("HHmm") & "' " &
                    " where ""U_Userid""='" & objMain.objCompany.UserName & "' and ""U_Stage""='" & StageV & "' and ""DocEntry""=(select ""DocEntry"" from ""@SBO_DD"" where ""U_docnum""='" & APPIDV & "')"
                rsup1.DoQuery(strUpdate)

                'Dim Qry As String = "SELECT COUNT(*) FROM ""@SBO_OWDD"" t1 INNER JOIN ""@SBO_WDD1"" t2 ON t1.""DocEntry"" = t2.""DocEntry"""
                'Qry = Qry & " WHERE ""U_docnum"" = '" & APPIDV & "' and t2.""U_statusN"" ='A' and ""U_Stage""='" & StageV & "'"
                'rsup3.DoQuery(Qry)

                Dim Qry As String = "SELECT COUNT(*) FROM ""@SBO_DD"" t1 INNER JOIN ""@SBO_DD1"" t2 ON t1.""DocEntry"" = t2.""DocEntry"""
                Qry = Qry & " WHERE ""U_docnum"" = '" & APPIDV & "' and t2.""U_statusN"" ='A' and t2.""U_Stage""='" & StageV & "'"
                rsup3.DoQuery(Qry)
                Dim ApprCount As Double = rsup3.Fields.Item(0).Value

                If ApprCount >= NoOfApprovals Then
                    Dim NextLineQry As String = "SELECT T3.""LineId"" + 1 FROM ""@SBO_APPHDR"" T0 " &
                    " INNER JOIN ""@SBO_APPAUT"" T3 ON T3.""Code"" = T0.""Code"" " &
                    " WHERE ""U_Active"" = 'Y' AND T0.""Code""='" & TemplateID & "' and T3.""U_M3_1""='" & StageV & "' "
                    Dim oRsNextLIne As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsNextLIne.DoQuery(NextLineQry)
                    Dim LineId As String = oRsNextLIne.Fields.Item(0).Value

                    Dim NextAuthQry As String = "SELECT T0.""Code"" As ""TemplateID"",S1.""Code"" As ""Stage"",S2.""U_UKey"" As ""Authorizers"" " &
                    " ,S1.""U_NAP"" As ""NoOfAuth"",S1.""U_NRJ"" As ""NoOfRej"",S2.""LineId"" FROM ""@SBO_APPHDR"" T0 " &
                    " INNER JOIN ""@SBO_APPAUT"" T3 ON T3.""Code"" = T0.""Code"" INNER JOIN ""@SBO_AST"" S1 ON T3.""U_M3_1""=S1.""Code"" " &
                    " INNER JOIN ""@SBO_AST_C0"" S2 ON S1.""Code""=S2.""Code"" WHERE ""U_Active"" = 'Y' AND IFNULL(S2.""U_UKey"",'')<>'' " &
                    " AND T0.""Code""='" & TemplateID & "' and T3.""LineId""='" & LineId & "';"
                    Dim oRsNextAuth As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsNextAuth.DoQuery(NextAuthQry)
                    If oRsNextAuth.RecordCount > 0 Then
                        Dim NextStage As String = oRsNextAuth.Fields.Item("Stage").Value

                        'Dim DraftQry As String = "UPDATE ""@SBO_WDD1"" SET ""U_statusN""='S',""U_Remarks""='" & Remarks & "',""U_Udate""='" & DateTime.Now.ToString("yyyyMMdd") & "',""U_Utime""='" & DateTime.Now.ToString("HHmm") & "'  " &
                        '" where ""U_Stage""='" & NextStage & "' and ""DocEntry""=(select ""DocEntry"" from ""@SBO_OWDD"" where ""U_docnum""='" & APPIDV & "')"

                        Dim DraftQry As String = "UPDATE ""@SBO_DD1"" SET ""U_statusN""='S',""U_Remarks""='" & Remarks & "',""U_Udate""='" & DateTime.Now.ToString("yyyyMMdd") & "',""U_Utime""='" & DateTime.Now.ToString("HHmm") & "'  " &
                        " where ""U_Stage""='" & NextStage & "' and ""DocEntry""=(select ""DocEntry"" from ""@SBO_DD"" where ""U_docnum""='" & APPIDV & "')"
                        Dim oRsDraft As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsDraft.DoQuery(DraftQry)
                        Me.SendMessageAlert(oRsNextAuth, APPIDV)

                    Else '-----after completing all the stages
                        'Dim DraftMasterQry As String = "UPDATE ""@SBO_OWDD"" SET ""U_statusN""='A',""U_Isdraft""='N' where ""U_docnum""='" & APPIDV & "'"
                        Dim DraftMasterQry As String = "UPDATE ""@SBO_DD"" SET ""U_statusN""='A',""U_Isdraft""='N' where ""U_docnum""='" & APPIDV & "'"
                        oRsDraftMaster.DoQuery(DraftMasterQry)
                        Dim DocUpdateQry As String = "UPDATE """ & Table & """ SET """ & AppStatField & """='A' where """ & AppIDField & """='" & APPIDV & "'"
                        oRsDocUpdate.DoQuery(DocUpdateQry)
                        Me.SendReplyMessageAlert(OriginatorV, Status, Remarks, StageV, APPIDV)


                        If DocumentType = "LvApp" Then
                            Me.LeaveUpdate(APPIDV)
                        ElseIf DocumentType = "LnApp" Then
                            Me.LoanUpdate(APPIDV)
                        ElseIf DocumentType = "Review" Then
                            Me.ReviewUpdate(APPIDV)
                        ElseIf DocumentType = "LnPr" Then
                            Me.LoanPreclosure(APPIDV)
                        ElseIf DocumentType = "CompOff" Then
                            Me.CompOffUpdate(APPIDV)
                        ElseIf DocumentType = "Trans" Then
                            Me.EmpTransferUpdate(APPIDV)
                        ElseIf DocumentType = "AReg" Then
                            Me.UpdateRegularization(APPIDV)
                        ElseIf DocumentType = "Encash" Then
                            Me.LeaveEncashmentUpdate(APPIDV)
                        ElseIf DocumentType = "LvRj" Then
                            Me.RejoinDateUpdate(APPIDV)
                        ElseIf DocumentType = "Air" Then
                            Me.UpdateEmpAirTicket(APPIDV)
                        ElseIf DocumentType = "BP" Then
                            Me.BPUpdate(APPIDV)
                        End If


                    End If

                    'If DocumentType = "AttSum" Then
                    '    If BranchId <> "" Then
                    '        oRsDocUpdate.DoQuery("UPDATE ""@SBO_PRCONTROL"" SET ""U_ATSUMSTS""='A' where ""U_BPLId""='" & BranchId & "' and Year(""U_LPDate"")='" & PYear & "' and (Month(""U_LPDate"")='" & PMonth & "' or Month(""U_NPDate"")='" & PMonth & "') ")
                    '    Else
                    '        oRsDocUpdate.DoQuery("UPDATE ""@SBO_PRCONTROL"" SET ""U_ATSUMSTS""='A' where Year(""U_LPDate"")='" & PYear & "' and (Month(""U_LPDate"")='" & PMonth & "' or Month(""U_NPDate"")='" & PMonth & "') ")
                    '    End If
                    'ElseIf DocumentType = "PrPre" Then
                    '    If BranchId <> "" Then
                    '        oRsDocUpdate.DoQuery("UPDATE ""@SBO_PRCONTROL"" SET ""U_PYPRPSTS""='A' where ""U_BPLId""='" & BranchId & "' and Year(""U_LPDate"")='" & PYear & "' and (Month(""U_LPDate"")='" & PMonth & "' or Month(""U_NPDate"")='" & PMonth & "') ")
                    '    Else
                    '        oRsDocUpdate.DoQuery("UPDATE ""@SBO_PRCONTROL"" SET ""U_PYPRPSTS""='A' where Year(""U_LPDate"")='" & PYear & "' and (Month(""U_LPDate"")='" & PMonth & "' or Month(""U_NPDate"")='" & PMonth & "') ")
                    '    End If
                    'End If

                End If

            ElseIf Status = "N" Then
                '                rsup1.DoQuery("UPDATE ""@SBO_WDD1"" SET ""U_statusN""='N',""U_Remarks""='" & Remarks & "',""U_Udate""='" & DateTime.Now.ToString("yyyyMMdd") & "',""U_Utime""='" & DateTime.Now.ToString("HHmm") & "' " &
                '"where ""U_Userid""='" & objMain.objCompany.UserName & "'  and ""U_Stage""='" & StageV & "' and ""DocEntry""=(select ""DocEntry"" from ""@SBO_OWDD"" where ""U_docnum""='" & APPIDV & "')")

                rsup1.DoQuery("UPDATE ""@SBO_DD1"" SET ""U_statusN""='N',""U_Remarks""='" & Remarks & "',""U_Udate""='" & DateTime.Now.ToString("yyyyMMdd") & "',""U_Utime""='" & DateTime.Now.ToString("HHmm") & "' " &
"where ""U_Userid""='" & objMain.objCompany.UserName & "'  and ""U_Stage""='" & StageV & "' and ""DocEntry""=(select ""DocEntry"" from ""@SBO_DD"" where ""U_docnum""='" & APPIDV & "')")


                'Dim Qry As String = "SELECT COUNT(*) FROM ""@SBO_OWDD"" t1 INNER JOIN ""@SBO_WDD1"" t2 ON t1.""DocEntry"" = t2.""DocEntry"""
                Dim Qry As String = "SELECT COUNT(*) FROM ""@SBO_DD"" t1 INNER JOIN ""@SBO_DD1"" t2 ON t1.""DocEntry"" = t2.""DocEntry"""
                Qry = Qry & " WHERE ""U_docnum"" = '" & APPIDV & "' and t2.""U_statusN"" ='N' and t2.""U_Stage""='" & StageV & "'"
                rsup3.DoQuery(Qry)
                Dim RegCount As Double = rsup3.Fields.Item(0).Value

                If RegCount >= NoOfRejects Then

                    oRsDocUpdate.DoQuery("UPDATE """ & Table & """ SET """ & AppStatField & """='N' where """ & AppIDField & """='" & APPIDV & "'")
                    'oRsDraftMaster.DoQuery("UPDATE ""@SBO_OWDD"" SET ""U_statusN"" ='N',""U_Isdraft""='N' where ""U_docnum""='" & APPIDV & "'")
                    oRsDraftMaster.DoQuery("UPDATE ""@SBO_DD"" SET ""U_statusN"" ='N',""U_Isdraft""='N' where ""U_docnum""='" & APPIDV & "'")
                    Me.SendReplyMessageAlert(OriginatorV, Status, Remarks, StageV, APPIDV)

                End If
            End If

            If DocumentType = "OL" Then
                UpdateOfferLetter(APPIDV, Status)
            End If

            'System.Runtime.InteropServices.Marshal.ReleaseComObject(rsup1)
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(rsup2)
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(rsup3)
            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try
    End Function
    Private Sub LeaveUpdate(ByVal AppIDV As String)
        '----------------------------If leave rejected------------------------------------
        'str = "SELECT T0.""U_empid"", T1.""U_Code"", T1.""U_ladays"",T2.""U_LeaBda"" , T2.""U_LeaBda"" - T1.""U_ladays"" ""BalanceLeave"" , T2.""U_PRCLAP"", T2.""U_PRCLAP"" + T1.""U_ladays"" ""AppliedLeaves"",T2.""U_LeaAda"", T2.""U_LeaAda"" +  T1.""U_ladays"" ""LeaveUsed"" FROM ""@SBO_PRLEVAPPMSTR"" T0 INNER JOIN ""@SBO_PRLEVAPPDET0"" T1 ON T0.""Code""=T1.""Code"" INNER JOIN ""@SBO_PREMPSALDET2"" T2 ON T2.""Code""= T0.""U_empid"" and T2.""U_Code"" = T1.""U_Code"" WHERE T0.""U_Appid""='" & AppIDV & "'"
        'rsGetLeaveDetails.DoQuery(str)
        'If rsGetLeaveDetails.RecordCount > 0 Then
        '    Dim EmpId, LeaveCode, BalanceLeave, AppliedLeaves, LeaveUsed As String
        '    EmpId = rsGetLeaveDetails.Fields.Item("U_empid").Value
        '    LeaveCode = rsGetLeaveDetails.Fields.Item("U_Code").Value
        '    BalanceLeave = rsGetLeaveDetails.Fields.Item("BalanceLeave").Value
        '    AppliedLeaves = rsGetLeaveDetails.Fields.Item("AppliedLeaves").Value
        '    LeaveUsed = rsGetLeaveDetails.Fields.Item("LeaveUsed").Value
        '    rsUpdateEmpLeaveQuota.DoQuery("Update  ""@SBO_PREMPSALDET2"" set  ""U_PRCLAP"" ='" & AppliedLeaves & "'  where ""Code""='" & Trim(EmpId) & "' and ""U_Code"" ='" & Trim(LeaveCode) & "'")
        'End If
        '----------------------------------------------------------------------------------

        'str = "SELECT T0.""U_empid"", T1.""U_Code"", T1.""U_ladays"",T2.""U_LeaBda"" , T2.""U_LeaBda"" - T1.""U_ladays"" ""BalanceLeave"" , T2.""U_PRCLAP"", T2.""U_PRCLAP"" + T1.""U_ladays"" ""AppliedLeaves"",T2.""U_LeaAda"", T2.""U_LeaAda"" +  T1.""U_ladays"" ""LeaveUsed"" FROM ""@SBO_PRLEVAPPMSTR"" T0 INNER JOIN ""@SBO_PRLEVAPPDET0"" T1 ON T0.""Code""=T1.""Code"" INNER JOIN ""@SBO_PREMPSALDET2"" T2 ON T2.""Code""= T0.""U_empid"" and T2.""U_Code"" = T1.""U_Code"" WHERE T0.""U_Appid""='" & AppIDV & "'"
        'rsGetLeaveDetails.DoQuery(str)
        'If rsGetLeaveDetails.RecordCount > 0 Then
        '    Dim EmpId, LeaveCode, BalanceLeave, AppliedLeaves, LeaveUsed As String
        '    EmpId = rsGetLeaveDetails.Fields.Item("U_empid").Value
        '    LeaveCode = rsGetLeaveDetails.Fields.Item("U_Code").Value
        '    BalanceLeave = rsGetLeaveDetails.Fields.Item("BalanceLeave").Value
        '    AppliedLeaves = rsGetLeaveDetails.Fields.Item("AppliedLeaves").Value
        '    LeaveUsed = rsGetLeaveDetails.Fields.Item("LeaveUsed").Value
        '    rsUpdateEmpLeaveQuota.DoQuery("Update  ""@SBO_PREMPSALDET2"" set ""U_LeaBda""='" & BalanceLeave & "' , ""U_PRCLAP"" ='" & AppliedLeaves & "' , ""U_LeaAda"" ='" & LeaveUsed & "' where ""Code""='" & Trim(EmpId) & "' and ""U_Code"" ='" & Trim(LeaveCode) & "'")
        '    '------Leave Transaction Table------------------------------------
        '    Dim LeaveAmt As Double = 0
        '    Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '    Dim Qry As String = "SELECT SUM(T0.""U_LeaveAmt"")/SUM(T0.""U_Leaves"") FROM ""@SBO_LEAVETRANS""  T0 WHERE T0.""U_Code""='" & EmpId & "'  and  T0.""U_LCode""='" & LeaveCode & "'"
        '    Try
        '        oRs.DoQuery(Qry)
        '        LeaveAmt = oRs.Fields.Item(0).Value
        '    Catch ex As Exception
        '        LeaveAmt = 0
        '    End Try


        '    objMain.sCmp = objMain.objCompany.GetCompanyService()
        '    oGeneralService = objMain.sCmp.GetGeneralService("SBO_LEAVETRANSUDO")
        '    oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
        '    oGeneralData.SetProperty("DocNum", Me.AutoIncrementCodeForTransaction())
        '    oGeneralData.SetProperty("U_Code", EmpId)
        '    oGeneralData.SetProperty("U_Name", "")
        '    oGeneralData.SetProperty("U_Date", DateTime.Now.ToString("yyyyMMdd"))
        '    oGeneralData.SetProperty("U_LCode", LeaveCode)
        '    oGeneralData.SetProperty("U_LName", "")
        '    oGeneralData.SetProperty("U_TrnsType", "LA")
        '    oGeneralData.SetProperty("U_TrnsNo", AppIDV)
        '    Dim Leaves As Double = CDbl("-" & rsGetLeaveDetails.Fields.Item("U_ladays").Value)
        '    oGeneralData.SetProperty("U_Leaves", Leaves)
        '    oGeneralData.SetProperty("U_LeaveAmt", "-" & CStr(LeaveAmt * CDbl(rsGetLeaveDetails.Fields.Item("U_ladays").Value)))
        '    oGeneralService.Add(oGeneralData)
        'End If

        '-----------------------------------------------------------------------
        str = "SELECT T0.""U_empid"", T1.""U_Code"", T1.""U_ladays"",T2.""U_LeaBda"" , T2.""U_LeaBda"" - T1.""U_ladays"" ""BalanceLeave"" , T2.""U_PRCLAP"", T2.""U_PRCLAP"" + " &
"T1.""U_ladays"" ""AppliedLeaves"",T2.""U_LeaAda"", T2.""U_LeaAda"" +  T1.""U_ladays"" ""LeaveUsed"" FROM ""@SBO_PRLEVAPPMSTR"" T0 INNER JOIN ""@SBO_PRLEVAPPDET0"" " &
"T1 ON T0.""Code""=T1.""Code"" INNER JOIN ""@SBO_PREMPSALDET2"" T2 ON T2.""Code""= T0.""U_empid"" and T2.""U_Code"" = T1.""U_Code"" WHERE T0.""U_Appid""='" & AppIDV & "'"
        rsGetLeaveDetails = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rsGetLeaveDetails.DoQuery(str)
        If rsGetLeaveDetails.RecordCount > 0 Then
            Dim EmpId, LeaveCode, BalanceLeave, AppliedLeaves, LeaveUsed As String
            EmpId = rsGetLeaveDetails.Fields.Item("U_empid").Value
            LeaveCode = rsGetLeaveDetails.Fields.Item("U_Code").Value
            BalanceLeave = rsGetLeaveDetails.Fields.Item("BalanceLeave").Value
            AppliedLeaves = rsGetLeaveDetails.Fields.Item("AppliedLeaves").Value
            LeaveUsed = rsGetLeaveDetails.Fields.Item("LeaveUsed").Value
            rsUpdateEmpLeaveQuota = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsUpdateEmpLeaveQuota.DoQuery("Update  ""@SBO_PREMPSALDET2"" set ""U_LeaBda""='" & BalanceLeave & "' , ""U_PRCLAP"" ='" & AppliedLeaves & "' , ""U_LeaAda"" ='" & LeaveUsed & "' where ""Code""='" & Trim(EmpId) & "' and ""U_Code"" ='" & Trim(LeaveCode) & "'")
            '------Leave Transaction Table------------------------------------
            Dim LeaveAmt As Double = 0
            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim Qry As String
            Try
                Qry = "SELECT SUM(T0.""U_LeaveAmt"")/SUM(T0.""U_Leaves"") FROM ""@SBO_LEAVETRANS""  T0 WHERE T0.""U_Code""='" & EmpId & "'  and  T0.""U_LCode""='" & LeaveCode & "'"
                oRs.DoQuery(Qry)
            Catch ex As Exception
                oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Qry = "Select Sum(""LeaveAmount"")/Sum(""Leaves"") From (Select ""MONTHLYLEAVEACCRUAL"" As ""Leaves"",""MONTHLYLEAVEAMOUNT"" As ""LeaveAmount"" From ""SBO_LEAVETRANS"" Where ""EMPID""='" & EmpId & "' " &
               " Union all SELECT T0.""U_LeaveAmt"",T0.""U_Leaves"" FROM ""@SBO_LEAVETRANS""  T0 WHERE T0.""U_Code""='" & EmpId & "'  and  T0.""U_LCode""='" & LeaveCode & "')"
                oRs.DoQuery(Qry)
            End Try

            LeaveAmt = oRs.Fields.Item(0).Value
            objMain.sCmp = objMain.objCompany.GetCompanyService()
            objMain.oGeneralService = objMain.sCmp.GetGeneralService("SBO_LEAVETRANSUDO")
            objMain.oGeneralData = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            objMain.oGeneralData.SetProperty("DocNum", Me.AutoIncrementCodeForTransaction())
            objMain.oGeneralData.SetProperty("U_Code", EmpId)
            objMain.oGeneralData.SetProperty("U_Name", "")
            objMain.oGeneralData.SetProperty("U_Date", DateTime.Now.ToString("yyyyMMdd"))
            objMain.oGeneralData.SetProperty("U_LCode", LeaveCode)
            objMain.oGeneralData.SetProperty("U_LName", "")
            objMain.oGeneralData.SetProperty("U_TrnsType", "LA")
            objMain.oGeneralData.SetProperty("U_TrnsNo", AppIDV)
            Dim Leaves As Double = CDbl("-" & rsGetLeaveDetails.Fields.Item("U_ladays").Value)
            objMain.oGeneralData.SetProperty("U_Leaves", Leaves)
            objMain.oGeneralData.SetProperty("U_LeaveAmt", "-" & CStr(LeaveAmt * CDbl(rsGetLeaveDetails.Fields.Item("U_ladays").Value)))
            objMain.oGeneralService.Add(objMain.oGeneralData)
        End If
    End Sub
    Private Sub LeaveEncashmentUpdate(ByVal AppIDV As String)
        '-------------Updating the details-------------------------------
        '-------------Leave Transaction Table----------------------------

        Try
            str = " Select LE.""U_Appid"", LE0.""U_EMPCODE"", LE0.""U_EMPNAME"", LE0.""U_LCode"", LE0.""U_LName"", LE0.""U_PLeaves"", LE0.""U_PAmt"" " &
            " from ""@SBO_LEAVENCASH"" LE INNER JOIN ""@SBO_LEAVENCASH_C0"" LE0 ON LE.""DocEntry"" = LE0.""DocEntry"" " &
            " Where LE.""U_Appid"" = '" & AppIDV & "' and LE0.""U_EMPCODE"" <> '' and LE0.""U_Post"" = 'Y' "
            rsGetLeaveDetails = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsGetLeaveDetails.DoQuery(str)
            If rsGetLeaveDetails.RecordCount > 0 Then
                Dim EmpId, EmployeeName, LeaveCode, LeaveName, PostingLeaves, PostingAmount As String
                EmpId = rsGetLeaveDetails.Fields.Item("U_EMPCODE").Value
                EmployeeName = rsGetLeaveDetails.Fields.Item("U_EMPNAME").Value
                LeaveCode = rsGetLeaveDetails.Fields.Item("U_LCode").Value
                LeaveName = rsGetLeaveDetails.Fields.Item("U_LName").Value
                PostingLeaves = rsGetLeaveDetails.Fields.Item("U_PLeaves").Value
                PostingAmount = rsGetLeaveDetails.Fields.Item("U_PAmt").Value
                objMain.sCmp = objMain.objCompany.GetCompanyService()
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("SBO_LEAVETRANSUDO")
                objMain.oGeneralData = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                objMain.oGeneralData.SetProperty("DocNum", Me.AutoIncrementCodeForTransaction())
                objMain.oGeneralData.SetProperty("U_Code", EmpId)
                objMain.oGeneralData.SetProperty("U_Name", EmployeeName)
                objMain.oGeneralData.SetProperty("U_Date", DateTime.Now.ToString("yyyyMMdd"))
                objMain.oGeneralData.SetProperty("U_LCode", LeaveCode)
                objMain.oGeneralData.SetProperty("U_LName", LeaveName)
                objMain.oGeneralData.SetProperty("U_TrnsType", "LE")
                objMain.oGeneralData.SetProperty("U_TrnsNo", AppIDV)
                Dim Leaves As Double = CDbl("-" & PostingLeaves)
                objMain.oGeneralData.SetProperty("U_Leaves", Leaves)
                objMain.oGeneralData.SetProperty("U_LeaveAmt", "-" & PostingAmount)
                Try
                    objMain.oGeneralService.Add(objMain.oGeneralData)
                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText("Error" & ex.Message)
                End Try


                '--------Update Employee salary setup-------------------------------
                Dim rs1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'Dim Query1 As String = "SELECT T1.""Code"",H1.""firstName"",T1.""LineId"",T1.""U_LeaBda"",L1.""U_Remarks"" " &
                '" FROM ""@SBO_PREMPSALMSTR"" T0 INNER JOIN ""@SBO_PREMPSALDET2"" T1 ON T0.""Code""=T1.""Code"" INNER JOIN ""@SBO_PRLEAVECODE"" L1 " &
                '" ON T1.""U_Code""=L1.""Code"" INNER JOIN ""OHEM"" H1 ON T0.""Code""=H1.""empID"" " &
                '" WHERE T1.""Code""='" & EmpId & "' and L1.""Code""='" & LeaveCode & "'"

                Dim Query1 As String = " Select T1.""Code"", T1.""LineId"", T1.""U_LeaBda"" - LE0.""U_PLeaves"" as ""BalanceLeave"", " &
                " T1.""U_PRCLAP"" + LE0.""U_PLeaves"" as ""AppliedLeaves"", T1.""U_LeaAda"" +  LE0.""U_PLeaves"" as ""LeaveUsed"" " &
                " from ""@SBO_LEAVENCASH"" LE INNER JOIN ""@SBO_LEAVENCASH_C0"" LE0 ON LE.""DocEntry"" = LE0.""DocEntry"" " &
                " INNER Join ""@SBO_PREMPSALMSTR"" T0 ON LE0.""U_EMPCODE"" = T0.""Code"" " &
                " INNER Join ""@SBO_PREMPSALDET2"" T1 ON T0.""Code""=T1.""Code"" And LE0.""U_LCode"" = T1.""U_Code"" " &
                " Where LE.""U_Appid"" = '" & AppIDV & "' and LE0.""U_EMPCODE"" <> '' and LE0.""U_Post"" = 'Y' "
                rs1.DoQuery(Query1)
                If rs1.RecordCount > 0 Then
                    objMain.oGeneralService = objMain.sCmp.GetGeneralService("SBO_EMPSALUDO")
                    objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                    objMain.oGeneralParams.SetProperty("Code", rs1.Fields.Item("Code").Value)
                    objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                    objMain.oChildren = objMain.oGeneralData.Child("SBO_PREMPSALDET2")
                    objMain.oChild = objMain.oChildren.Item(rs1.Fields.Item("LineId").Value - 1)
                    objMain.oChild.SetProperty("U_LeaBda", rs1.Fields.Item("BalanceLeave").Value)
                    objMain.oChild.SetProperty("U_LeaAda", rs1.Fields.Item("LeaveUsed").Value)
                    objMain.oGeneralService.Update(objMain.oGeneralData)
                End If

                'Dim rs2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'Dim Query2 As String = "SELECT SUM(IFNULL(""U_Leaves"",0)) FROM ""@SBO_LEAVETRANS"" Where ""U_Code""='" & rs1.Fields.Item("Code").Value & "' and ""U_LCode""='" & LeaveCode & "'"
                'rs2.DoQuery(Query2)
                'oChild.SetProperty("U_LeaBda", rs2.Fields.Item(0).Value)
                'oGeneralService.Update(oGeneralData)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error while updating leave transaction table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
        '-------------------------------------------------------------------
    End Sub
    Sub RejoinDateUpdate(ByVal AppIDV As String)
        Try
            'RejoinAppID = Me.m_SBO_Form.Items.Item("LAppno").Specific.Value.ToString().Trim()
            If AppIDV <> "" Then
                'RejoinAppID = LeaveRejoinAppId

                Dim oRsRejoinDate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim UpdateRejoinQry As String = "Update OHEM Set ""U_RJDate""=(Select ""U_RDuty"" From ""@SBO_NOTIF"" Where ""U_Appid""='" & AppIDV & "') Where ""empID""=(Select ""U_EmpId"" From ""@SBO_NOTIF"" Where ""U_Appid""='" & AppIDV & "') "
                oRsRejoinDate.DoQuery(UpdateRejoinQry)

                Dim StatusQry As String = "SELECT T0.""statusID"" FROM OHST T0 WHERE T0.""U_Status"" ='Active'"
                Dim oRsStatus As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsStatus.DoQuery(StatusQry)

                Dim EmpIDQry As String = "Select ""U_EmpId"" From ""@SBO_NOTIF"" Where ""U_Appid""='" & AppIDV & "'"
                Dim oRsEmpID As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsEmpID.DoQuery(EmpIDQry)

                Dim oEmployee As SAPbobsCOM.EmployeesInfo
                oEmployee = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo)
                If (oEmployee.GetByKey(oRsEmpID.Fields.Item(0).Value)) = True Then
                    oEmployee.StatusCode = oRsStatus.Fields.Item(0).Value
                    oEmployee.Update()
                End If

                '-------Checking rejoin date----------------
                Dim Qry As String = " SELECT DAYS_BETWEEN(""U_RDuty"",""U_LTo"") + 1 As ""EarlyDays"",""U_LAppno"" AS ""LeaveAppID"" " &
                " FROM ""@SBO_NOTIF"" WHERE  ""U_Appid""='" & AppIDV & "' AND ""U_RDuty"" BETWEEN ""U_LFrom"" AND ""U_LTo"""
                Dim oRsEarlyDays As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsEarlyDays.DoQuery(Qry)
                If oRsEarlyDays.RecordCount > 0 Then
                    Me.LeaveRejoinUpdate(oRsEarlyDays.Fields.Item(0).Value, oRsEarlyDays.Fields.Item(1).Value, AppIDV)
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error while update:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
    Private Sub LeaveRejoinUpdate(ByVal EarlyDays As Double, ByVal LeaveAppID As String, ByVal RejoinAppID As String)
        str = "SELECT T0.""U_empid"", T1.""U_Code"", T1.""U_ladays"",T2.""U_LeaBda"" , T2.""U_LeaBda"" AS ""BalanceLeave"", T2.""U_LeaAda"" ""LeaveUsed"" FROM ""@SBO_PRLEVAPPMSTR"" T0 INNER JOIN ""@SBO_PRLEVAPPDET0"" T1 ON T0.""Code""=T1.""Code"" INNER JOIN ""@SBO_PREMPSALDET2"" T2 ON T2.""Code""= T0.""U_empid"" and T2.""U_Code"" = T1.""U_Code"" WHERE T0.""U_Appid""='" & LeaveAppID & "'"
        Dim rsGetLeaveDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rsGetLeaveDetails.DoQuery(str)
        If rsGetLeaveDetails.RecordCount > 0 Then
            Dim EmpId, LeaveCode, BalanceLeave, AppliedLeaves, LeaveUsed As String
            EmpId = rsGetLeaveDetails.Fields.Item("U_empid").Value
            LeaveCode = rsGetLeaveDetails.Fields.Item("U_Code").Value
            BalanceLeave = rsGetLeaveDetails.Fields.Item("BalanceLeave").Value + EarlyDays
            'AppliedLeaves = rsGetLeaveDetails.Fields.Item("AppliedLeaves").Value
            LeaveUsed = rsGetLeaveDetails.Fields.Item("LeaveUsed").Value - EarlyDays
            '------------Update Salary setup----------------------------------
            Dim rsUpdateEmpLeaveQuota As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'rsUpdateEmpLeaveQuota.DoQuery("Update  ""@SBO_PREMPSALDET2"" set ""U_LeaBda""='" & BalanceLeave & "' , ""U_PRCLAP"" ='" & AppliedLeaves & "' , ""U_LeaAda"" ='" & LeaveUsed & "' where ""Code""='" & Trim(EmpId) & "' and ""U_Code"" ='" & Trim(LeaveCode) & "'")
            Dim UpdateQry As String = "Update  ""@SBO_PREMPSALDET2"" set ""U_LeaBda""='" & BalanceLeave & "',""U_LeaAda"" ='" & LeaveUsed & "' where ""Code""='" & Trim(EmpId) & "' and ""U_Code"" ='" & Trim(LeaveCode) & "'"
            rsUpdateEmpLeaveQuota.DoQuery(UpdateQry)

            '------Leave Transaction Table------------------------------------
            Dim LeaveAmt As Double = 0
            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim Qry As String
            Try
                Qry = "SELECT SUM(T0.""U_LeaveAmt"")/SUM(T0.""U_Leaves"") FROM ""@SBO_LEAVETRANS""  T0 WHERE T0.""U_Code""='" & EmpId & "'  and  T0.""U_LCode""='" & LeaveCode & "'"
                oRs.DoQuery(Qry)
            Catch ex As Exception
                oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Qry = "Select Sum(""LeaveAmount"")/Sum(""Leaves"") From (Select ""MONTHLYLEAVEACCRUAL"" As ""Leaves"",""MONTHLYLEAVEAMOUNT"" As ""LeaveAmount"" From ""SBO_LEAVETRANS"" Where ""EMPID""='" & EmpId & "' " &
               " Union all SELECT T0.""U_LeaveAmt"",T0.""U_Leaves"" FROM ""@SBO_LEAVETRANS""  T0 WHERE T0.""U_Code""='" & EmpId & "'  and  T0.""U_LCode""='" & LeaveCode & "')"
                oRs.DoQuery(Qry)
            End Try
            LeaveAmt = oRs.Fields.Item(0).Value
            objMain.sCmp = objMain.objCompany.GetCompanyService()
            objMain.oGeneralService = objMain.sCmp.GetGeneralService("SBO_LEAVETRANSUDO")
            objMain.oGeneralData = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            objMain.oGeneralData.SetProperty("DocNum", Me.AutoIncrementCodeForTransaction())
            objMain.oGeneralData.SetProperty("U_Code", EmpId)
            objMain.oGeneralData.SetProperty("U_Name", "")
            objMain.oGeneralData.SetProperty("U_Date", DateTime.Now.ToString("yyyyMMdd"))
            objMain.oGeneralData.SetProperty("U_LCode", LeaveCode)
            objMain.oGeneralData.SetProperty("U_LName", "")
            objMain.oGeneralData.SetProperty("U_TrnsType", "REJ")
            objMain.oGeneralData.SetProperty("U_TrnsNo", RejoinAppID)
            objMain.oGeneralData.SetProperty("U_Leaves", EarlyDays)
            'oGeneralData.SetProperty("U_LeaveAmt", "-" & CStr(LeaveAmt * CDbl(EarlyDays)))
            objMain.oGeneralData.SetProperty("U_LeaveAmt", CStr(LeaveAmt * CDbl(EarlyDays)))
            objMain.oGeneralService.Add(objMain.oGeneralData)
        End If
    End Sub
    Function AutoIncrementCodeForTransaction()
        Dim rs4 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs4.DoQuery("Select Max(Ifnull(""DocNum"",0))+1 From ""@SBO_LEAVETRANS"" ")
        Return rs4.Fields.Item(0).Value
    End Function

    'Private Sub LoanUpdate()
    '    'Dim LCode As String = objMatrix.Columns.Item("LCode").Cells.Item(i).Specific.Value.ToString().Trim()
    '    Dim rs5 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    objMain.sCmp = objMain.objCompany.GetCompanyService()

    '    Dim LQuery As String = " Select T0.""U_empid"",T0.""U_Appid"",T1.""U_Code"",T1.""U_dscr"",T1.""U_TAmt"",T1.""U_Inst"",T1.""U_IAmt"",T1.""U_Reason"",T1.""U_ApprAmt"" " & _
    '      " from ""@SBO_LOAN"" T0 INNER JOIN ""@SBO_LOAN_C0"" T1 ON T0.""DocEntry""=T1.""DocEntry"" " & _
    '      " Where T0.""U_Appid"" = '" & AppIDV & "' "
    '    rs5.DoQuery(LQuery)
    '    If rs5.RecordCount > 0 Then
    '        oGeneralService = objMain.sCmp.GetGeneralService("SBO_EMPSALUDO")
    '        oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
    '        oGeneralParams.SetProperty("Code", rs5.Fields.Item("U_empid").Value)
    '        oGeneralData = oGeneralService.GetByParams(oGeneralParams)
    '        oChildren = oGeneralData.Child("SBO_PREMPSALDET4")

    '        oChild = oChildren.Add()
    '        'oChild = oChildren.Item(rs5.Fields.Item("LineId").Value - 1)
    '        oChild.SetProperty("U_LnCD", rs5.Fields.Item("U_Code").Value)
    '        oChild.SetProperty("U_Desc", rs5.Fields.Item("U_dscr").Value)
    '        oChild.SetProperty("U_Lamt", rs5.Fields.Item("U_ApprAmt").Value)
    '        oChild.SetProperty("U_insMnth", rs5.Fields.Item("U_Inst").Value)
    '        oChild.SetProperty("U_insAmt", rs5.Fields.Item("U_IAmt").Value)
    '        oChild.SetProperty("U_PAmt", 0)
    '        oChild.SetProperty("U_RemInst", 0)
    '        oChild.SetProperty("U_BAmt", rs5.Fields.Item("U_ApprAmt").Value)
    '        oChild.SetProperty("U_Comments", rs5.Fields.Item("U_Reason").Value)
    '        oChild.SetProperty("U_NIPay", 1)
    '        oChild.SetProperty("U_LStat", "P")
    '        oChild.SetProperty("U_PFlag", "Y")
    '        'oChildren.Add()
    '        oGeneralService.Update(oGeneralData)
    '    End If
    'End Sub
    Private Sub LoanUpdate(ByVal AppIDV As String)
        'Dim LCode As String = objMatrix.Columns.Item("LCode").Cells.Item(i).Specific.Value.ToString().Trim()
        Dim rs5 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMain.sCmp = objMain.objCompany.GetCompanyService()

        Dim LQuery As String = " Select T0.""U_empid"",T0.""U_Appid"",T1.""U_Code"",T1.""U_dscr"",T1.""U_TAmt"",T1.""U_Inst"",T1.""U_IAmt"",T1.""U_Reason"",T1.""U_ApprAmt"" " &
          " from ""@SBO_LOAN"" T0 INNER JOIN ""@SBO_LOAN_C0"" T1 ON T0.""DocEntry""=T1.""DocEntry"" " &
          " Where T0.""U_Appid"" = '" & AppIDV & "' "
        rs5.DoQuery(LQuery)
        If rs5.RecordCount > 0 Then
            Dim oRsLineCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim LineCountQry As String = "SELECT *  FROM ""@SBO_PREMPSALDET4""  T0 WHERE T0.""Code"" ='" & rs5.Fields.Item("U_empid").Value & "' and IFNULL(""U_LnCD"",'')='' and ""LineId""=1"
            oRsLineCount.DoQuery(LineCountQry)
            If oRsLineCount.RecordCount = 0 Then

                objMain.oGeneralService = objMain.sCmp.GetGeneralService("SBO_EMPSALUDO")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("Code", rs5.Fields.Item("U_empid").Value)
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                objMain.oChildren = objMain.oGeneralData.Child("SBO_PREMPSALDET4")

                objMain.oChild = objMain.oChildren.Add()
                'oChild = oChildren.Item(rs5.Fields.Item("LineId").Value - 1)
                objMain.oChild.SetProperty("U_LnCD", rs5.Fields.Item("U_Code").Value)
                objMain.oChild.SetProperty("U_Desc", rs5.Fields.Item("U_dscr").Value)
                objMain.oChild.SetProperty("U_Lamt", rs5.Fields.Item("U_ApprAmt").Value)
                objMain.oChild.SetProperty("U_insMnth", rs5.Fields.Item("U_Inst").Value)
                objMain.oChild.SetProperty("U_insAmt", rs5.Fields.Item("U_IAmt").Value)
                objMain.oChild.SetProperty("U_PAmt", 0)
                objMain.oChild.SetProperty("U_RemInst", 0)
                objMain.oChild.SetProperty("U_BAmt", rs5.Fields.Item("U_ApprAmt").Value)
                objMain.oChild.SetProperty("U_Comments", rs5.Fields.Item("U_Reason").Value)
                objMain.oChild.SetProperty("U_NIPay", 1)
                objMain.oChild.SetProperty("U_LStat", "P")
                objMain.oChild.SetProperty("U_PFlag", "Y")
                'oChildren.Add()
                objMain.oGeneralService.Update(objMain.oGeneralData)

            Else

                objMain.oGeneralService = objMain.sCmp.GetGeneralService("SBO_EMPSALUDO")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("Code", rs5.Fields.Item("U_empid").Value)
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                objMain.oChildren = objMain.oGeneralData.Child("SBO_PREMPSALDET4")

                objMain.oChild = objMain.oChildren.Item(oRsLineCount.Fields.Item("LineId").Value - 1)
                objMain.oChild.SetProperty("U_LnCD", rs5.Fields.Item("U_Code").Value)
                objMain.oChild.SetProperty("U_Desc", rs5.Fields.Item("U_dscr").Value)
                objMain.oChild.SetProperty("U_Lamt", rs5.Fields.Item("U_ApprAmt").Value)
                objMain.oChild.SetProperty("U_insMnth", rs5.Fields.Item("U_Inst").Value)
                objMain.oChild.SetProperty("U_insAmt", rs5.Fields.Item("U_IAmt").Value)
                objMain.oChild.SetProperty("U_PAmt", 0)
                objMain.oChild.SetProperty("U_RemInst", 0)
                objMain.oChild.SetProperty("U_BAmt", rs5.Fields.Item("U_ApprAmt").Value)
                objMain.oChild.SetProperty("U_Comments", rs5.Fields.Item("U_Reason").Value)
                objMain.oChild.SetProperty("U_NIPay", 1)
                objMain.oChild.SetProperty("U_LStat", "P")
                objMain.oChild.SetProperty("U_PFlag", "Y")
                'oChildren.Add()
                objMain.oGeneralService.Update(objMain.oGeneralData)


            End If
        End If
    End Sub
    Private Sub CompOffUpdate(ByVal AppIDV As String)
        str = "SELECT T0.""U_empid"", T1.""U_Code"", T1.""U_ladays"",T2.""U_LeaBda"" , T2.""U_LeaBda"" + T1.""U_ladays"" ""BalanceLeave"" " &
            " FROM ""@SBO_COMPOFF"" T0 " &
            " INNER JOIN ""@SBO_COMPOFF_C0"" T1 ON T0.""DocEntry""=T1.""DocEntry"" INNER JOIN ""@SBO_PREMPSALDET2"" T2 ON T2.""Code""= T0.""U_empid"" and T2.""U_Code"" = T1.""U_Code"" WHERE T0.""U_Appid""='" & AppIDV & "'"
        rsGetLeaveDetails = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rsGetLeaveDetails.DoQuery(str)
        If rsGetLeaveDetails.RecordCount > 0 Then
            Dim EmpId, LeaveCode, BalanceLeave, AppliedLeaves, LeaveUsed As String
            EmpId = rsGetLeaveDetails.Fields.Item("U_empid").Value
            LeaveCode = rsGetLeaveDetails.Fields.Item("U_Code").Value
            BalanceLeave = rsGetLeaveDetails.Fields.Item("BalanceLeave").Value
            rsUpdateEmpLeaveQuota = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsUpdateEmpLeaveQuota.DoQuery("Update  ""@SBO_PREMPSALDET2"" set ""U_LeaBda""='" & BalanceLeave & "' where ""Code""='" & Trim(EmpId) & "' and ""U_Code"" ='" & Trim(LeaveCode) & "'")
            '------Leave Transaction Table------------------------------------
            objMain.sCmp = objMain.objCompany.GetCompanyService()
            objMain.oGeneralService = objMain.sCmp.GetGeneralService("SBO_LEAVETRANSUDO")
            objMain.oGeneralData = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            objMain.oGeneralData.SetProperty("DocNum", Me.AutoIncrementCodeForTransaction())
            objMain.oGeneralData.SetProperty("U_Code", EmpId)
            objMain.oGeneralData.SetProperty("U_Name", "")
            objMain.oGeneralData.SetProperty("U_Date", DateTime.Now.ToString("yyyyMMdd"))
            objMain.oGeneralData.SetProperty("U_LCode", LeaveCode)
            objMain.oGeneralData.SetProperty("U_LName", "")
            objMain.oGeneralData.SetProperty("U_TrnsType", "COMP")
            objMain.oGeneralData.SetProperty("U_TrnsNo", AppIDV)
            objMain.oGeneralData.SetProperty("U_Leaves", rsGetLeaveDetails.Fields.Item("U_ladays").Value)
            objMain.oGeneralData.SetProperty("U_LeaveAmt", 0)
            objMain.oGeneralService.Add(objMain.oGeneralData)
        End If
    End Sub
    Sub LoanPreclosure(ByVal AppIDV As String)
        str = "Select *,(Select T.""U_LnType"" From ""@SBO_PRLOAN"" T Where T.""Code""=""U_RefNo"") As ""LoanType"" From ""@SBO_LOANPRECLOSR"" WHERE ""U_Appid""='" & AppIDV & "'"
        rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs.DoQuery(str)
        If rs.RecordCount > 0 Then
            Dim EmpId, LineId As String
            EmpId = rs.Fields.Item("U_EmpId").Value
            LineId = rs.Fields.Item("U_LoanNo").Value

            Dim loanType As String = rs.Fields.Item("LoanType").Value
            Dim oRsUpdt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim LoanAmt As Double = CDbl(rs.Fields.Item("U_LAdvAmt").Value)
            Dim BalAmt As Double = CDbl(rs.Fields.Item("U_BalAmt").Value)
            Dim PaidAmt As Double = CDbl(rs.Fields.Item("U_PaidAmt").Value)
            Dim PaidInst As Double = CDbl(rs.Fields.Item("U_PaidInst").Value)
            Dim TotalInst As Double = CDbl(rs.Fields.Item("U_PaidInst").Value) + CDbl(rs.Fields.Item("U_BalInst").Value)
            Dim ClsAmt As Double = CDbl(rs.Fields.Item("U_Amount").Value)
            Dim InstAmt As Double = CDbl(rs.Fields.Item("U_InstAmt").Value)
            Dim ClsInst As Double = CDbl(rs.Fields.Item("U_ClsInst").Value)

            Dim PaidAmount As Double = PaidAmt + ClsAmt
            Dim BalanceAmount As Double = LoanAmt - PaidAmount
            Dim PaidInstallments As Double = ClsInst + PaidInst
            Dim UpdateQry As String
            If loanType = "1" Then
                If BalanceAmount <> 0 Then
                    UpdateQry = "Update  ""@SBO_PREMPSALDET4"" set ""U_PAmt""='" & PaidAmount & "',""U_BAmt""='" & BalanceAmount & "',""U_RemInst""='" & PaidInstallments & "' where ""Code""='" & Trim(EmpId) & "' and ""LineId"" ='" & Trim(LineId) & "'"
                Else
                    UpdateQry = "Update  ""@SBO_PREMPSALDET4"" set ""U_PAmt""='" & PaidAmount & "',""U_BAmt""='" & BalanceAmount & "',""U_RemInst""='" & PaidInstallments & "',""U_LStat""='Y',""U_PFlag""='N' where ""Code""='" & Trim(EmpId) & "' and ""LineId"" ='" & Trim(LineId) & "'"
                End If
            Else
                If BalanceAmount <> 0 Then
                    UpdateQry = "Update  ""@SBO_PREMPSALDET4"" set ""U_PAmt""='" & PaidAmount & "',""U_BAmt""='" & BalanceAmount & "' where ""Code""='" & Trim(EmpId) & "' and ""LineId"" ='" & Trim(LineId) & "'"
                Else
                    UpdateQry = "Update  ""@SBO_PREMPSALDET4"" set ""U_PAmt""='" & PaidAmount & "',""U_BAmt""='" & BalanceAmount & "',""U_LStat""='Y',""U_PFlag""='N' where ""Code""='" & Trim(EmpId) & "' and ""LineId"" ='" & Trim(LineId) & "'"
                End If
            End If
            oRsUpdt.DoQuery(UpdateQry)
        End If
    End Sub
    Sub UpdateOfferLetter(ByVal OffAppID As String, ByVal OffStatus As String)
        Try
            If OffAppID <> "" And OffStatus = "A" Then

                Dim strStatusUpdate As String = " Update ""@SBO_OFFLETTER"" set ""U_STATUS"" = 'Offer Issued' where ""U_OLNo"" = '" & OffAppID & "' "
                Dim rsStatus As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rsStatus.DoQuery(strStatusUpdate)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rsStatus)

                Dim strIntervSch As String = "Update ""@SBO_INTWSCHLR"" set ""U_IntwSts"" ='C' " &
                " where ""U_CNDID"" = (Select ""U_CNDNO"" from ""@SBO_OFFLETTER"" where ""U_OLNo"" = '" & OffAppID & "') "
                Dim rsIntw As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rsIntw.DoQuery(strIntervSch)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rsIntw)

                Dim strRec As String = " Select ""U_NoPos"", ""U_RcNo"" from ""@SBO_RCRTMNTREQST"" " &
                " where ""U_RcNo"" = (Select ""U_RECNO"" from ""@SBO_OFFLETTER"" where ""U_OLNo""='" & OffAppID & "') "
                Dim rsRec As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rsRec.DoQuery(strRec)
                Dim RecPositions As Integer = 0
                Dim OffPositions As Integer = 0
                If rsRec.RecordCount > 0 Then
                    RecPositions = rsRec.Fields.Item("U_NoPos").Value
                    Dim RecruitmentNo As String = rsRec.Fields.Item("U_RcNo").Value
                    If RecruitmentNo <> "" Then
                        Dim strOff As String = " Select count(""U_RECNO"") from ""@SBO_OFFLETTER""  " &
                        " where ""U_RECNO"" = '" & RecruitmentNo & "' and ""U_AppStat"" = 'A' "
                        Dim rsOff As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rsOff.DoQuery(strOff)
                        If rsOff.RecordCount > 0 Then
                            OffPositions = rsOff.Fields.Item(0).Value
                        End If
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(rsOff)
                    End If
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rsRec)

                If RecPositions = OffPositions Then
                    Dim rsRecruitment As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim UpdateRecruitmentStatus As String = "UPDATE ""@SBO_RCRTMNTREQST"" Set ""U_Status"" = 'C' " &
                    " where ""U_RcNo"" = (Select ""U_RECNO"" from ""@SBO_OFFLETTER"" where ""U_OLNo""='" & OffAppID & "') "
                    rsRecruitment.DoQuery(UpdateRecruitmentStatus)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsRecruitment)
                End If

            ElseIf OffAppID <> "" And OffStatus = "R" Then
                Dim strStatusUpdate As String = " Update ""@SBO_OFFLETTER"" set ""U_STATUS"" = 'Offer Rejected' where ""U_OLNo"" = '" & OffAppID & "' "
                Dim rsStatus As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rsStatus.DoQuery(strStatusUpdate)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rsStatus)
            End If
        Catch ex As Exception

        End Try

    End Sub
    Private Sub EmpTransferUpdate(ByVal AppIDV As String)
        Try

            Dim strTransfer As String = " Select distinct A.""U_empID"", A.""U_CurLocId"", A.""U_TrLocId"" " &
            " from ""@SBO_TRANSFER_C0"" A INNER JOIN  ""@SBO_TRANSFER"" B ON A. ""DocEntry"" = B.""DocEntry"" Where B.""U_AppId"" = '" & AppIDV & "' "
            Dim rsTransferDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsTransferDtls.DoQuery(strTransfer)
            If rsTransferDtls.RecordCount > 0 Then
                For i As Integer = 1 To rsTransferDtls.RecordCount
                    'rsTransferDtls.MoveFirst()
                    Dim CurrentEmpId As String = rsTransferDtls.Fields.Item("U_empID").Value
                    Dim CurrentLocation As String = rsTransferDtls.Fields.Item("U_CurLocId").Value
                    Dim TargetLocation As String = rsTransferDtls.Fields.Item("U_TrLocId").Value

                    Dim strBranchUpdate As String = "Update ""HEM10"" Set ""BPLId""='" & TargetLocation & "' Where ""empID""='" & CurrentEmpId & "' "
                    Dim oRsBranchUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsBranchUpdate.DoQuery(strBranchUpdate)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRsBranchUpdate)

                    Dim strEmpUpdate As String = "Update ""OHEM"" Set ""BPLId""='" & TargetLocation & "' Where ""empID""='" & CurrentEmpId & "' "
                    Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsUpdate.DoQuery(strEmpUpdate)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRsUpdate)

                    rsTransferDtls.MoveNext()
                Next

            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rsTransferDtls)
        Catch ex As Exception

        End Try
    End Sub
    Sub UpdateRegularization(ByVal AppIDV As String)
        Try
            Dim AttDocEntry As String = ""
            Dim AttDocStatus As String = ""
            Dim selectQuery As String = ""
            Dim rsSelect As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            selectQuery = "Select ""DocEntry"", ""U_AppStat"" from ""@SBO_ATTREGLSN"" Where ""U_DocNo"" = '" & AppIDV & "' "
            rsSelect.DoQuery(selectQuery)
            If rsSelect.RecordCount > 0 Then
                AttDocEntry = rsSelect.Fields.Item("DocEntry").Value
                AttDocStatus = "A"
                'AttDocStatus = rsSelect.Fields.Item("U_AppStat").Value
            End If

            If AttDocEntry <> "" And AttDocStatus <> "" Then
                Dim rsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim UpdateQuery As String = "UPDATE ""@SBO_ATTREGLSN_C0"" Set ""U_AppStat"" = '" & AttDocStatus & "' " &
                                    " Where ""DocEntry"" = '" & AttDocEntry & "' "
                rsUpdate.DoQuery(UpdateQuery)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rsUpdate)
                objMain.objApplication.ActivateMenuItem("1304")
            End If

        Catch ex As Exception

        End Try

    End Sub
    Private Sub UpdateEmpAirTicket(ByVal AppIDV As String)
        Try

            Dim strTransfer As String = "  Select ""U_EmpId"", ""Code"", ""U_ETOO"" from ""@SBO_AIRREQ"" A INNER JOIN ""@SBO_AIRREQ_C0"" A0 " &
            " ON A.""DocEntry"" = A0.""DocEntry"" Where A.""U_Appid"" = '" & AppIDV & "' and ""U_ETOO"" is not null and ""U_ACT"" = 'Y' "
            Dim rsTransferDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsTransferDtls.DoQuery(strTransfer)
            If rsTransferDtls.RecordCount > 0 Then
                For i As Integer = 1 To rsTransferDtls.RecordCount
                    Dim CurrentEmpId As String = rsTransferDtls.Fields.Item("U_EmpId").Value
                    Dim BCode As String = rsTransferDtls.Fields.Item("Code").Value

                    Dim strBranchUpdate As String = " Update ""@SBO_EMPDETAILS_C2"" Set ""U_Lrmtd"" = (Select ""U_ETOO"" from ""@SBO_AIRREQ"" A INNER JOIN ""@SBO_AIRREQ_C0"" A0 " &
                    " ON A.""DocEntry"" = A0.""DocEntry"" Where A.""U_Appid"" = '" & AppIDV & "' and ""Code"" = '" & BCode & "' and ""U_ETOO"" is not null and ""U_ACT"" = 'Y')  " &
                    " Where ""Code"" = '" & CurrentEmpId & "' and (""Code"" = '" & BCode & "' or ""U_Applcble"" is null) "
                    Dim oRsBranchUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsBranchUpdate.DoQuery(strBranchUpdate)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRsBranchUpdate)

                    rsTransferDtls.MoveNext()
                Next

            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rsTransferDtls)
        Catch ex As Exception

        End Try
    End Sub

    Sub BPUpdate(ByVal CardCode As String)
        Try
            Dim oBP As SAPbobsCOM.BusinessPartners = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            If oBP.GetByKey(CardCode) = True Then

                'Dim DateQry As String = "Select CURRENT_DATE As ""ValidFrom"", Add_Days(CURRENT_DATE,2950) As ""ValidTo"" From DUMMY "
                'Dim rsDate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'rsDate.DoQuery(DateQry)

                oBP.Valid = SAPbobsCOM.BoYesNoEnum.tYES
                oBP.ValidFrom = DateTime.Now()
                oBP.ValidTo = DateTime.Now().AddYears(10)

                oBP.Frozen = SAPbobsCOM.BoYesNoEnum.tYES
                oBP.FrozenFrom = DateTime.Now().AddDays(-2)
                oBP.FrozenTo = DateTime.Now().AddDays(-1)

                If oBP.Update() = 0 Then
                    objMain.objApplication.StatusBar.SetText("Business Partner Master Data Updated", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objMain.objApplication.StatusBar.SetText("Error while Updating Business Partner Master Data" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    '--- Review BackUp----
    'Sub ReviewUpdate()

    '    '----------Employee review Earnings--------------------------
    '    str = "SELECT * FROM ""@SBO_EMPREV"" T0  INNER JOIN ""@SBO_EMPREV_C0"" T1 ON T0.""DocEntry""=T1.""DocEntry"" WHERE T0.""U_Appid"" ='" & AppIDV & "'"
    '    rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    rs.DoQuery(str)
    '    If rs.RecordCount > 0 Then
    '        Dim EmpId, LineId As String
    '        EmpId = rs.Fields.Item("U_EmpId").Value
    '        '------Earnings---------------------------------------
    '        Dim oRsDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        Dim DetailsQry As String = "Select ""Code"",""U_Code"",""U_Amount"" From ""@SBO_PREMPSALDET0"" Where ""Code""='" & EmpId & "'"
    '        oRsDetails.DoQuery(DetailsQry)
    '        Dim NetPay As Double = 0
    '        For i As Integer = 1 To oRsDetails.RecordCount
    '            rs.MoveFirst()
    '            For j As Integer = 1 To rs.RecordCount
    '                If oRsDetails.Fields.Item("U_Code").Value = rs.Fields.Item("U_ECode").Value Then
    '                    Dim Update As String = "Update ""@SBO_PREMPSALDET0"" Set ""U_Amount""='" & rs.Fields.Item("U_NewSal").Value & "' Where ""Code""='" & EmpId & "' and ""U_Code""='" & oRsDetails.Fields.Item("U_Code").Value & "'"
    '                    Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                    oRsUpdate.DoQuery(Update)
    '                    NetPay = NetPay + rs.Fields.Item("U_NewSal").Value
    '                End If
    '                rs.MoveNext()
    '            Next
    '            oRsDetails.MoveNext()
    '        Next
    '        '-----Deductions-------------------------------------
    '        Dim DetailsDeductionQry As String = "Select ""Code"",""U_Code"",""U_Amount"" From ""@SBO_PREMPSALDET1"" Where ""Code""='" & EmpId & "'"
    '        oRsDetails.DoQuery(DetailsDeductionQry)
    '        Dim NetPay1 As Double = 0
    '        For i As Integer = 1 To oRsDetails.RecordCount
    '            rs.MoveFirst()
    '            For j As Integer = 1 To rs.RecordCount
    '                If oRsDetails.Fields.Item("U_Code").Value = rs.Fields.Item("U_DCode").Value Then
    '                    Dim Update As String = "Update ""@SBO_PREMPSALDET1"" Set ""U_Amount""='" & rs.Fields.Item("U_NewSal").Value & "' Where ""Code""='" & EmpId & "' and ""U_Code""='" & oRsDetails.Fields.Item("U_Code").Value & "'"
    '                    Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                    oRsUpdate.DoQuery(Update)
    '                    NetPay1 = NetPay1 + rs.Fields.Item("U_NewSal").Value
    '                End If
    '                rs.MoveNext()
    '            Next
    '            oRsDetails.MoveNext()
    '        Next
    '        '------------------------------------------
    '        objMain.sCmp = objMain.objCompany.GetCompanyService()
    '        oGeneralService = objMain.sCmp.GetGeneralService("SBO_EMPSALUDO")
    '        oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
    '        oGeneralParams.SetProperty("Code", EmpId)
    '        oGeneralData = oGeneralService.GetByParams(oGeneralParams)

    '        Dim oRsDedcution As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        Dim Deductions As String = "SELECT IFNULL(SUM(""U_Amount""),0) From ""@SBO_PREMPSALDET1"" Where ""Code""='" & EmpId & "' "
    '        oRsDedcution.DoQuery(Deductions)
    '        Dim Deduction As Double = CDbl(oRsDedcution.Fields.Item(0).Value)
    '        '----------Benfits------------------------------------------------
    '        Dim DetailsBenQry As String = "Select ""Code"",""U_Code"",""U_Amount"" From ""@SBO_PREMPSALDET6"" Where ""Code""='" & EmpId & "'"
    '        oRsDetails.DoQuery(DetailsBenQry)
    '        Dim NetPay2 As Double = 0
    '        For i As Integer = 1 To oRsDetails.RecordCount
    '            rs.MoveFirst()
    '            For j As Integer = 1 To rs.RecordCount
    '                If oRsDetails.Fields.Item("U_Code").Value = rs.Fields.Item("Code").Value Then
    '                    Dim Update As String = "Update ""@SBO_PREMPSALDET6"" Set ""U_Amount""='" & rs.Fields.Item("U_NewSal").Value & "' Where ""Code""='" & EmpId & "' and ""U_Code""='" & oRsDetails.Fields.Item("U_Code").Value & "'"
    '                    Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                    oRsUpdate.DoQuery(Update)
    '                    NetPay2 = NetPay2 + rs.Fields.Item("U_NewSal").Value
    '                End If
    '                rs.MoveNext()
    '            Next
    '            oRsDetails.MoveNext()
    '        Next
    '        '----------------------------------------------------------------
    '        oGeneralData.SetProperty("U_NetPay", NetPay - Deduction)
    '        oGeneralData.SetProperty("U_EmpCtc", NetPay * 12)
    '        oGeneralData.SetProperty("U_EarTot", NetPay)
    '        oGeneralService.Update(oGeneralData)

    '        '-----Updating Employee Master data
    '        rs.MoveFirst()
    '        Try
    '            Dim oEmployee As SAPbobsCOM.EmployeesInfo
    '            oEmployee = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo)
    '            If (oEmployee.GetByKey(EmpId)) = True Then
    '                oEmployee.UserFields.Fields.Item("U_RDate").Value = rs.Fields.Item("U_DocDt").Value
    '                oEmployee.UserFields.Fields.Item("U_NtRevDt").Value = rs.Fields.Item("U_NtRevDt").Value
    '                oEmployee.JobTitle = rs.Fields.Item("U_PropDes").Value
    '                oEmployee.Update()
    '            End If
    '        Catch ex As Exception

    '        End Try
    '    End If
    'End Sub

    '-----Review Update---------
    Sub ReviewUpdate(ByVal AppIDV As String)

        '----------Employee review Earnings--------------------------
        str = "SELECT * FROM ""@SBO_EMPREV"" T0  INNER JOIN ""@SBO_EMPREV_C0"" T1 ON T0.""DocEntry""=T1.""DocEntry"" " &
       "WHERE T0.""U_Appid"" ='" & AppIDV & "'"
        rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs.DoQuery(str)
        If rs.RecordCount > 0 Then
            Dim EmpId, LineId As String
            EmpId = rs.Fields.Item("U_EmpId").Value
            '------Earnings---------------------------------------
            Dim oRsDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim DetailsQry As String = "Select ""Code"",""U_Code"",""U_Amount"" From ""@SBO_PREMPSALDET0"" Where ""Code""='" & EmpId & "'"
            oRsDetails.DoQuery(DetailsQry)
            Dim NetPay As Double = 0
            For i As Integer = 1 To oRsDetails.RecordCount
                rs.MoveFirst()
                For j As Integer = 1 To rs.RecordCount
                    If oRsDetails.Fields.Item("U_Code").Value = rs.Fields.Item("U_ECode").Value Then
                        Dim Update As String = "Update ""@SBO_PREMPSALDET0"" Set ""U_Amount""='" & rs.Fields.Item("U_NewSal").Value & "' Where ""Code""='" & EmpId & "' and ""U_Code""='" & oRsDetails.Fields.Item("U_Code").Value & "'"
                        Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsUpdate.DoQuery(Update)
                        NetPay = NetPay + rs.Fields.Item("U_NewSal").Value
                    End If
                    rs.MoveNext()
                Next
                oRsDetails.MoveNext()
            Next
            '-----Deductions-------------------------------------
            str = "SELECT * FROM ""@SBO_EMPREV"" T0  INNER JOIN ""@SBO_EMPREV_C2"" T1 ON T0.""DocEntry""=T1.""DocEntry"" " &
    "WHERE T0.""U_Appid"" ='" & AppIDV & "'"
            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rs.DoQuery(str)
            EmpId = rs.Fields.Item("U_EmpId").Value
            Dim DetailsDeductionQry As String = "Select ""Code"",""U_Code"",""U_Amount"" From ""@SBO_PREMPSALDET1"" Where ""Code""='" & EmpId & "'"
            oRsDetails.DoQuery(DetailsDeductionQry)
            'Dim NetPay1 As Double = 0
            For i As Integer = 1 To oRsDetails.RecordCount
                rs.MoveFirst()
                For j As Integer = 1 To rs.RecordCount
                    If oRsDetails.Fields.Item("U_Code").Value = rs.Fields.Item("U_DCode").Value Then
                        Dim Update As String = "Update ""@SBO_PREMPSALDET1"" Set ""U_Amount""='" & rs.Fields.Item("U_Amount").Value & "' Where ""Code""='" & EmpId & "' and ""U_Code""='" & oRsDetails.Fields.Item("U_Code").Value & "'"
                        Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsUpdate.DoQuery(Update)
                        'NetPay1 = NetPay1 + rs.Fields.Item("U_Amount").Value
                    End If
                    rs.MoveNext()
                Next
                oRsDetails.MoveNext()
            Next
            '------------------------------------------
            objMain.sCmp = objMain.objCompany.GetCompanyService()
            objMain.oGeneralService = objMain.sCmp.GetGeneralService("SBO_EMPSALUDO")
            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            objMain.oGeneralParams.SetProperty("Code", EmpId)
            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)

            Dim oRsDedcution As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim Deductions As String = "SELECT IFNULL(SUM(""U_Amount""),0) From ""@SBO_PREMPSALDET1"" Where ""Code""='" & EmpId & "' "
            oRsDedcution.DoQuery(Deductions)
            Dim Deduction As Double = CDbl(oRsDedcution.Fields.Item(0).Value)
            '----------Benfits------------------------------------------------
            str = "SELECT * FROM ""@SBO_EMPREV"" T0  INNER JOIN ""@SBO_EMPREV_C3"" T1 ON T0.""DocEntry""=T1.""DocEntry"" " &
                 "WHERE T0.""U_Appid"" ='" & AppIDV & "'"
            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rs.DoQuery(str)
            EmpId = rs.Fields.Item("U_EmpId").Value
            Dim DetailsBenQry As String = "Select ""Code"",""U_Code"",""U_Amount"" From ""@SBO_PREMPSALDET6"" Where ""Code""='" & EmpId & "'"
            oRsDetails.DoQuery(DetailsBenQry)
            'Dim NetPay2 As Double = 0
            For i As Integer = 1 To oRsDetails.RecordCount
                rs.MoveFirst()
                For j As Integer = 1 To rs.RecordCount
                    If oRsDetails.Fields.Item("U_Code").Value = rs.Fields.Item("Code").Value Then
                        Dim Update As String = "Update ""@SBO_PREMPSALDET6"" Set ""U_Amount""='" & rs.Fields.Item("U_Amount").Value & "' Where ""Code""='" & EmpId & "' and ""U_Code""='" & oRsDetails.Fields.Item("U_Code").Value & "'"
                        Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsUpdate.DoQuery(Update)
                        'NetPay2 = NetPay2 + rs.Fields.Item("U_Amount").Value
                    End If
                    rs.MoveNext()
                Next
                oRsDetails.MoveNext()
            Next
            '----------------------------------------------------------------
            objMain.oGeneralData.SetProperty("U_NetPay", NetPay - Deduction)
            objMain.oGeneralData.SetProperty("U_EmpCtc", NetPay * 12)
            objMain.oGeneralData.SetProperty("U_EarTot", NetPay)
            objMain.oGeneralService.Update(objMain.oGeneralData)

            '-----Updating Employee Master data
            rs.MoveFirst()
            Try
                Dim oEmployee As SAPbobsCOM.EmployeesInfo
                oEmployee = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo)
                If (oEmployee.GetByKey(EmpId)) = True Then
                    oEmployee.Position = rs.Fields.Item("U_PropDes").Value
                    oEmployee.UserFields.Fields.Item("U_RDate").Value = CDate(rs.Fields.Item("U_DocDt").Value)
                    oEmployee.UserFields.Fields.Item("U_NtRevDt").Value = CDate(rs.Fields.Item("U_NtRevDt").Value)
                    oEmployee.Update()
                End If
            Catch ex As Exception
            End Try
        End If
    End Sub
    Public Sub SendMessageAlert(ByRef oRsRecipients As SAPbobsCOM.Recordset, ByVal AppIDV As String)
        Dim oMessage As SAPbobsCOM.Message
        Dim pMessageDataColumns As SAPbobsCOM.MessageDataColumns '  MessageDataColumns
        Dim pMessageDataColumn As SAPbobsCOM.MessageDataColumn
        Dim oRecipientCollection As SAPbobsCOM.RecipientCollection
        Dim i As Integer
        Dim str, str1 As String

        Dim NextStage As String = oRsRecipients.Fields.Item("Stage").Value
        Try
            oCmpSrv = objMain.objCompany.GetCompanyService
            oMessageService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService)
            oMessage = oMessageService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage)

            Dim Receip As String = ""
            oRecipientCollection = oMessage.RecipientCollection
            For i = 0 To oRsRecipients.RecordCount - 1
                If oRsRecipients.Fields.Item("Authorizers").Value <> "" Then
                    If Not Receip.Contains(oRsRecipients.Fields.Item("Authorizers").Value) Then
                        oRecipientCollection.Add()
                        oRecipientCollection.Item(i).UserCode = oRsRecipients.Fields.Item("Authorizers").Value
                        oRecipientCollection.Item(i).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES
                        Receip = Receip & "," & oRsRecipients.Fields.Item("Authorizers").Value
                        oRsRecipients.MoveNext()
                    End If
                End If
            Next
            pMessageDataColumns = oMessage.MessageDataColumns
            pMessageDataColumn = pMessageDataColumns.Add()
            pMessageDataColumn.ColumnName = "Application Number"
            pMessageDataColumn.MessageDataLines.Add()
            pMessageDataColumn.MessageDataLines.Item(0).Value = AppIDV

            pMessageDataColumn = pMessageDataColumns.Add()
            pMessageDataColumn.ColumnName = "Requester"
            pMessageDataColumn.MessageDataLines.Add()
            pMessageDataColumn.MessageDataLines.Item(0).Value = OriginatorV


            pMessageDataColumn = pMessageDataColumns.Add()
            pMessageDataColumn.ColumnName = "Stage"
            pMessageDataColumn.MessageDataLines.Add()
            pMessageDataColumn.MessageDataLines.Item(0).Value = NextStage

            pMessageDataColumn = pMessageDataColumns.Add()
            pMessageDataColumn.ColumnName = "Send by"
            pMessageDataColumn.MessageDataLines.Add()
            pMessageDataColumn.MessageDataLines.Item(0).Value = objMain.objCompany.UserName

            oMessage.Subject = objMain.objCompany.UserName & ":Request to Approve Document"
            oMessage.Text = DraftMessage

            oMessageService.SendMessage(oMessage)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Sub SendReplyMessageAlert(ByVal RecipientUserID As String, ByVal Status As String, ByVal remarks As String, ByVal StageV As String, ByVal AppIDV As String)
        Dim oMessage As SAPbobsCOM.Message
        Dim pMessageDataColumns As SAPbobsCOM.MessageDataColumns '  MessageDataColumns
        Dim pMessageDataColumn As SAPbobsCOM.MessageDataColumn
        Dim oRecipientCollection As SAPbobsCOM.RecipientCollection
        Try
            oCmpSrv = objMain.objCompany.GetCompanyService
            oMessageService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService)
            oMessage = oMessageService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage)

            oRecipientCollection = oMessage.RecipientCollection

            oRecipientCollection.Add()
            oRecipientCollection.Item(0).UserCode = RecipientUserID
            oRecipientCollection.Item(0).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES

            pMessageDataColumns = oMessage.MessageDataColumns
            pMessageDataColumn = pMessageDataColumns.Add()
            pMessageDataColumn.ColumnName = "Application Number"
            pMessageDataColumn.MessageDataLines.Add()

            If DocumentType = "AttSum" Then
                pMessageDataColumn.MessageDataLines.Item(0).Value = "AT" & "-" & AppIDV
            ElseIf DocumentType = "PrPre" Then
                pMessageDataColumn.MessageDataLines.Item(0).Value = "PR" & "-" & AppIDV
            Else
                pMessageDataColumn.MessageDataLines.Item(0).Value = AppIDV
            End If

            pMessageDataColumn = pMessageDataColumns.Add()
            pMessageDataColumn.ColumnName = "Status"
            pMessageDataColumn.MessageDataLines.Add()
            pMessageDataColumn.MessageDataLines.Item(0).Value = Status


            pMessageDataColumn = pMessageDataColumns.Add()
            pMessageDataColumn.ColumnName = "Stage"
            pMessageDataColumn.MessageDataLines.Add()
            pMessageDataColumn.MessageDataLines.Item(0).Value = StageV


            pMessageDataColumn = pMessageDataColumns.Add()
            pMessageDataColumn.ColumnName = "Approved by"
            pMessageDataColumn.MessageDataLines.Add()
            pMessageDataColumn.MessageDataLines.Item(0).Value = objMain.objCompany.UserName

            oMessage.Subject = objMain.objCompany.UserName & ":Application Approved Status"
            oMessage.Text = remarks
            oMessageService.SendMessage(oMessage)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub




End Class





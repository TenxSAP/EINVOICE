Imports System.Security.Cryptography

Public Class clsGRID
#Region "       Declaration             "
    Public objForm, objSTRForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim str, str1, str2, str3, Query, str4, Query2 As String
    Public rs, RsNum, rs1, ORS, rs3, rsGetLeaveDetails, rsUpdateEmpLeaveQuota As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False
    Dim objutilities As Utilities
    Dim oDt As SAPbouiCOM.DataTable
    Dim oLinkedButton As SAPbouiCOM.LinkedButton
    Dim oGrid, oSubGrid As SAPbouiCOM.Grid
    Dim r1, r2, r3 As RadioButton
    Dim column, column1 As SAPbouiCOM.EditTextColumn
    Dim objGridColumn As SAPbouiCOM.GridColumn
    Dim objGridColumns As SAPbouiCOM.GridColumns
    Dim objComboGrid As SAPbouiCOM.ComboBoxColumn
    Public TemplateID, StageV, Table, AppIDField, AppStatField, OriginatorV As String
    Public AppIDV, qrystring, DocumentType As String
    Public oCmpSrv As SAPbobsCOM.CompanyService
    Public oMessageService As SAPbobsCOM.MessagesService




#End Region
    Sub CreateForm(ByVal Query1 As String)

        objMain.objUtilities.LoadForm("grid.xml", "GRID", ResourceType.Embeded)
        objForm = objMain.objApplication.Forms.GetForm("GRID", objMain.objApplication.Forms.ActiveForm.TypeCount)
        objForm.Items.Item("Item_01").TextStyle = 1

        objMain.objApplication.StatusBar.SetText("Please Wait Data is Loading.....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Me.GRID(objForm.UniqueID, Query1)
        oGrid.AutoResizeColumns()
        objMain.objApplication.StatusBar.SetText("Data is Loaded Sucessfully.....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        'Try
        '    If pVal.MenuUID = "GRID" And pVal.BeforeAction = False Then
        '        Me.CreateForm(Query1)
        '    objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        'End Try
    End Sub
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)

        Select Case pVal.EventType
            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                If pVal.ItemUID = "GRID_1" And pVal.ColUID = "Application ID" And pVal.BeforeAction = True Then

                    BubbleEvent = False

                    Dim row As Integer = pVal.Row
                    Dim AppID As String = oGrid.DataTable.GetValue("Application ID", row)
                    Dim table As String = oGrid.DataTable.GetValue("Document Table", row)

                    If AppID <> "" Then

                        ''================ SUPPLIER REBATE =================
                        'If table = "@SBO_SUPREBATE" Then

                        '    objMain.objSupplierRebate.CreateForm(FormUID, "", "", "", AppID, "", "")

                        '    Dim objNewForm As SAPbouiCOM.Form
                        '    objNewForm = objMain.objApplication.Forms.ActiveForm

                        '    objNewForm.Freeze(True)
                        '    objNewForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        '    objNewForm.Items.Item("AppID").Specific.Value = AppID
                        '    objNewForm.Items.Item("1").Click()
                        '    objNewForm.Freeze(False)

                        '    '================ PRICE CHANGE REQUEST =================
                        'ElseIf table = "@TNX_PCRT" Then

                        '    objMain.ObjPriceChangeRequest.CreateForm(AppID)

                        '    Dim objNewForm As SAPbouiCOM.Form
                        '    objNewForm = objMain.objApplication.Forms.ActiveForm

                        '    objNewForm.Freeze(True)
                        '    objNewForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        '    objNewForm.Items.Item("AppID").Specific.Value = AppID
                        '    objNewForm.Items.Item("1").Click()
                        '    objNewForm.Freeze(False)

                        '    'madhu item listing screen
                        'ElseIf table = "@SBO_LIST" Then

                        '    objMain.objItemList.CreateForm_Find(AppID)

                        '    Dim objNewForm As SAPbouiCOM.Form
                        '    objNewForm = objMain.objApplication.Forms.ActiveForm

                        '    objNewForm.Freeze(True)
                        '    objNewForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        '    objNewForm.Items.Item("AppID").Specific.Value = AppID
                        '    objNewForm.Items.Item("1").Click()
                        '    objNewForm.Freeze(False)

                        'ElseIf table = "@TNX_ITEMUP" Then


                        '    Dim objGRForm As SAPbouiCOM.Form

                        '    Try

                        '        objMain.objApplication.ActivateMenuItem("ITEMUP")
                        '        objGRForm = objMain.objApplication.Forms.ActiveForm
                        '        objGRForm.Freeze(True)
                        '        objGRForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        '        objGRForm.Items.Item("AppID").Specific.Value = AppID
                        '        objGRForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        '        Dim oDS As SAPbouiCOM.DBDataSource
                        '        oDS = objGRForm.DataSources.DBDataSources.Item("@TNX_ITEMUP")
                        '        objGRForm.Freeze(False)

                        '    Catch ex As Exception

                        '        objGRForm.Freeze(False)

                        '    End Try


                        'End If
                    End If
                End If


            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                If pVal.ItemUID = "41" And pVal.BeforeAction = False Then
                    objForm.Close()
                End If
                '--------------------------------

                '-------------

                If pVal.ItemUID = "GRID_1" And pVal.ColUID = "RowsHeader" And pVal.BeforeAction = False Then
                    'Me.subgrid(objForm.UniqueID)
                    Me.GridDetails(objForm.UniqueID, pVal)

                End If


                If pVal.ItemUID = "SUBGRID" And pVal.ColUID = "RowsHeader" And pVal.BeforeAction = False Then

                End If

                'If pVal.ItemUID = "1" And pVal.BeforeAction = True Then
                '    For i As Integer = 0 To oSubGrid.DataTable.Rows.Count - 1
                '        Dim AppID As String = oSubGrid.DataTable.GetValue("Application Num", i)




                '    Next

                'End If

                If pVal.ItemUID = "1" And pVal.BeforeAction = True Then


                    Dim AUTH, stage As String
                    Dim ROWID As Integer
                    Dim AppIDV, authorizer, StageStatus, STATUS, Approvaldocstatus, ORGINATOR, remarks, StageSts, AppID, Approvaldocsts, DocType As String

                    For i As Integer = 0 To oSubGrid.DataTable.Rows.Count - 1
                        AUTH = oSubGrid.DataTable.GetValue("Authorizer", i)
                        stage = oSubGrid.DataTable.GetValue("Approval Stage Status", i)
                        StageSts = oSubGrid.DataTable.GetValue("Approval Stage", i)
                        AppID = oSubGrid.DataTable.GetValue("Application Num", i)
                        TemplateID = oSubGrid.DataTable.GetValue("Template ID", i)
                        Approvaldocsts = oSubGrid.DataTable.GetValue("Approval Doc Status", i)

                        If Me.Validate(FormUID, AppID, TemplateID) = False Then
                            BubbleEvent = False
                            Exit Sub   ' <<< THIS will stop entire execution
                        End If

                        Dim oRsCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim CountQry As String = "SELECT IFNULL(T0.""U_NAP"",0) As ""NAP"", IFNULL(T0.""U_NRJ"",0) As ""NRJ"" FROM ""@SBO_AST""  T0 WHERE T0.""Code"" ='" & StageSts & "'"
                        oRsCount.DoQuery(CountQry)
                        Dim NoOfApp As Double = oRsCount.Fields.Item("NAP").Value
                        Dim NoOfRej As Double = oRsCount.Fields.Item("NRJ").Value
                        Dim rsup3 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        Dim Qry As String = "SELECT COUNT(*) FROM ""@SBO_DD"" t1 INNER JOIN ""@SBO_DD1"" t2 ON t1.""DocEntry"" = t2.""DocEntry"""
                        Qry = Qry & " WHERE ""U_docnum"" = '" & AppID & "' and t2.""U_statusN"" ='A' and t2.""U_Stage""='" & StageSts & "'"
                        rsup3.DoQuery(Qry)
                        Dim ApprCount As Double = rsup3.Fields.Item(0).Value

                        Dim USERNAME As String = objMain.objCompany.UserName
                        '  If AUTH = USERNAME And stage = "Submitted" And Approvaldocsts = "Submitted" And ApprCount < NoOfApp Then
                        If AUTH = USERNAME And stage = "Submitted" And ApprCount < NoOfApp Then

                            '    Dim USERNAME As String = objMain.objCompany.UserName
                            'If AUTH = USERNAME And stage = "Submitted" Then
                            ROWID = i
                            AppIDV = oSubGrid.DataTable.GetValue("Application Num", i)
                            StageStatus = oSubGrid.DataTable.GetValue("Approval Stage", i)
                            authorizer = oSubGrid.DataTable.GetValue("Authorizer", i)
                            STATUS = oSubGrid.DataTable.GetValue("Approval Doc Status", i)
                            remarks = oSubGrid.DataTable.GetValue("Remarks", i)
                            ORGINATOR = oSubGrid.DataTable.GetValue("Requester Emp ID", i)
                            DocType = oSubGrid.DataTable.GetValue("Doc Type", i)
                            Exit For

                        End If
                    Next
                    rs1 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    rs1 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim Qry1 As String = "Select * From ""@SBO_DD"" T0 INNER JOIN ""@SBO_DD1"" T1 ON T0.""DocEntry""=T1.""DocEntry"" Where T0.""U_docnum""='" & AppIDV & "' and T1.""U_Userid""='" & authorizer & "' and T1.""U_Stage""='" & StageStatus & "' "
                    rs1.DoQuery(Qry1)
                    DocumentType = rs1.Fields.Item("U_objtype").Value
                    TemplateID = rs1.Fields.Item("U_tempid").Value
                    Table = rs1.Fields.Item("U_Table").Value
                    AppIDField = rs1.Fields.Item("U_AppID").Value
                    AppStatField = rs1.Fields.Item("U_AppStat").Value
                    '  Dim Status1 As String = rs1.Fields.Item("U_statusN").Value/
                    If STATUS = "A" Or STATUS = "N" Then

                        If updateDraftTable(STATUS, StageStatus, TemplateID, remarks, DocumentType, authorizer, AppIDV, DocType) = True Then


                            objMain.objApplication.StatusBar.SetText("Operation Completed Sucessfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                        End If
                        '  ElseIf STATUS = "R" Or STATUS = "C" Then
                    Else
                        objMain.objApplication.StatusBar.SetText("Operation Completed Sucessfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                    End If

                    'objMain.sCmp = objMain.objCompany.GetCompanyService
                    'objMain.oGeneralService = objMain.sCmp.GetGeneralService("SBO_Draft")
                    'objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                    'objMain.oGeneralParams.SetProperty("DocEntry", rs1.Fields.Item("DocEntry").Value)
                    'objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                    'objMain.oChildren = objMain.oGeneralData.Child("SBO_WDD1")
                    '' objMain.oGeneralData.SetProperty("U_statusN", STATUS)

                    'objMain.oChild = objMain.oChildren.Item(rs1.Fields.Item("LineId").Value - 1)
                    'objMain.oChild.SetProperty("U_statusN", STATUS)
                    ''  objMain.oChild.SetProperty("U_MenuNM", MenuName)
                    'objMain.oGeneralService.Update(objMain.oGeneralData)
                    ' updateDraftTable(ByVal Status As String, ByVal Stage As String, ByVal tempid As String, ByVal Remarks As String, ByVal DocumentType As String, ByVal APPID As String)
                    'updateDraftTable(STATUS, stage, TemplateID, remarks, DocumentType, AppIDV)
                    'test1

                End If
                '                '  

        End Select
    End Sub
    Public Function Validate(ByVal FormUID As String, ByVal AppID As String, ByVal TemplateID As String) As Boolean

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)


            Dim rs4 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim QRYS As String = "Select ""U_FOUT"" from ""@TNX_PCRT"" where ""U_AppID"" = '" & AppID & "'"
            rs4.DoQuery(QRYS)

            If TemplateID = "FLIER" Then
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
    Sub GridDetails(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent)


        Try
            oGrid = objForm.Items.Item("GRID_1").Specific

            Dim AppID, TableNM As String
            'For i As Integer = 1 To oGrid.Rows.Count
            '    If oGrid.Rows.IsSelected(i) = True Then
            Dim Test As Integer = pVal.Row
            AppID = oGrid.DataTable.GetValue("Application ID", Test)
            TableNM = oGrid.DataTable.GetValue("Document", Test)
            If AppID <> "" Then
                Me.subgrid(FormUID, AppID, TableNM)

            End If

            '        Exit Try
            '    End If
            'Next

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub subgrid(ByVal FormUID As String, ByVal AppID As String, ByVal DocType As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)
            oSubGrid = objForm.Items.Item("SUBGRID").Specific
            oSubGrid.DataTable = objForm.DataSources.DataTables.Item("DT_GRD")
            oSubGrid.DataTable.Clear()
            Query2 = "CALL ""TNX_ApprovalStatusDetailReport""('" & AppID & "','" & DocType & "')"
            oSubGrid.DataTable.ExecuteQuery(Query2)
            If oSubGrid.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("No records found")
                Exit Sub

            End If
            DISABLE()

            Dim AUTH, stage, Approvaldocstatus, StageStatus As String
            For i As Integer = 0 To oSubGrid.DataTable.Rows.Count - 1
                AUTH = oSubGrid.DataTable.GetValue("Authorizer", i)
                stage = oSubGrid.DataTable.GetValue("Approval Stage Status", i)
                Approvaldocstatus = oSubGrid.DataTable.GetValue("Approval Doc Status", i)
                StageStatus = oSubGrid.DataTable.GetValue("Approval Stage", i)

                Dim oRsCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim CountQry As String = "SELECT IFNULL(T0.""U_NAP"",0) As ""NAP"", IFNULL(T0.""U_NRJ"",0) As ""NRJ"" FROM ""@SBO_AST""  T0 WHERE T0.""Code"" ='" & StageStatus & "'"
                oRsCount.DoQuery(CountQry)
                Dim NoOfApp As Double = oRsCount.Fields.Item("NAP").Value
                Dim NoOfRej As Double = oRsCount.Fields.Item("NRJ").Value
                Dim rsup3 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                Dim Qry As String = "SELECT COUNT(*) FROM ""@SBO_OWDD"" t1 INNER JOIN ""@SBO_WDD1"" t2 ON t1.""DocEntry"" = t2.""DocEntry"""
                Qry = Qry & " WHERE ""U_docnum"" = '" & AppID & "' and t2.""U_statusN"" ='A' and ""U_Stage""='" & StageStatus & "'"
                rsup3.DoQuery(Qry)
                Dim ApprCount As Double = rsup3.Fields.Item(0).Value

                Dim USERNAME As String = objMain.objCompany.UserName
                If AUTH = USERNAME And stage = "Submitted" And Approvaldocstatus = "Submitted" And ApprCount < NoOfApp Then

                    'Dim USERNAME As String = objMain.objCompany.UserName
                    'If AUTH = USERNAME And stage = "Submitted" And Approvaldocstatus = "Submitted" Then
                    oSubGrid.CommonSetting.SetCellEditable(i + 1, 13, True)
                    oSubGrid.CommonSetting.SetCellEditable(i + 1, 14, True)


                End If
            Next
            oSubGrid.Columns.Item("Approval Doc Status").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            objComboGrid = oSubGrid.Columns.Item("Approval Doc Status")
            objComboGrid.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
            'objComboGrid.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objComboGrid.ValidValues.Add("A", "Approved")
            objComboGrid.ValidValues.Add("N", "Reject")
            objComboGrid.ValidValues.Add("C", "Cancelled")

            'Dim oRsCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'Dim CountQry As String = "SELECT IFNULL(T0.""U_NAP"",0) As ""NAP"", IFNULL(T0.""U_NRJ"",0) As ""NRJ"" FROM ""@SBO_AST""  T0 WHERE T0.""Code"" ='" & StageStatus & "'"
            'oRsCount.DoQuery(CountQry)
            'Dim NoOfApp As Double = oRsCount.Fields.Item("NAP").Value
            'Dim NoOfRej As Double = oRsCount.Fields.Item("NRJ").Value
            'Dim rsup3 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'Dim Qry As String = "SELECT COUNT(*) FROM ""@SBO_OWDD"" t1 INNER JOIN ""@SBO_WDD1"" t2 ON t1.""DocEntry"" = t2.""DocEntry"""
            'Qry = Qry & " WHERE ""U_docnum"" = '" & AppID & "' and t2.""U_statusN"" ='A' and ""U_Stage""='" & StageStatus & "'"
            'rsup3.DoQuery(Qry)
            'Dim ApprCount As Double = rsup3.Fields.Item(0).Value

            'Dim USERNAME As String = objMain.objCompany.UserName
            'If AUTH = USERNAME And stage = "Submitted" And Approvaldocstatus = "Submitted" And ApprCount >= NoOfApp Then
            '    oSubGrid.CommonSetting.SetCellEditable(i + 1, 13, True)

            'If  Then
            '    Dim NextLineQry As String = "SELECT T3.""LineId"" + 1 FROM ""@SBO_APPHDR"" T0 " &
            '    " INNER JOIN ""@SBO_APPAUT"" T3 ON T3.""Code"" = T0.""Code"" " &
            '    " WHERE ""U_Active"" = 'Y' AND T0.""Code""='" & TemplateID & "' and T3.""U_Name""='" & stage & "' "
            '    Dim oRsNextLIne As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '    oRsNextLIne.DoQuery(NextLineQry)
            '    Dim LineId As String = oRsNextLIne.Fields.Item(0).Value

            '    Dim NextAuthQry As String = "SELECT T0.""Code"" As ""TemplateID"",S1.""Code"" As ""Stage"",S2.""U_UKey"" As ""Authorizers"" " &
            '    " ,S1.""U_NAP"" As ""NoOfAuth"",S1.""U_NRJ"" As ""NoOfRej"",S2.""LineId"" FROM ""@SBO_APPHDR"" T0 " &
            '    " INNER JOIN ""@SBO_APPAUT"" T3 ON T3.""Code"" = T0.""Code"" INNER JOIN ""@SBO_AST"" S1 ON T3.""U_Name""=S1.""Code"" " &
            '    " INNER JOIN ""@SBO_AST_C0"" S2 ON S1.""Code""=S2.""Code"" WHERE ""U_Active"" = 'Y' AND IFNULL(S2.""U_UKey"",'')<>'' " &
            '    " AND T0.""Code""='" & TemplateID & "' and T3.""LineId""='" & LineId & "';"
            '    Dim oRsNextAuth As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '    oRsNextAuth.DoQuery(NextAuthQry)
            'End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub

    Sub GRID(ByVal FormUID As String, ByVal Query1 As String)
        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("GRID_1").Specific

            oGrid.DataTable = objForm.DataSources.DataTables.Item("DT_GRID")
            oGrid.DataTable.Clear()

            Query1 = Query1
            oGrid.DataTable.ExecuteQuery(Query1)

            If oGrid.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("No records found")
                Exit Sub
            End If
            oGrid.AutoResizeColumns()
            For i As Integer = 0 To oGrid.DataTable.Columns.Count - 1
                oGrid.Columns.Item(i).Editable = False
            Next
            ' oGrid.CollapseLevel = 1
            oGrid.Columns.Item("Document Table").Visible = False
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            column = oGrid.Columns.Item("Application ID")
            column.LinkedObjectType = "2"

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Function DISABLE()
        oSubGrid.Columns.Item("Doc Type").Editable = False
        oSubGrid.Columns.Item("DocEntry").Editable = False
        oSubGrid.Columns.Item("Type").Editable = False
        oSubGrid.Columns.Item("Application Num").Editable = False
        oSubGrid.Columns.Item("Requester SAP ID").Editable = False
        oSubGrid.Columns.Item("Requester Emp ID").Editable = False
        oSubGrid.Columns.Item("Requester Name").Editable = False
        oSubGrid.Columns.Item("Requester Branch").Editable = False
        oSubGrid.Columns.Item("DocumentDate").Editable = False
        oSubGrid.Columns.Item("Create Date").Editable = False
        oSubGrid.Columns.Item("Approval Doc Status").Editable = False
        oSubGrid.Columns.Item("Approver SAP ID").Editable = False
        oSubGrid.Columns.Item("Approver Name").Editable = False
        oSubGrid.Columns.Item("Approval Stage Status").Editable = False
        oSubGrid.Columns.Item("Approver Stage").Editable = False
        oSubGrid.Columns.Item("Approver Emp ID").Editable = False
        oSubGrid.Columns.Item("Authorizer").Editable = False
        oSubGrid.Columns.Item("Approval Stage").Editable = False
        oSubGrid.Columns.Item("Template ID").Editable = False
        oSubGrid.Columns.Item("Remarks").Editable = False

    End Function

    Public Function updateDraftTable(ByVal Status As String, ByVal Stage As String, ByVal tempid As String, ByVal Remarks As String, ByVal DocumentType As String, ByVal authorizer As String, ByVal APPIDV As String, ByVal DocType As String)
        Try

            Dim oRsDraftDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim DraftDetQry As String = "Select * From ""@SBO_DD"" T0 INNER JOIN ""@SBO_DD1"" T1 ON T0.""DocEntry""=T1.""DocEntry"" " &
            "Where T0.""U_docnum""='" & APPIDV & "' and T1.""U_Userid""='" & authorizer & "' and T1.""U_Stage""='" & Stage & "' And T0.""U_Doc"" = '" & DocType & "'"
            oRsDraftDetails.DoQuery(DraftDetQry)
            Table = oRsDraftDetails.Fields.Item("U_Table").Value
            AppIDField = oRsDraftDetails.Fields.Item("U_AppID").Value
            AppStatField = oRsDraftDetails.Fields.Item("U_AppStat").Value
            OriginatorV = oRsDraftDetails.Fields.Item("U_userid").Value


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
            Dim CountQry As String = "SELECT IFNULL(T0.""U_NAP"",0) As ""NAP"", IFNULL(T0.""U_NRJ"",0) As ""NRJ"" FROM ""@SBO_AST""  T0 WHERE T0.""Code"" ='" & Stage & "'"
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
                Try
                    strUpdate = "UPDATE ""@SBO_DD1"" SET ""U_statusN""='A',""U_Remarks""='" & Remarks & "',""U_Udate""='" & DateTime.Now.ToString("yyyyMMdd") & "',""U_Utime""='" & DateTime.Now.ToString("HHmm") & "' " &
                    " where ""U_Userid""='" & objMain.objCompany.UserName & "' and ""U_Stage""='" & Stage & "' and ""DocEntry""=(select ""DocEntry"" from ""@SBO_DD"" where ""U_docnum""='" & APPIDV & "' And ""U_Doc"" = '" & DocType & "' And ""U_statusN"" = 'S')"
                    rsup1.DoQuery(strUpdate)
                    If DocType = "Salary Setup" Then
                        '        objMain.objUtilities.AddErrorLog(objForm.UniqueID, APPIDV, "Method Add", strUpdate)
                    End If
                Catch ex As Exception
                    '      objMain.objUtilities.AddErrorLog(objForm.UniqueID, APPIDV, "Method Add", ex.Message)
                    objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End Try

                Dim Qry As String = "SELECT COUNT(*) FROM ""@SBO_DD"" t1 INNER JOIN ""@SBO_DD1"" t2 ON t1.""DocEntry"" = t2.""DocEntry"""
                Qry = Qry & " WHERE ""U_docnum"" = '" & APPIDV & "' and t2.""U_statusN"" ='A' and t2.""U_Stage""='" & Stage & "' And t1.""U_Doc"" = '" & DocType & "'"
                rsup3.DoQuery(Qry)
                Dim ApprCount As Double = rsup3.Fields.Item(0).Value

                If ApprCount >= NoOfApprovals Then
                    'Dim NextLineQry As String = "SELECT T3.""LineId"" + 1 FROM ""@SBO_APPHDR"" T0 " &
                    '" INNER JOIN ""@SBO_APPAUT"" T3 ON T3.""Code"" = T0.""Code"" " &
                    '" WHERE ""U_Active"" = 'Y' AND T0.""Code""='" & TemplateID & "' and T3.""U_Name""='" & Stage & "' "
                    Dim NextLineQry As String = "SELECT T3.""LineId"" + 1 FROM ""@SBO_APPHDR"" T0 " &
                    " INNER JOIN ""@SBO_APPAUT"" T3 ON T3.""Code"" = T0.""Code"" " &
                    " WHERE ""U_Active"" = 'Y' AND T0.""Code""='" & TemplateID & "' and T3.""U_NAMES""='" & Stage & "' "
                    Dim oRsNextLIne As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsNextLIne.DoQuery(NextLineQry)
                    Dim LineId As String = oRsNextLIne.Fields.Item(0).Value

                    'Dim NextAuthQry As String = "SELECT T0.""Code"" As ""TemplateID"",S1.""Code"" As ""Stage"",S2.""U_UKey"" As ""Authorizers"" " &
                    '" ,S1.""U_NAP"" As ""NoOfAuth"",S1.""U_NRJ"" As ""NoOfRej"",S2.""LineId"" FROM ""@SBO_APPHDR"" T0 " &
                    '" INNER JOIN ""@SBO_APPAUT"" T3 ON T3.""Code"" = T0.""Code"" INNER JOIN ""@SBO_AST"" S1 ON T3.""U_Name""=S1.""Code"" " &
                    '" INNER JOIN ""@SBO_AST_C0"" S2 ON S1.""Code""=S2.""Code"" WHERE ""U_Active"" = 'Y' AND IFNULL(S2.""U_UKey"",'')<>'' " &
                    '" AND T0.""Code""='" & TemplateID & "' and T3.""LineId""='" & LineId & "';"
                    Dim NextAuthQry As String = "SELECT T0.""Code"" As ""TemplateID"",S1.""Code"" As ""Stage"",S2.""U_UKey"" As ""Authorizers"" " &
                    " ,S1.""U_NAP"" As ""NoOfAuth"",S1.""U_NRJ"" As ""NoOfRej"",S2.""LineId"" FROM ""@SBO_APPHDR"" T0 " &
                    " INNER JOIN ""@SBO_APPAUT"" T3 ON T3.""Code"" = T0.""Code"" INNER JOIN ""@SBO_AST"" S1 ON T3.""U_NAMES""=S1.""Code"" " &
                    " INNER JOIN ""@SBO_AST_C0"" S2 ON S1.""Code""=S2.""Code"" WHERE ""U_Active"" = 'Y' AND IFNULL(S2.""U_UKey"",'')<>'' " &
                    " AND T0.""Code""='" & TemplateID & "' and T3.""LineId""='" & LineId & "';"

                    Dim oRsNextAuth As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsNextAuth.DoQuery(NextAuthQry)
                    If oRsNextAuth.RecordCount > 0 Then
                        Dim NextStage As String = oRsNextAuth.Fields.Item("Stage").Value
                        Dim DraftQry As String = "UPDATE ""@SBO_DD1"" SET ""U_statusN""='S',""U_Remarks""='" & Remarks & "',""U_Udate""='" & DateTime.Now.ToString("yyyyMMdd") & "',""U_Utime""='" & DateTime.Now.ToString("HHmm") & "'  " &
                        " where ""U_Stage""='" & NextStage & "' and ""DocEntry""=(select ""DocEntry"" from ""@SBO_DD"" where ""U_docnum""='" & APPIDV & "' And ""U_Doc"" = '" & DocType & "')"
                        Dim oRsDraft As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsDraft.DoQuery(DraftQry)
                        Me.SendMessageAlert(oRsNextAuth, APPIDV)

                    Else '-----after completing all the stages
                        Dim DraftMasterQry As String = "UPDATE ""@SBO_DD"" SET ""U_statusN""='A',""U_Isdraft""='N' where ""U_docnum""='" & APPIDV & "' And ""U_Doc"" = '" & DocType & "'"
                        oRsDraftMaster.DoQuery(DraftMasterQry)
                        Dim DocUpdateQry As String = "UPDATE """ & Table & """ SET """ & AppStatField & """='A' where """ & AppIDField & """='" & APPIDV & "'"
                        oRsDocUpdate.DoQuery(DocUpdateQry)
                        Me.SendReplyMessageAlert(OriginatorV, Status, Remarks, Stage, APPIDV)
                        'DocumentType = oDS.GetValue("U_objtype", 0)
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
                        ElseIf DocumentType = "EmpSal" Then
                            'Me.UpdateEmpSalarySetup(APPIDV)
                        ElseIf DocumentType = "EOS" Then
                            Me.FinalSettlement(APPIDV)
                        ElseIf DocumentType = "EmpMstr" Then
                            Dim oEmploy As SAPbobsCOM.EmployeesInfo = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo)
                            If oEmploy.GetByKey(CInt(APPIDV)) = True Then
                                oEmploy.TreminationReason = "7"
                                If oEmploy.Update <> 0 Then
                                    objMain.objApplication.StatusBar.SetText("Failed to update employee status with empID '" & APPIDV & "'   -->", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                            End If
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
                'rsup1.DoQuery("UPDATE ""@SBO_WDD1"" SET ""U_statusN""='N',""U_Remarks""='" & Remarks & "',""U_Udate""='" & DateTime.Now.ToString("yyyyMMdd") & "',""U_Utime""='" & DateTime.Now.ToString("HHmm") & "' " &
                '    " where ""U_Userid""='" & objMain.objCompany.UserName & "' and ""U_Stage""='" & Stage & "' and ""DocEntry""=(select ""DocEntry"" from ""@SBO_OWDD"" where ""U_docnum""='" & APPIDV & "' And ""U_Doc"" = '" & DocType & "')")

                'Dim Qry As String = "SELECT COUNT(*) FROM ""@SBO_OWDD"" t1 INNER JOIN ""@SBO_WDD1"" t2 ON t1.""DocEntry"" = t2.""DocEntry"""
                'Qry = Qry & " WHERE ""U_docnum"" = '" & APPIDV & "' and t2.""U_statusN"" ='N' and ""U_Stage""='" & Stage & "' And t1.""U_Doc"" = '" & DocType & "'"
                'rsup3.DoQuery(Qry)
                'Dim RegCount As Double = rsup3.Fields.Item(0).Value

                'If RegCount >= NoOfRejects Then '

                '    oRsDocUpdate.DoQuery("UPDATE """ & Table & """ SET """ & AppStatField & """='N' where """ & AppIDField & """='" & APPIDV & "'")
                '    oRsDraftMaster.DoQuery("UPDATE ""@SBO_OWDD"" SET ""U_statusN"" ='N',""U_Isdraft""='N' where ""U_docnum""='" & APPIDV & "' And ""U_Doc"" = '" & DocType & "'")
                '    Me.SendReplyMessageAlert(OriginatorV, Status, Remarks, Stage, APPIDV)

                'End If             ********* Comment on 20240822
                rsup1.DoQuery("UPDATE ""@SBO_DD1"" SET ""U_statusN""='N',""U_Remarks""='" & Remarks & "',""U_Udate""='" & DateTime.Now.ToString("yyyyMMdd") & "',""U_Utime""='" & DateTime.Now.ToString("HHmm") & "' " &
                    " where ""U_Userid""='" & objMain.objCompany.UserName & "' and ""U_Stage""='" & Stage & "' and ""DocEntry""=(select ""DocEntry"" from ""@SBO_DD"" where ""U_docnum""='" & APPIDV & "' And ""U_Doc"" = '" & DocType & "')")


                Dim Qry As String = "SELECT COUNT(*) FROM ""@SBO_DD"" t1 INNER JOIN ""@SBO_DD1"" t2 ON t1.""DocEntry"" = t2.""DocEntry"""
                Qry = Qry & " WHERE ""U_docnum"" = '" & APPIDV & "' and t2.""U_statusN"" ='N' and ""U_Stage""='" & Stage & "' And t1.""U_Doc"" = '" & DocType & "'"
                rsup3.DoQuery(Qry)
                Dim RegCount As Double = rsup3.Fields.Item(0).Value

                If RegCount >= NoOfRejects Then
                    oRsDocUpdate.DoQuery("UPDATE """ & Table & """ SET """ & AppStatField & """='N' where """ & AppIDField & """='" & APPIDV & "'")
                    oRsDraftMaster.DoQuery("UPDATE ""@SBO_DD"" SET ""U_statusN"" ='N',""U_Isdraft""='N' where ""U_docnum""='" & APPIDV & "' And ""U_Doc"" = '" & DocType & "'")
                    Me.SendReplyMessageAlert(OriginatorV, Status, Remarks, Stage, APPIDV)
                End If

                If DocumentType = "LvApp" Then
                    Me.ReverseLeaveUpdate(APPIDV)
                ElseIf DocumentType = "LvRj" Then
                    Me.ReverseRejoinUpdate(APPIDV)
                ElseIf DocumentType = "Resign" Then
                    Me.ReverseResignUpdate(APPIDV)
                ElseIf DocumentType = "EOS" Then
                    Me.RejectFinalSetUpdate(APPIDV)
                End If
            End If

            If DocumentType = "OL" Then
                UpdateOfferLetter(APPIDV, Status)
            ElseIf DocumentType = "AReg" Then
                Me.UpdateRegularization(Status)
            End If

            'System.Runtime.InteropServices.Marshal.ReleaseComObject(rsup1)
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(rsup2)
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(rsup3)
            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try
    End Function
    Private Sub RejectFinalSetUpdate(ByVal APPIDV As String)
        Try
            Dim FindCode As String = "Select ""Code"" from ""@SBO_EOSPMSTR"" Where ""U_Appid"" = '" & APPIDV & "'"
            Dim oRsFindCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsFindCode.DoQuery(FindCode)
            If oRsFindCode.RecordCount > 0 Then
                Dim EmpCode As String = oRsFindCode.Fields.Item("Code").Value

                Dim UpdateMstrCancel As String = "Update ""@SBO_EOSPMSTR"" A Set A.""U_AppStat"" = 'C',A.""Code"" = A.""Code"" || 'C' Where A.""Code"" = '" & EmpCode & "'"
                Dim oRsUpdateMstrCancel As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsUpdateMstrCancel.DoQuery(UpdateMstrCancel)
                Dim UpdateC0Cancel As String = "Update ""@SBO_EOSP_C0"" A Set A.""Code"" = A.""Code"" || 'C' Where A.""Code"" = '" & EmpCode & "'"
                Dim oRsUpdateC0Cancel As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsUpdateC0Cancel.DoQuery(UpdateC0Cancel)
                Dim UpdateC1Cancel As String = "Update ""@SBO_EOSP_C1"" A Set A.""Code"" = A.""Code"" || 'C' Where A.""Code"" = '" & EmpCode & "'"
                Dim oRsUpdateC1Cancel As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsUpdateC1Cancel.DoQuery(UpdateC1Cancel)
                Dim UpdateC2Cancel As String = "Update ""@SBO_EOSP_C2"" A Set A.""Code"" = A.""Code"" || 'C' Where A.""Code"" = '" & EmpCode & "'"
                Dim oRsUpdateC2Cancel As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsUpdateC2Cancel.DoQuery(UpdateC2Cancel)
                Dim UpdateC3Cancel As String = "Update ""@SBO_EOSP_C3"" A Set A.""Code"" = A.""Code"" || 'C' Where A.""Code"" = '" & EmpCode & "'"
                Dim oRsUpdateC3Cancel As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsUpdateC3Cancel.DoQuery(UpdateC3Cancel)
                Dim UpdateC4Cancel As String = "Update ""@SBO_EOSP_C4"" A Set A.""Code"" = A.""Code"" || 'C' Where A.""Code"" = '" & EmpCode & "'"
                Dim oRsUpdateC4Cancel As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsUpdateC4Cancel.DoQuery(UpdateC4Cancel)
                Dim UpdateC5Cancel As String = "Update ""@SBO_EOSP_C5"" A Set A.""Code"" = A.""Code"" || 'C' Where A.""Code"" = '" & EmpCode & "'"
                Dim oRsUpdateC5Cancel As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsUpdateC5Cancel.DoQuery(UpdateC5Cancel)
                Dim UpdateC6Cancel As String = "Update ""@SBO_EOSP_C6"" A Set A.""Code"" = A.""Code"" || 'C' Where A.""Code"" = '" & EmpCode & "'"
                Dim oRsUpdateC6Cancel As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsUpdateC6Cancel.DoQuery(UpdateC6Cancel)
                Dim UpdateC7Cancel As String = "Update ""@SBO_EOSP_C7"" A Set A.""Code"" = A.""Code"" || 'C' Where A.""Code"" = '" & EmpCode & "'"
                Dim oRsUpdateC7Cancel As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsUpdateC7Cancel.DoQuery(UpdateC7Cancel)
                Dim UpdateC8Cancel As String = "Update ""@SBO_EOSP_C8"" A Set A.""Code"" = A.""Code"" || 'C' Where A.""Code"" = '" & EmpCode & "'"
                Dim oRsUpdateC8Cancel As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsUpdateC8Cancel.DoQuery(UpdateC8Cancel)
                Dim UpdateC9Cancel As String = "Update ""@SBO_EOSP_C9"" A Set A.""Code"" = A.""Code"" || 'C' Where A.""Code"" = '" & EmpCode & "'"
                Dim oRsUpdateC9Cancel As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsUpdateC9Cancel.DoQuery(UpdateC9Cancel)
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub ReverseLeaveUpdate(ByVal APPIDV As String)
        Try
            Dim LeaveAppDet As String = "Select * from ""@SBO_PRLEVAPPMSTR"" Where ""U_Appid"" = '" & APPIDV & "'"
            Dim oRsLeaveAppDet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsLeaveAppDet.DoQuery(LeaveAppDet)
            Dim oEmp As SAPbobsCOM.EmployeesInfo = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo)
            If oEmp.GetByKey(CInt(oRsLeaveAppDet.Fields.Item("U_empid").Value)) = True Then
                Dim GetActiveID As String = "Select ""reasonID"" From OHTR Where ""reasonID""='" & oEmp.TreminationReason & "' And ""reasonID"" IS NOT NULL"
                Dim oRsGetActiveID As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetActiveID.DoQuery(GetActiveID)

                If oRsGetActiveID.RecordCount > 0 Then

                    oEmp.TreminationReason = CStr(oEmp.UserFields.Fields.Item("U_PRSTS").Value).ToString
                    oEmp.UserFields.Fields.Item("U_PRSTS").Value = CStr(oRsGetActiveID.Fields.Item(0).Value).ToString
                    oEmp.UserFields.Fields.Item("U_STUPDT").Value = DateTime.Now.ToString

                    If oEmp.Update = 0 Then
                        objMain.objApplication.StatusBar.SetText("Status Changed from  On Leave to" & CStr(oEmp.UserFields.Fields.Item("U_PRSTS").Value).ToString & "  with empID '" & oRsLeaveAppDet.Fields.Item("U_empid").Value & "'   ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Else
                        objMain.objApplication.StatusBar.SetText("Failed to update employee master  with empID '" & oRsLeaveAppDet.Fields.Item("U_empid").Value & "'   -->", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If

                Else
                    objMain.objApplication.StatusBar.SetText("ReasonID not fetched  ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Try
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    Private Sub ReverseRejoinUpdate(ByVal APPIDV As String)
        Try
            Dim LeaveAppDet As String = "Select * from ""@SBO_NOTIF"" Where ""U_Appid"" = '" & APPIDV & "'"
            Dim oRsLeaveAppDet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsLeaveAppDet.DoQuery(LeaveAppDet)
            Dim oEmp As SAPbobsCOM.EmployeesInfo = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo)
            If oEmp.GetByKey(CInt(oRsLeaveAppDet.Fields.Item("U_EmpId").Value)) = True Then
                Dim GetActiveID As String = "Select ""reasonID"" From OHTR Where ""name""='" & oEmp.TreminationReason & "' And ""reasonID"" IS NOT NULL"
                Dim oRsGetActiveID As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetActiveID.DoQuery(GetActiveID)

                If oRsGetActiveID.RecordCount > 0 Then

                    oEmp.TreminationReason = CStr(oEmp.UserFields.Fields.Item("U_PRSTS").Value).ToString
                    oEmp.UserFields.Fields.Item("U_PRSTS").Value = CStr(oRsGetActiveID.Fields.Item(0).Value).ToString
                    oEmp.UserFields.Fields.Item("U_STUPDT").Value = DateTime.Now.ToString

                    If oEmp.Update = 0 Then
                        objMain.objApplication.StatusBar.SetText("Status Changed to  " & CStr(oEmp.UserFields.Fields.Item("U_PRSTS").Value).ToString & "  with empID '" & oRsLeaveAppDet.Fields.Item("U_EmpId").Value & "'   ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Else
                        objMain.objApplication.StatusBar.SetText("Failed to update employee master  with empID '" & oRsLeaveAppDet.Fields.Item("U_EmpId").Value & "'   -->", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If

                Else
                    objMain.objApplication.StatusBar.SetText("ReasonID not fetched  ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Try
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    Private Sub ReverseResignUpdate(ByVal APPIDV As String)
        Try
            Dim LeaveAppDet As String = "Select * from ""@SBO_RESIGN"" Where ""U_Appid"" = '" & APPIDV & "'"
            Dim oRsLeaveAppDet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsLeaveAppDet.DoQuery(LeaveAppDet)
            Dim oEmp As SAPbobsCOM.EmployeesInfo = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo)
            If oEmp.GetByKey(CInt(oRsLeaveAppDet.Fields.Item("U_empID").Value)) = True Then
                Dim GetActiveID As String = "Select ""reasonID"" From OHTR Where ""name""='" & oEmp.TreminationReason & "' And ""reasonID"" IS NOT NULL"
                Dim oRsGetActiveID As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetActiveID.DoQuery(GetActiveID)

                If oRsGetActiveID.RecordCount > 0 Then

                    oEmp.TreminationReason = CStr(oEmp.UserFields.Fields.Item("U_PRSTS").Value).ToString
                    oEmp.UserFields.Fields.Item("U_PRSTS").Value = CStr(oRsGetActiveID.Fields.Item(0).Value).ToString
                    oEmp.UserFields.Fields.Item("U_STUPDT").Value = DateTime.Now.ToString

                    If oEmp.Update = 0 Then
                        objMain.objApplication.StatusBar.SetText("Status Changed to " & CStr(oEmp.UserFields.Fields.Item("U_PRSTS").Value).ToString & "  with empID '" & oRsLeaveAppDet.Fields.Item("U_empID").Value & "'   ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Else
                        objMain.objApplication.StatusBar.SetText("Failed to update employee master  with empID '" & oRsLeaveAppDet.Fields.Item("U_empID").Value & "'   -->", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If

                Else
                    objMain.objApplication.StatusBar.SetText("ReasonID not fetched  ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Try
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub


    Private Sub FinalSettlement(ByVal AppIDV As String)
        Try
            Dim DocumentDetails As String = "Select ""Code"",""U_FDate"" From ""@SBO_EOSPMSTR"" WHERE ""U_Appid"" = '" & AppIDV & "'"
            Dim oRsDocumentDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsDocumentDetails.DoQuery(DocumentDetails)

            Dim GetLastPyDet As String = "Select IFNULL((Select ADD_DAYS(MAX(""PayrollEndDate""),1)  From ""SBO_PRTRMSTRTMBACKUP_GULFAR_TEST01"" Where ""EmpId"" = " &
                                         " '" & oRsDocumentDetails.Fields.Item("Code").Value & "'), (Select ""U_LPDate""  From ""@SBO_PRCONTROL""))," &
                                        "((SELECT ""termDate"" FROM ""OHEM"" WHERE ""empID""='" & oRsDocumentDetails.Fields.Item("Code").Value & "') ) From DUMMY"
            Dim oRsGetLastPyDet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetLastPyDet.DoQuery(GetLastPyDet)

            Dim UpdatePD As String = "Update ""@SBO_PAYDED_C0"" B Set ""U_PAYSET"" = 'Y',""U_LSETID"" = '" & AppIDV & "' From ""@SBO_PAYDED"" A Inner join ""@SBO_PAYDED_C0"" B  On A.""DocEntry""=B.""DocEntry"" " &
                                    "Where B.""U_Code""='" & oRsDocumentDetails.Fields.Item("Code").Value & "' And A.""U_AppStat""='A' " &
                                    "And B.""U_APPSTAT""='A' And ifnull(""U_PAYSET"",'N') <> 'Y' " &
                                    "And ((A.""U_Month""='" & Month(oRsGetLastPyDet.Fields.Item(0).Value) & "' And A.""U_Year""='" & Year(oRsGetLastPyDet.Fields.Item(0).Value) & "') " &
                                    "Or (A.""U_Month""='" & Month(oRsGetLastPyDet.Fields.Item(1).Value) & "' And A.""U_Year""='" & Year(oRsGetLastPyDet.Fields.Item(1).Value) & "'))"
            Dim oRsUpdatePD As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsUpdatePD.DoQuery(UpdatePD)

            Dim UpdateResign As String = "Update ""@SBO_RESIGN"" SET ""U_CLID"" = '" & AppIDV & "',""U_CLDATE"" = '" & Format(CType(oRsDocumentDetails.Fields.Item("U_FDate").Value, Date), "yyyyMMdd") & "' WHERE ""U_empID"" = '" & oRsDocumentDetails.Fields.Item("Code").Value & "'"
            Dim oRsUpdateResign As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsUpdateResign.DoQuery(UpdateResign)
            Dim updateOHEM As String = "Update OHEM Set ""Active"" = 'N',""U_LeavSts"" = 'Paid',""U_STLSTAT"" = 'Yes',""U_EOSID"" = '" & AppIDV & "', ""U_EOSDT"" = '" & Format(CType(oRsDocumentDetails.Fields.Item("U_FDate").Value, Date), "yyyyMMdd") & "' Where ""empID"" = '" & CInt(oRsDocumentDetails.Fields.Item("Code").Value) & "'"
            Dim OrsUpdateOHEM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            OrsUpdateOHEM.DoQuery(updateOHEM)
            'Dim oEmploy As SAPbobsCOM.EmployeesInfo = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo)  '" & Format(CType(oRSchksettlement.Fields.Item("U_DocDate").Value, Date), "yyyyMMdd") & "'
            'If oEmploy.GetByKey(CInt(oRsDocumentDetails.Fields.Item("Code").Value)) = True Then
            '    oEmploy.UserFields.Fields.Item("U_LeavSts").Value = "Paid"
            '    'oEmploy.Active = "N"
            '    oEmploy.UserFields.Fields.Item("U_STLSTAT").Value = "Yes"
            '    oEmploy.UserFields.Fields.Item("U_EOSID").Value = AppIDV
            '    oEmploy.UserFields.Fields.Item("U_EOSDT").Value = Format(CType(oRsDocumentDetails.Fields.Item("U_FDate").Value, Date), "yyyyMMdd")
            '    If oEmploy.Update <> 0 Then
            '        objMain.objApplication.StatusBar.SetText("Failed to update employee pay status with empID '" & AppIDV & "'   -->", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    End If
            'End If
        Catch ex As Exception
            objMain.objCompany.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
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
            For i As Integer = 1 To rsGetLeaveDetails.RecordCount
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
            Next
            '---------Update query----------
            Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim Update As String = "Update ""@SBO_PRLEVAPPMSTR"" Set ""U_AppStat""='A' Where ""U_Appid""='" & AppIDV & "' "
            oRsUpdate.DoQuery(Update)
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
            Dim chksettlement As String = "Select A.""U_SEMPCD"",A.""U_LAPPID"", B.""U_Appid"", CAST(B.""U_DocDate"" as DATE) As ""U_DocDate"" ,MONTH(A.""U_LFRM"") As ""LFRMONTH"", YEAR(A.""U_LFRM"") As ""LFRYEAR"", " &
                                              "MONTH(ADD_DAYS(A.""U_LSTPRLDT"",1)) As ""PRLMONTH"", YEAR(ADD_DAYS(A.""U_LSTPRLDT"",1)) As ""PRLYEAR"" From ""@SBO_LEAVENCASH_C1"" " &
                                              "  A Inner Join ""@SBO_LEAVENCASH"" B on A.""DocEntry"" = B.""DocEntry"" Where B.""U_Appid"" = '" & AppIDV & "' "
            Dim oRSchksettlement As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRSupdatePayDed As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRSupdateLeavApp As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRSchksettlement.DoQuery(chksettlement)

            If oRSchksettlement.RecordCount > 0 Then
                For ip As Integer = 0 To oRSchksettlement.RecordCount - 1
                    If oRSchksettlement.Fields.Item("U_SEMPCD").Value <> "" Then
                        'Dim updateEmpMas As String = "Update ""OHEM"" Set ""U_STLSTAT"" = 'Yes', ""U_LeavSts"" = 'Paid' Where ""empID"" = '" & oRSchksettlement.Fields.Item("U_SEMPCD").Value & "'"
                        Dim updateLeavApp As String = "Update ""@SBO_PRLEVAPPMSTR"" Set ""U_STLSTAT"" = 'Yes', ""U_SETLEID"" = '" & oRSchksettlement.Fields.Item("U_Appid").Value & "', " &
                                                      " ""U_SETLEDT"" = '" & Format(CType(oRSchksettlement.Fields.Item("U_DocDate").Value, Date), "yyyyMMdd") & "' Where ""U_Appid"" = '" & oRSchksettlement.Fields.Item("U_LAPPID").Value & "' And" &
                                                      """U_empid"" = '" & oRSchksettlement.Fields.Item("U_SEMPCD").Value & "'"
                        Dim updatePayDed As String = "Update ""@SBO_PAYDED_C0"" A Set  A.""U_PAYSET""='Y', A.""U_LSETID"" = '" & AppIDV & "' " &
                                                     "Where A.""DocEntry"" in (Select B.""DocEntry"" from ""@SBO_PAYDED"" B Inner join ""@SBO_PAYDED_C0"" C on C.""DocEntry"" = B.""DocEntry"" " &
                                                     "Where C.""U_Code""='" & oRSchksettlement.Fields.Item("U_SEMPCD").Value & "' And C.""U_APPSTAT""='A' " &
                                                     "And ((B.""U_Month""='" & oRSchksettlement.Fields.Item("LFRMONTH").Value & "' And B.""U_Year""= " &
                                                     "'" & oRSchksettlement.Fields.Item("LFRYEAR").Value & "')  Or (B.""U_Month""= '" & oRSchksettlement.Fields.Item("PRLMONTH").Value & "' " &
                    "And B.""U_Year""='" & oRSchksettlement.Fields.Item("PRLYEAR").Value & "'))) And A.""U_Code""='" & oRSchksettlement.Fields.Item("U_SEMPCD").Value & "' And A.""U_APPSTAT""='A'"
                        oRSupdatePayDed.DoQuery(updatePayDed)
                        oRSupdateLeavApp.DoQuery(updateLeavApp)

                        Dim oEmploy As SAPbobsCOM.EmployeesInfo = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo)
                        If oEmploy.GetByKey(CInt(oRSchksettlement.Fields.Item("U_SEMPCD").Value)) = True Then
                            oEmploy.UserFields.Fields.Item("U_LeavSts").Value = "Paid"
                            'oEmploy.UserFields.Fields.Item("U_STLSTAT").Value = "Yes"
                            If oEmploy.Update <> 0 Then
                                objMain.objApplication.StatusBar.SetText("Failed to update employee pay status with empID '" & AppIDV & "'   -->", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                        End If
                        oRSchksettlement.MoveNext()
                    End If
                Next

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
        str = "SELECT T0.""U_empid"",T0.""U_empname"", ""U_LName"",T1.""U_Code"", T1.""U_ladays"",T2.""U_LeaBda"" , T2.""U_LeaBda"" AS ""BalanceLeave"", " &
              "T2.""U_LeaAda"" ""LeaveUsed"" FROM ""@SBO_PRLEVAPPMSTR"" T0 INNER JOIN ""@SBO_PRLEVAPPDET0"" T1 ON T0.""Code""=T1.""Code"" " &
              "INNER JOIN ""@SBO_PREMPSALDET2"" T2 ON T2.""Code""= T0.""U_empid"" And T2.""U_Code"" = T1.""U_Code"" WHERE T0.""U_Appid""='" & LeaveAppID & "'"
        Dim APPDATE As String = "SELECT Case IFNULL(""U_DDate"",'') when '' then CURDATE() else IFNULL(""U_DDate"",'')  End as ""docdate"" FROM ""@SBO_NOTIF"" " &
                                "where ""U_Appid""='" & RejoinAppID & "'"
        Dim rsGetLeaveDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rsGetLeaveDetails.DoQuery(str)
        Dim rsGetLeaveDetail As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rsGetLeaveDetail.DoQuery(APPDATE)

        If rsGetLeaveDetails.RecordCount > 0 Then
            Dim EmpId, LeaveCode, BalanceLeave, AppliedLeaves, LeaveName, EMPNAME, LeaveUsed As String
            EmpId = rsGetLeaveDetails.Fields.Item("U_empid").Value
            LeaveCode = rsGetLeaveDetails.Fields.Item("U_Code").Value
            BalanceLeave = rsGetLeaveDetails.Fields.Item("BalanceLeave").Value + EarlyDays
            'AppliedLeaves = rsGetLeaveDetails.Fields.Item("AppliedLeaves").Value
            LeaveUsed = rsGetLeaveDetails.Fields.Item("LeaveUsed").Value - EarlyDays
            EMPNAME = rsGetLeaveDetails.Fields.Item("U_empname").Value
            LeaveName = rsGetLeaveDetails.Fields.Item("U_LName").Value
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
            objMain.oGeneralData.SetProperty("U_Name", EMPNAME)
            objMain.oGeneralData.SetProperty("U_Date", DateTime.Now.ToString("yyyyMMdd"))
            objMain.oGeneralData.SetProperty("U_LCode", LeaveCode)
            objMain.oGeneralData.SetProperty("U_LName", LeaveName)
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

            ElseIf OffAppID <> "" And OffStatus = "N" Then
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

    'Sub ReviewUpdate(ByVal AppIDV As String)

    '    '----------Employee review Earnings--------------------------
    '    str = "SELECT * FROM ""@SBO_EMPREV"" T0  INNER JOIN ""@SBO_EMPREV_C0"" T1 ON T0.""DocEntry""=T1.""DocEntry"" " &
    '   "WHERE T0.""U_Appid"" ='" & AppIDV & "' AND IFNULL(""U_ECode"",'')<>''"
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
    '        str = "SELECT * FROM ""@SBO_EMPREV"" T0  INNER JOIN ""@SBO_EMPREV_C2"" T1 ON T0.""DocEntry""=T1.""DocEntry"" " &
    '"WHERE T0.""U_Appid"" ='" & AppIDV & "'"
    '        rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        rs.DoQuery(str)
    '        EmpId = rs.Fields.Item("U_EmpId").Value
    '        Dim DetailsDeductionQry As String = "Select ""Code"",""U_Code"",""U_Amount"" From ""@SBO_PREMPSALDET1"" Where ""Code""='" & EmpId & "'"
    '        oRsDetails.DoQuery(DetailsDeductionQry)
    '        'Dim NetPay1 As Double = 0
    '        For i As Integer = 1 To oRsDetails.RecordCount
    '            rs.MoveFirst()
    '            For j As Integer = 1 To rs.RecordCount
    '                If oRsDetails.Fields.Item("U_Code").Value = rs.Fields.Item("U_DCode").Value Then
    '                    Dim Update As String = "Update ""@SBO_PREMPSALDET1"" Set ""U_Amount""='" & rs.Fields.Item("U_Amount").Value & "' Where ""Code""='" & EmpId & "' and ""U_Code""='" & oRsDetails.Fields.Item("U_Code").Value & "'"
    '                    Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                    oRsUpdate.DoQuery(Update)
    '                    'NetPay1 = NetPay1 + rs.Fields.Item("U_Amount").Value
    '                End If
    '                rs.MoveNext()
    '            Next
    '            oRsDetails.MoveNext()
    '        Next
    '        '------------------------------------------
    '        objMain.sCmp = objMain.objCompany.GetCompanyService()
    '        objMain.oGeneralService = objMain.sCmp.GetGeneralService("SBO_EMPSALUDO")
    '        objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
    '        objMain.oGeneralParams.SetProperty("Code", EmpId)
    '        objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)

    '        Dim oRsDedcution As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        Dim Deductions As String = "SELECT IFNULL(SUM(""U_Amount""),0) From ""@SBO_PREMPSALDET1"" Where ""Code""='" & EmpId & "' "
    '        oRsDedcution.DoQuery(Deductions)
    '        Dim Deduction As Double = CDbl(oRsDedcution.Fields.Item(0).Value)
    '        '----------Benfits------------------------------------------------
    '        str = "SELECT * FROM ""@SBO_EMPREV"" T0  INNER JOIN ""@SBO_EMPREV_C3"" T1 ON T0.""DocEntry""=T1.""DocEntry"" " &
    '             "WHERE T0.""U_Appid"" ='" & AppIDV & "'"
    '        rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        rs.DoQuery(str)
    '        EmpId = rs.Fields.Item("U_EmpId").Value
    '        Dim DetailsBenQry As String = "Select ""Code"",""U_Code"",""U_Amount"" From ""@SBO_PREMPSALDET6"" Where ""Code""='" & EmpId & "'"
    '        oRsDetails.DoQuery(DetailsBenQry)
    '        'Dim NetPay2 As Double = 0
    '        For i As Integer = 1 To oRsDetails.RecordCount
    '            rs.MoveFirst()
    '            For j As Integer = 1 To rs.RecordCount
    '                If oRsDetails.Fields.Item("U_Code").Value = rs.Fields.Item("Code").Value Then
    '                    Dim Update As String = "Update ""@SBO_PREMPSALDET6"" Set ""U_Amount""='" & rs.Fields.Item("U_Amount").Value & "' Where ""Code""='" & EmpId & "' and ""U_Code""='" & oRsDetails.Fields.Item("U_Code").Value & "'"
    '                    Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                    oRsUpdate.DoQuery(Update)
    '                    'NetPay2 = NetPay2 + rs.Fields.Item("U_Amount").Value
    '                End If
    '                rs.MoveNext()
    '            Next
    '            oRsDetails.MoveNext()
    '        Next
    '        '----------------------------------------------------------------
    '        objMain.oGeneralData.SetProperty("U_NetPay", NetPay - Deduction)
    '        objMain.oGeneralData.SetProperty("U_EmpCtc", NetPay * 12)
    '        objMain.oGeneralData.SetProperty("U_EarTot", NetPay)
    '        objMain.oGeneralService.Update(objMain.oGeneralData)

    '        '-----Updating Employee Master data
    '        rs.MoveFirst()
    '        Try
    '            Dim oEmployee As SAPbobsCOM.EmployeesInfo
    '            oEmployee = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo)
    '            If (oEmployee.GetByKey(EmpId)) = True Then
    '                oEmployee.Position = rs.Fields.Item("U_PropDes").Value
    '                oEmployee.UserFields.Fields.Item("U_RDate").Value = CDate(rs.Fields.Item("U_DocDt").Value)
    '                oEmployee.UserFields.Fields.Item("U_NtRevDt").Value = CDate(rs.Fields.Item("U_NtRevDt").Value)
    '                oEmployee.Update()
    '            End If
    '        Catch ex As Exception
    '        End Try
    '    End If
    'End Sub

    Sub ReviewUpdate(ByVal AppIDV As String)
        objMain.sCmp = objMain.objCompany.GetCompanyService()
        '----------Employee review Earnings--------------------------
        'objMain.sCmp = objMain.objCompany.GetCompanyService
        'objMain.oGeneralService = objMain.sCmp.GetGeneralService("TNX_PORDER_UDO")
        str = "SELECT * FROM ""@SBO_EMPREV"" T0  INNER JOIN ""@SBO_EMPREV_C0"" T1 ON T0.""DocEntry""=T1.""DocEntry"" " &
       "WHERE T0.""U_Appid"" ='" & AppIDV & "' AND IFNULL(""U_ECode"",'')<>''"
        Dim ors As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ors.DoQuery(str)
        Dim EmpID As String = ors.Fields.Item("U_EmpId").Value
        Dim Earning As Double = 0
        Dim Deduction As Double = 0
        If ors.RecordCount > 0 Then
            objMain.oGeneralService = objMain.sCmp.GetGeneralService("SBO_EMPSALUDO")
            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            objMain.oGeneralParams.SetProperty("Code", EmpID)
            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)

            If ors.Fields.Item("U_PCat").Value <> "" Then
                Dim CatValue As String = "Select ""Name"" from ""@SBO_CAT"" WHERE ""Code"" = '" & ors.Fields.Item("U_PCat").Value & "'"
                Dim oRsCatValue As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsCatValue.DoQuery(CatValue)
                objMain.oGeneralData.SetProperty("U_Cat", oRsCatValue.Fields.Item("Name").Value)
            End If
            If ors.Fields.Item("U_PGRADE").Value <> "" Then
                Dim GradValue As String = "Select ""Name"" from ""@SBO_GRADE"" WHERE ""Code"" = '" & ors.Fields.Item("U_PGRADE").Value & "'"
                Dim oRsGradValue As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGradValue.DoQuery(GradValue)
                objMain.oGeneralData.SetProperty("U_GRADE", oRsGradValue.Fields.Item("Name").Value)
            End If
            'If ors.Fields.Item("U_PropDes").Value <> "" Then
            '    Dim PositionValue As String = "Select ""descriptio"" From OHPS Where ""posID"" = '" & ors.Fields.Item("U_PropDes").Value & "'"
            '    Dim oRsPositionValue As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '    oRsPositionValue.DoQuery(PositionValue)
            '    objMain.oGeneralData.SetProperty("U_DESGN", oRsPositionValue.Fields.Item("descriptio").Value)
            'End If

            objMain.oChildren = objMain.oGeneralData.Child("SBO_PREMPSALDET0")
            For j As Integer = 1 To ors.RecordCount
                Dim EarningCode As String = ors.Fields.Item("U_ECode").Value
                Earning = Earning + ors.Fields.Item("U_NewSal").Value
                Dim oRsDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim DetailsQry As String = "Select ""LineId"" From ""@SBO_PREMPSALDET0"" " &
                    "Where ""Code""='" & EmpID & "' and ""U_Code""='" & EarningCode & "'"
                oRsDetails.DoQuery(DetailsQry)
                If oRsDetails.RecordCount > 0 Then
                    Dim Amount As Double = ors.Fields.Item("U_NewSal").Value
                    Dim LineID As Integer = oRsDetails.Fields.Item("LineId").Value
                    Dim Status As String = ors.Fields.Item("U_Status").Value
                    Dim EffDate As String = ors.Fields.Item("U_NRevDt").Value
                    Dim Appid As String = ors.Fields.Item("U_Appid").Value
                    objMain.oChild = objMain.oChildren.Item(LineID - 1)
                    If ors.Fields.Item("U_Diff").Value <> 0 Then
                        objMain.oChild.SetProperty("U_Amount", Amount)
                        objMain.oChild.SetProperty("U_INCRAID", Appid)
                        objMain.oChild.SetProperty("U_EFFDATE", EffDate)

                    End If
                    If Status = "" Then
                        objMain.oChild.SetProperty("U_Status", "1")
                    Else
                        objMain.oChild.SetProperty("U_Status", Status)
                    End If
                Else
                    objMain.oChild = objMain.oChildren.Add()
                    objMain.oChild.SetProperty("U_Code", ors.Fields.Item("U_ECode").Value)
                    objMain.oChild.SetProperty("U_dscr", ors.Fields.Item("U_EDes").Value)
                    objMain.oChild.SetProperty("U_GLCode", ors.Fields.Item("U_EGL").Value)
                    objMain.oChild.SetProperty("U_Amount", ors.Fields.Item("U_NewSal").Value)
                    If ors.Fields.Item("U_Status").Value = "" Then
                        objMain.oChild.SetProperty("U_Status", "1")
                    Else
                        objMain.oChild.SetProperty("U_Status", ors.Fields.Item("U_Status").Value)
                    End If
                    objMain.oChild.SetProperty("U_Pay", "Y")
                End If
                ors.MoveNext()
            Next
            objMain.oGeneralService.Update(objMain.oGeneralData)
        End If

        '----------Employee review Deductions--------------------------
        Dim Dedstr As String = "SELECT * FROM ""@SBO_EMPREV"" T0  INNER JOIN ""@SBO_EMPREV_C2"" T1 ON T0.""DocEntry""=T1.""DocEntry"" " &
       "WHERE T0.""U_Appid"" ='" & AppIDV & "' AND IFNULL(""U_DCode"",'')<>'' AND IFNULL(""U_Status"",'')<>''"
        rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs.DoQuery(Dedstr)
        If rs.RecordCount > 0 Then
            objMain.oGeneralService = objMain.sCmp.GetGeneralService("SBO_EMPSALUDO")
            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            objMain.oGeneralParams.SetProperty("Code", EmpID)
            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
            objMain.oChildren = objMain.oGeneralData.Child("SBO_PREMPSALDET1")
            For j As Integer = 1 To rs.RecordCount
                Dim DeductionCode As String = rs.Fields.Item("U_DCode").Value
                Deduction = Deduction + rs.Fields.Item("U_NewDed").Value
                Dim oRsDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim DetailsQry As String = "Select ""LineId"" From ""@SBO_PREMPSALDET1"" " &
                    "Where ""Code""='" & EmpID & "' and ""U_Code""='" & DeductionCode & "'"
                oRsDetails.DoQuery(DetailsQry)
                If oRsDetails.RecordCount > 0 Then
                    Dim Amount As Double = rs.Fields.Item("U_NewDed").Value
                    Dim LineID As Integer = oRsDetails.Fields.Item("LineId").Value
                    Dim Status As String = rs.Fields.Item("U_Status").Value
                    objMain.oChild = objMain.oChildren.Item(LineID - 1)
                    objMain.oChild.SetProperty("U_Amount", Amount)
                    If Status = "" Then
                        objMain.oChild.SetProperty("U_Status", "1")
                    Else
                        objMain.oChild.SetProperty("U_Status", Status)
                    End If
                Else
                    objMain.oChild = objMain.oChildren.Add()
                    objMain.oChild.SetProperty("U_Code", rs.Fields.Item("U_DCode").Value)
                    objMain.oChild.SetProperty("U_dscr", rs.Fields.Item("U_DDes").Value)
                    objMain.oChild.SetProperty("U_GLCode", rs.Fields.Item("U_GLcode").Value)
                    objMain.oChild.SetProperty("U_Amount", rs.Fields.Item("U_NewDed").Value)
                    If rs.Fields.Item("U_Status").Value = "" Then
                        objMain.oChild.SetProperty("U_Status", "1")
                    Else
                        objMain.oChild.SetProperty("U_Status", rs.Fields.Item("U_Status").Value)
                    End If

                    objMain.oChild.SetProperty("U_Pay", "Y")
                End If
                rs.MoveNext()
            Next
            objMain.oGeneralService.Update(objMain.oGeneralData)
        End If


        '----------Employee review Benefits--------------------------
        Dim Benstr As String = "SELECT * FROM ""@SBO_EMPREV"" T0  INNER JOIN ""@SBO_EMPREV_C3"" T1 ON T0.""DocEntry""=T1.""DocEntry"" " &
       "WHERE T0.""U_Appid"" ='" & AppIDV & "' AND IFNULL(""Code"",'')<>'' AND IFNULL(""U_Status"",'')<>''"
        rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs.DoQuery(Benstr)
        If rs.RecordCount > 0 Then
            objMain.oGeneralService = objMain.sCmp.GetGeneralService("SBO_EMPSALUDO")
            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            objMain.oGeneralParams.SetProperty("Code", EmpID)
            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
            objMain.oChildren = objMain.oGeneralData.Child("SBO_PREMPSALDET6")
            For j As Integer = 1 To rs.RecordCount
                Dim DeductionCode As String = rs.Fields.Item("Code").Value
                Dim oRsDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim DetailsQry As String = "Select ""LineId"" From ""@SBO_PREMPSALDET6"" " &
                    "Where ""Code""='" & EmpID & "' and ""U_Code""='" & DeductionCode & "'"
                oRsDetails.DoQuery(DetailsQry)
                If oRsDetails.RecordCount > 0 Then
                    Dim Amount As Double = rs.Fields.Item("U_NewBen").Value
                    Dim LineID As Integer = oRsDetails.Fields.Item("LineId").Value
                    Dim Status As String = rs.Fields.Item("U_Status").Value
                    objMain.oChild = objMain.oChildren.Item(LineID - 1)
                    objMain.oChild.SetProperty("U_Amount", Amount)
                    objMain.oChild.SetProperty("U_EffDate", rs.Fields.Item("U_EffDate").Value)
                    objMain.oChild.SetProperty("U_EndDate", rs.Fields.Item("U_EDate").Value)
                    If Status = "" Then
                        objMain.oChild.SetProperty("U_BenStu", "1")
                    Else
                        objMain.oChild.SetProperty("U_BenStu", Status)
                    End If
                Else
                    objMain.oChild = objMain.oChildren.Add()
                    objMain.oChild.SetProperty("U_Code", rs.Fields.Item("Code").Value)
                    objMain.oChild.SetProperty("U_dscr", rs.Fields.Item("U_BDes").Value)
                    objMain.oChild.SetProperty("U_GLCode", rs.Fields.Item("U_GLAcnt").Value)
                    objMain.oChild.SetProperty("U_Amount", rs.Fields.Item("U_NewBen").Value)
                    objMain.oChild.SetProperty("U_EffDate", rs.Fields.Item("U_EffDate").Value)
                    objMain.oChild.SetProperty("U_EndDate", rs.Fields.Item("U_EDate").Value)
                    If rs.Fields.Item("U_Status").Value = "" Then
                        objMain.oChild.SetProperty("U_BenStu", "1")
                    Else
                        objMain.oChild.SetProperty("U_BenStu", rs.Fields.Item("U_Status").Value)
                    End If
                End If
                rs.MoveNext()
            Next
            'oGeneralService.Update(oGeneralData)
        End If
        objMain.oGeneralData.SetProperty("U_NetPay", Earning - Deduction)
        objMain.oGeneralData.SetProperty("U_EmpCtc", Earning * 12)
        objMain.oGeneralData.SetProperty("U_EarTot", Earning)
        objMain.oGeneralData.SetProperty("U_DedTot", Deduction)
        objMain.oGeneralService.Update(objMain.oGeneralData)
        ors.MoveFirst()


        '-----Updating Employee Master data
        'Try
        '    Dim oEmployee As SAPbobsCOM.EmployeesInfo
        '    oEmployee = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo)
        '    If (oEmployee.GetByKey(EmpID)) = True Then
        '        oEmployee.UserFields.Fields.Item("U_RDate").Value = ors.Fields.Item("U_DocDt").Value
        '        oEmployee.UserFields.Fields.Item("U_NtRevDt").Value = ors.Fields.Item("U_NtRevDt").Value
        '        If ors.Fields.Item("U_PropDes").Value <> "" Then
        '            oEmployee.Position = ors.Fields.Item("U_PropDes").Value
        '            Dim PosDesc As String = "Select ""descriptio"" from OHPS Where ""posID"" = '" & ors.Fields.Item("U_PropDes").Value & "'"
        '            Dim oRsPosDesc As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '            oRsPosDesc.DoQuery(PosDesc)
        '            oEmployee.UserFields.Fields.Item("U_FULLDSC").Value = oRsPosDesc.Fields.Item("descriptio").Value

        '        End If
        '        If ors.Fields.Item("U_PGRADE").Value <> "" Then
        '            oEmployee.UserFields.Fields.Item("U_Grade").Value = ors.Fields.Item("U_PGRADE").Value
        '        End If
        '        If ors.Fields.Item("U_PFStatus").Value <> "" Then
        '            oEmployee.UserFields.Fields.Item("U_FStatus").Value = ors.Fields.Item("U_PFStatus").Value
        '        End If
        '        If ors.Fields.Item("U_PFInsur").Value <> "" Then
        '            oEmployee.UserFields.Fields.Item("U_FInsur").Value = ors.Fields.Item("U_PFInsur").Value
        '        End If
        '        If ors.Fields.Item("U_PCat").Value <> "" Then
        '            oEmployee.UserFields.Fields.Item("U_Cat").Value = ors.Fields.Item("U_PCat").Value
        '        End If
        '        If ors.Fields.Item("U_PEType").Value <> "" Then
        '            oEmployee.UserFields.Fields.Item("U_EType").Value = ors.Fields.Item("U_PEType").Value
        '        End If
        '        If ors.Fields.Item("U_PDept").Value <> "" Then
        '            oEmployee.Department = ors.Fields.Item("U_PDept").Value
        '        End If
        '        oEmployee.Update()
        '    End If
        'Catch ex As Exception
        'End Try
        Try
            Dim oEmployee As SAPbobsCOM.EmployeesInfo
            oEmployee = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo)

            If (oEmployee.GetByKey(EmpID)) = True Then

                '--- Existing fields ---
                oEmployee.UserFields.Fields.Item("U_RDate").Value = ors.Fields.Item("U_DocDt").Value
                oEmployee.UserFields.Fields.Item("U_NtRevDt").Value = ors.Fields.Item("U_NtRevDt").Value

                Dim q As String = "SELECT ""U_PropDes"" FROM ""@SBO_EMPREV"" WHERE ""U_EmpId"" = '" & EmpID & "' and ""U_Appid"" ='" & AppIDV & "'"
                Dim oRsFull As SAPbobsCOM.Recordset
                oRsFull = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsFull.DoQuery(q)

                Dim proposedName As String = ""

                If oRsFull.RecordCount > 0 Then
                    proposedName = oRsFull.Fields.Item("U_PropDes").Value
                End If



                If proposedName <> "" Then

                    Dim qPos As String = "SELECT ""posID"" FROM OHPS WHERE UPPER(""descriptio"") = UPPER('" & proposedName.Replace("'", "''") & "')"
                    Dim oRsPos As SAPbobsCOM.Recordset
                    oRsPos = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsPos.DoQuery(qPos)

                    If oRsPos.RecordCount > 0 Then

                        Dim posId As Integer = oRsPos.Fields.Item("posID").Value
                        ' After setting position:
                        oEmployee.UserFields.Fields.Item("U_FULLDSC").Value = proposedName


                        oEmployee.Position = posId

                    End If

                End If


                If ors.Fields.Item("U_PGRADE").Value <> "" Then oEmployee.UserFields.Fields.Item("U_Grade").Value = ors.Fields.Item("U_PGRADE").Value
                If ors.Fields.Item("U_PFStatus").Value <> "" Then oEmployee.UserFields.Fields.Item("U_FStatus").Value = ors.Fields.Item("U_PFStatus").Value
                If ors.Fields.Item("U_PFInsur").Value <> "" Then oEmployee.UserFields.Fields.Item("U_FInsur").Value = ors.Fields.Item("U_PFInsur").Value
                If ors.Fields.Item("U_PCat").Value <> "" Then oEmployee.UserFields.Fields.Item("U_Cat").Value = ors.Fields.Item("U_PCat").Value
                If ors.Fields.Item("U_PEType").Value <> "" Then oEmployee.UserFields.Fields.Item("U_EType").Value = ors.Fields.Item("U_PEType").Value
                If ors.Fields.Item("U_PDept").Value <> "" Then oEmployee.Department = ors.Fields.Item("U_PDept").Value



                Dim updateResult As Integer = oEmployee.Update()

                If updateResult <> 0 Then
                    '    objMain.objApplication.StatusBar.SetText("Employee update failed: " &
                    'objMain.objCompany.GetLastErrorDescription(),
                    'SAPbouiCOM.BoMessageTime.bmt_Short,
                    'SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    objMain.objApplication.StatusBar.SetText("Employee Master Updated (Position + FullDesignation) Successfully",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If

            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try


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




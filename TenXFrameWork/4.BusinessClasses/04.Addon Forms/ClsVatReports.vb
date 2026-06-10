

Imports SAPbouiCOM
Imports SAPbobsCOM
Imports System.IO
Imports System.Text
Imports System.Diagnostics
Imports System.Runtime.InteropServices

Public Class ClsVatReports

#Region "Declaration"

    Public objForm As SAPbouiCOM.Form
    Public objMatrix, objMatrix1, objMatrix2, objMatrix3 As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Details As SAPbouiCOM.DBDataSource
    Dim oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim oDBs_Attach As SAPbouiCOM.DBDataSource
    Dim oDS As SAPbouiCOM.DBDataSource


#End Region

#Region "Create Form"

    Public Sub CreateForm()

        Try
            objMain.objUtilities.LoadForm("VATReport.xml", "VATR", ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm("VATR",
                      objMain.objApplication.Forms.ActiveForm.TypeCount)


            objForm.Freeze(True)
            Dim oStatic As SAPbouiCOM.StaticText
            objForm.Items.Item("vaturl").TextStyle = 4
            oStatic = CType(objForm.Items.Item("vaturl").Specific, SAPbouiCOM.StaticText)

            Dim Vaturl As String = oStatic.Caption

            oStatic.Caption = Vaturl
            'Dim oStatic As SAPbouiCOM.StaticText

            'oStatic = CType(objForm.Items.Item("vaturl").Specific, SAPbouiCOM.StaticText)

            'Dim sText As String = oStatic.Caption

            'oStatic.Caption = "<html><font color='blue'><u>" & sText & "</u></font></html>"

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VATRP")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_VATRP_C0")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_VATCTM_C1")
            oDBs_Attach = objForm.DataSources.DBDataSources.Item("@TNX_ATTACH_C3")
            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNXVATRUDO", "Primary"))
            'oDBs_Head.SetValue("U_DA", 0, DateTime.Now.ToString("yyyyMMdd"))
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            ' objForm.Items.Item("DA").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            '  LoadCompanyDetails(objForm.UniqueID)
            ' LoadVATReportLines(objForm.UniqueID, "2026-01-01", "2026-03-31")
            'SetVATTabs(objForm)
            objForm.PaneLevel = 1
            'Disable in all modes
            objForm.Items.Item("VRPY").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'Enable only in Add Mode
            objForm.Items.Item("VRPY").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("VATTO").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("VATTO").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("SD").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("SD").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            ' objForm.Items.Item("APPI").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Item_10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("TRNM").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("TRNM").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("2").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("FTY").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("TPNE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("TPNE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("TPNA").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("TPNA").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("TPA").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("TPA").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("APPI").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("APPI").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("NVAT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("NVAT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("APPI").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("APPI").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("Item_11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            SetExportButtonByMode(objForm)
            LoadVATHeaderOnly(objForm.UniqueID)
            Me.objForm.EnableMenu("1282", True)
            Me.objForm.EnableMenu("519", True)
            Me.objForm.EnableMenu("520", True)
            'SetDefault(objForm.UniqueID)
            Me.objForm.EnableMenu("1292", True)
            SetDefault(objForm.UniqueID)
            ' Me.SetNewLine(objForm.UniqueID)
            Me.objForm.EnableMenu("1293", True)
            objForm.Items.Item("Item_5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("APPI").Enabled = False
            objForm.Items.Item("Item_7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(
            "VAT Report Form Loaded Successfully",
            BoMessageTime.bmt_Short,
            BoStatusBarMessageType.smt_Success)

        Catch ex As Exception

            Try
                objForm.Freeze(False)
            Catch
            End Try

            objMain.objApplication.StatusBar.SetText(
            ex.Message,
            BoMessageTime.bmt_Short,
            BoStatusBarMessageType.smt_Warning)

        End Try

    End Sub

#End Region

#Region "Default Values"

    Public Sub SetDefault(ByVal FormUID As String)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VATRP")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_VATRP_C0")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_VATCTM_C1")
            oDBs_Attach = objForm.DataSources.DBDataSources.Item("@TNX_ATTACH_C3")
            objForm.Freeze(True)
            'Dim oStatic As SAPbouiCOM.StaticText

            'oStatic = CType(objForm.Items.Item("vaturl").Specific, SAPbouiCOM.StaticText)

            'Dim sText As String = oStatic.Caption

            'oStatic.Caption = "<html><font color='blue'><u>" & sText & "</u></font></html>"
            ' If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            AutoDocentryNumber(objForm.UniqueID)
            oDBs_Head.SetValue("U_FTY", 0, "Open")
            '    AddAttachmentNewLine(FormUID)
            'End If
            '  Me.CreateForm()

            objForm.PaneLevel = 1
            LoadVATHeaderOnly(objForm.UniqueID)
            SetExportButtonByMode(objForm)
            SetNewLine(FormUID)
            '  AddAttachmentNewLine(FormUID)
            objForm.Items.Item("APPI").Enabled = False
            objForm.Freeze(False)

        Catch ex As Exception

            Try
                objForm.Freeze(False)
            Catch
            End Try

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub


#End Region

#Region "Add New Line"

    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VATRP")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_VATRP_C0")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_VATCTM_C1")
            oDBs_Attach = objForm.DataSources.DBDataSources.Item("@TNX_ATTACH_C3")
            objMatrix = objForm.Items.Item("MXT_3").Specific


            objMatrix.AddRow()
            oDBs_Attach.SetValue("LineId", oDBs_Attach.Offset, objMatrix.VisualRowCount)
            oDBs_Attach.SetValue("U_TPA", oDBs_Attach.Offset, "")   'Target Path
            oDBs_Attach.SetValue("U_FN", oDBs_Attach.Offset, "")    'File Name
            oDBs_Attach.SetValue("U_FTT", oDBs_Attach.Offset, "")   'Free Text
            oDBs_Attach.SetValue("U_ATD", oDBs_Attach.Offset, "") 'Attachment Date
            objMatrix.SetLineData(objMatrix.VisualRowCount)

            objMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


#End Region



#Region "Menu Event"
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "VATR" And pVal.BeforeAction = False Then
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                objMatrix3 = objForm.Items.Item("MXT_3").Specific
                Me.SetDefault(objForm.UniqueID)
                objForm.Items.Item("APPI").Enabled = False

                objForm.Items.Item("Item_5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                '  Me.SetDefault(objForm.UniqueID)
                Me.SetNewLine(objForm.UniqueID)

                objMatrix = objForm.Items.Item("MXT_3").Specific
                objForm.Items.Item("APPI").Enabled = False
                ' Me.SetNewLine(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1281" And pVal.BeforeAction = False Then

            ElseIf pVal.MenuUID = "1293" AndAlso pVal.BeforeAction = True Then

                Try

                    objForm = objMain.objApplication.Forms.ActiveForm

                    If objForm.TypeEx <> "VATR" Then Exit Sub

                    BubbleEvent = False

                    objForm.Freeze(True)

                    objMatrix = CType(objForm.Items.Item("MXT_3").Specific,
                                       SAPbouiCOM.Matrix)

                    oDBs_Attach = objForm.DataSources.DBDataSources.Item("@TNX_ATTACH_C3")

                    Dim selectedRow As Integer =
                    objMatrix.GetNextSelectedRow(0,
                    SAPbouiCOM.BoOrderType.ot_RowOrder)

                    If selectedRow <= 0 Then

                        objMain.objApplication.StatusBar.SetText(
                        "Please select row to delete",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        Exit Try

                    End If

                    'Delete selected row
                    objMatrix.DeleteRow(selectedRow)

                    'Sync Matrix to DB
                    objMatrix.FlushToDataSource()

                    'Remove extra datasource rows
                    While oDBs_Attach.Size > objMatrix.VisualRowCount
                        oDBs_Attach.RemoveRecord(oDBs_Attach.Size - 1)
                    End While

                    'Keep minimum one row
                    If oDBs_Attach.Size = 0 Then

                        oDBs_Attach.InsertRecord(0)

                        oDBs_Attach.SetValue("LineId", 0, "1")
                        oDBs_Attach.SetValue("U_TPA", 0, "")
                        oDBs_Attach.SetValue("U_FN", 0, "")
                        oDBs_Attach.SetValue("U_ATD", 0, "")
                        oDBs_Attach.SetValue("U_FTT", 0, "")

                    End If

                    'Re-sequence line numbers
                    For i As Integer = 0 To oDBs_Attach.Size - 1

                        oDBs_Attach.SetValue("LineId",
                                             i,
                                             (i + 1).ToString())

                    Next

                    objMatrix.LoadFromDataSource()
                    objMatrix.AutoResizeColumns()

                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If

                Catch ex As Exception

                    objMain.objApplication.StatusBar.SetText(
                    "Delete Row Error : " & ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Finally

                    Try
                        objForm.Freeze(False)
                    Catch
                    End Try

                End Try
            ElseIf pVal.MenuUID = "519" AndAlso pVal.BeforeAction = True Then

                Try

                    objForm = objMain.objApplication.Forms.ActiveForm

                    If objForm.TypeEx <> "VATR" Then Exit Sub

                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VATRP")

                    Dim frs As SAPbobsCOM.Recordset =
           objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'VATREPORT' AND ""Type"" = 'C'")

                    If frs.RecordCount = 0 Then

                        objMain.objApplication.MessageBox("VATREPORT Layout not found. Please import Crystal Layout with same name.", 0, "OK")

                    Else

                        objMain.objApplication.ActivateMenuItem(
               frs.Fields.Item(0).Value.ToString())

                        Dim CrForm As SAPbouiCOM.Form
                        Dim oedt As SAPbouiCOM.EditText

                        CrForm = objMain.objApplication.Forms.ActiveForm

                        oedt = CrForm.Items.Item("1000003").Specific

                        oedt.Value = oDBs_Head.GetValue("DocEntry", 0)

                        CrForm.Items.Item("1").Click(
               SAPbouiCOM.BoCellClickType.ct_Regular)

                    End If

                Catch ex As Exception

                    objMain.objApplication.StatusBar.SetText("Preview Error : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                End Try



            ElseIf pVal.MenuUID = "520" AndAlso pVal.BeforeAction = True Then

                Try

                    Dim LayoutSelection As Integer =
           objMain.objApplication.MessageBox("Please select layout for printing", 1, "VAT Report", "")

                    If LayoutSelection = 1 Then

                        objForm = objMain.objApplication.Forms.ActiveForm

                        If objForm.TypeEx <> "VATR" Then Exit Sub

                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VATRP")

                        Dim frs As SAPbobsCOM.Recordset =
               objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'VATREPORT' AND ""Type"" = 'C'")

                        If frs.RecordCount = 0 Then

                            objMain.objApplication.MessageBox("VATREPORT Layout not found. Please import Crystal Layout with same name.", 0, "OK")

                        Else

                            objMain.objApplication.ActivateMenuItem(frs.Fields.Item(0).Value.ToString())

                            Dim CrForm As SAPbouiCOM.Form
                            Dim oedt As SAPbouiCOM.EditText

                            CrForm = objMain.objApplication.Forms.ActiveForm

                            oedt = CrForm.Items.Item("1000003").Specific

                            oedt.Value = oDBs_Head.GetValue("DocEntry", 0)

                            CrForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        End If

                    End If

                Catch ex As Exception

                    objMain.objApplication.StatusBar.SetText("Print Error : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                End Try

            End If

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


#End Region
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("VATR", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        'Dim oStatic As SAPbouiCOM.StaticText

                        'oStatic = CType(objForm.Items.Item("vaturl").Specific, SAPbouiCOM.StaticText)

                        'Dim sText As String = oStatic.Caption

                        'oStatic.Caption = "<html><font color='blue'><u>" & sText & "</u></font></html>"

                        Dim ApprovalStatus As String = ""

                        Try
                            ApprovalStatus = oDBs_Head.GetValue("U_FTY", oDBs_Head.Offset).Trim()
                        Catch
                            ApprovalStatus = ""
                        End Try
                        If ApprovalStatus = "A" _
    OrElse ApprovalStatus = "Approved" OrElse ApprovalStatus = "S" OrElse ApprovalStatus = "N" Then

                            objForm.Items.Item("Item_5").Enabled = False

                        Else

                            objForm.Items.Item("Item_5").Enabled = True
                        End If
                        If ApprovalStatus = "O" _
OrElse ApprovalStatus = "Open" Then

                            objForm.Items.Item("Item_8").Enabled = False
                            objForm.Items.Item("Item_11").Enabled = False

                        Else

                            objForm.Items.Item("Item_11").Enabled = True
                            objForm.Items.Item("Item_8").Enabled = True
                        End If
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub ApprovalTrigger(ByVal DocType As String, ByVal AppID As String, ByVal Table As String, ByVal AppIDField As String, ByVal AppStatField As String, ByVal Document As String)
        Dim ff As Integer
        Dim draftquery As String
        Dim approverid As String
        Dim docdate As String = objForm.Items.Item("SD").Specific.Value
        Dim rsapp As SAPbobsCOM.Recordset
        Dim posrs As SAPbobsCOM.Recordset

        ff = objMain.objApplication.MessageBox("Do You want Submit the Application?", 1, "Ok", "Cancel")
        If ff = 1 Then
            Dim result As Integer
            result = objMain.objApplication.MessageBox("You can not change the Application after you Sending for Approval,continue", 1, "Yes", "No")
            If result = 1 Then
                rsapp = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim str As String = "SELECT T0.""Code"" As ""TemplateID"",T1.""U_Name"" As ""Originator"",S1.""Code"" As ""Stage"",S2.""U_UKey"" As ""Authorizers"" " &
         " ,S1.""U_NAP"" As ""NoOfAuth"",S1.""U_NRJ"" As ""NoOfRej"",T0.""CreateDate"",(Case When T3.""LineId""='1' Then 'S' Else 'O' End) As ""Status"" FROM ""@SBO_APPHDR"" T0 " &
         " INNER JOIN ""@SBO_APPREQ"" T1 ON T1.""Code""=T0.""Code"" " &
         " INNER JOIN ""@SBO_APPDOC"" T2 ON T2.""Code"" = T0.""Code"" " &
         " INNER JOIN ""@SBO_APPAUT"" T3 ON T3.""Code"" = T0.""Code"" " &
         " INNER JOIN ""@SBO_AST"" S1 ON T3.""U_M3_1""=S1.""Code"" " &
         " INNER JOIN ""@SBO_AST_C0"" S2 ON S1.""Code""=S2.""Code"" " &
         " WHERE ""U_Active"" = 'Y' AND T1.""U_Name"" = '" & objMain.objCompany.UserName & "' AND T2.""U_" & DocType & """  = 'Y' AND IFNULL(S2.""U_UKey"",'')<>''  order by T3.""LineId""  ;"
                rsapp.DoQuery(str)

                If rsapp.RecordCount > 0 Then
                    posrs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    draftquery = " SELECT * FROM ""@SBO_OWDD"" WHERE ""U_statusN"" IN ('S','A') AND ""U_Isdraft"" IN ('Y','N') AND ""U_docnum"" = '" & AppID & "';"
                    posrs.DoQuery(draftquery)
                    If posrs.RecordCount > 0 Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(posrs)
                        objMain.objApplication.StatusBar.SetText("Document Already sent for approval", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    Else
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(posrs)

                        objMain.ObjDraftProcedure.CreateForm(objForm, DocType, AppID, objMain.objCompany.UserName, rsapp, Table, AppIDField, AppStatField, Document)

                    End If
                Else
                    objMain.objApplication.StatusBar.SetText("Login User Don't have Approvals Please Define Approvals and Add the Document", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        End If
    End Sub


    Public Sub AutoDocentryNumber(ByVal FormUID As String)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VATRP")

            If oDBs_Head.Size = 0 Then
                oDBs_Head.InsertRecord(0)
            End If

            Dim oRsDocNum As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            Dim Query1 As String = "SELECT IFNULL(MAX(""DocNum""),0)+1 AS ""DocNum"" FROM ""@TNX_VATRP"""

            oRsDocNum.DoQuery(Query1)

            oDBs_Head.SetValue("DocNum", 0, oRsDocNum.Fields.Item("DocNum").Value.ToString())


            Dim rsAppId As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            Dim str As String = "SELECT 'V' || LPAD(" & "TO_NVARCHAR(IFNULL(MAX(TO_INTEGER(REPLACE(""U_APPI"", 'V', ''))),0)+1), 6, '0') AS ""AppId"" " & "FROM ""@TNX_VATRP"""

            rsAppId.DoQuery(str)

            oDBs_Head.SetValue("U_APPI", 0, rsAppId.Fields.Item("AppId").Value.ToString())

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("AutoDocentryNumber Error : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
    Private Function ValidateVATInputDates(ByVal oForm As SAPbouiCOM.Form) As Boolean

        Try

            Dim FromDate As String = GetEditValue(oForm, "VRPY")
            Dim ToDate As String = GetEditValue(oForm, "VATTO")

            If FromDate = "" OrElse ToDate = "" Then

                objMain.objApplication.StatusBar.SetText(
            "Please enter VAT Return Period From and To Date",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                Return False

            End If

            Dim dtFrom As DateTime
            Dim dtTo As DateTime

            If Not DateTime.TryParseExact(
            FromDate,
            "yyyyMMdd",
            System.Globalization.CultureInfo.InvariantCulture,
            Globalization.DateTimeStyles.None,
            dtFrom) Then

                objMain.objApplication.StatusBar.SetText(
            "Invalid From Date",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Return False

            End If

            If Not DateTime.TryParseExact(
            ToDate,
            "yyyyMMdd",
            System.Globalization.CultureInfo.InvariantCulture,
            Globalization.DateTimeStyles.None,
            dtTo) Then

                objMain.objApplication.StatusBar.SetText(
            "Invalid To Date",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Return False

            End If

            If dtFrom > dtTo Then

                objMain.objApplication.StatusBar.SetText(
            "From Date should not be greater than To Date",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Return False

            End If

            Return True

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
        "Date Validation Error : " & ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            Return False

        End Try

    End Function


#Region "Item Event"
    Public Sub ItemEvent(ByVal FormUID As String,
                 ByRef pVal As SAPbouiCOM.ItemEvent,
                 ByRef BubbleEvent As Boolean)

        Try

            If pVal.FormTypeEx <> "VATR" Then Exit Sub

            If pVal.EventType = BoEventTypes.et_ITEM_PRESSED _
                    AndAlso pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.Item(FormUID)

                If pVal.ItemUID = "Item_8" Then

                    Dim FromDate As String = objForm.Items.Item("VRPY").Specific.Value
                    Dim ToDate As String = objForm.Items.Item("VATTO").Specific.Value

                    If FromDate = "" OrElse ToDate = "" Then
                        objMain.objApplication.StatusBar.SetText(
                      "Please enter From Date and To Date",
                     SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Exit Sub
                    End If

                    Try
                        'Open Tax Report
                        objMain.objApplication.ActivateMenuItem("13068")

                        Dim oTaxForm As SAPbouiCOM.Form
                        oTaxForm = objMain.objApplication.Forms.ActiveForm

                        'Set From Date and To Date
                        CType(oTaxForm.Items.Item("5").Specific, SAPbouiCOM.EditText).Value = FromDate
                        CType(oTaxForm.Items.Item("7").Specific, SAPbouiCOM.EditText).Value = ToDate
                        oTaxForm.Items.Item("9").Click(BoCellClickType.ct_Regular)

                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(
            "Tax Report Open Error : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try

                End If

                Select Case pVal.ItemUID


                    Case "Item_7"      'Sales
                        objForm.PaneLevel = 1

                    Case "Item_9"      'Purchase
                        objForm.PaneLevel = 2

                    Case "Item_13"     'Attachment
                        objForm.PaneLevel = 3
                    Case "Item_11"
                        Dim FromDate As String = GetEditValue(objForm, "VRPY")
                        Dim ToDate As String = GetEditValue(objForm, "VATTO")



                        FromDate = ConvertSAPDateToSQL(FromDate)
                        ToDate = ConvertSAPDateToSQL(ToDate)

                        ExportVATToExcel(FromDate, ToDate)


                    Case "Item_10"     'Get Data Button

                        Dim FromDate As String = GetEditValue(objForm, "VRPY")
                        Dim ToDate As String = GetEditValue(objForm, "VATTO")

                        If FromDate = "" OrElse ToDate = "" Then
                            objMain.objApplication.StatusBar.SetText(
                                "Please enter VAT Return Period From and To Date",
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            Exit Sub
                        End If

                        If Not ValidateVATInputDates(objForm) Then
                            Exit Sub
                        End If


                        FromDate = ConvertSAPDateToSQL(FromDate)
                        ToDate = ConvertSAPDateToSQL(ToDate)

                        LoadVATReportLines(FormUID, FromDate, ToDate)

                        '  objForm.Items.Item("Item_10").Enabled = False


                End Select

            End If

            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK AndAlso pVal.BeforeAction = False Then

                If pVal.ItemUID = "btn_Del" Then

                    Try
                        objForm = objMain.objApplication.Forms.Item(FormUID)
                        objForm.Freeze(True)

                        objMatrix2 = CType(objForm.Items.Item("MXT_3").Specific, SAPbouiCOM.Matrix)
                        oDBs_Attach = objForm.DataSources.DBDataSources.Item("@TNX_ATTACH_C3")

                        Dim selectedRow As Integer = 0

                        For i As Integer = 1 To objMatrix2.VisualRowCount
                            If objMatrix2.IsRowSelected(i) = True Then
                                selectedRow = i
                                Exit For
                            End If
                        Next

                        If selectedRow = 0 Then
                            objMain.objApplication.StatusBar.SetText("Please select attachment row.")
                            Exit Try
                        End If

                        objMatrix2.FlushToDataSource()

                        oDBs_Attach.RemoveRecord(selectedRow - 1)

                        If oDBs_Attach.Size = 0 Then
                            oDBs_Attach.InsertRecord(0)
                            oDBs_Attach.SetValue("LineId", 0, "1")
                            oDBs_Attach.SetValue("U_TPA", 0, "")
                            oDBs_Attach.SetValue("U_FN", 0, "")
                            oDBs_Attach.SetValue("U_ATD", 0, "")
                            oDBs_Attach.SetValue("U_FTT", 0, "")
                        Else

                            For i As Integer = 0 To oDBs_Attach.Size - 1
                                oDBs_Attach.SetValue("LineId", i, (i + 1).ToString())
                            Next
                        End If

                        objMatrix2.LoadFromDataSource()
                        objMatrix2.AutoResizeColumns()

                        ' Ensure form goes into update mode after deleting a row when it was in OK mode.
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If

                        objForm.Items.Item("btn_Del").Enabled = False

                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(
                "Delete Attachment Error : " & ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Finally
                        Try
                            objForm.Freeze(False)
                        Catch
                        End Try
                    End Try

                End If
                If pVal.ItemUID = "vaturl" Then

                    Try

                        Dim rs As SAPbobsCOM.Recordset =
                        objMain.objCompany.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset)

                        rs.DoQuery("SELECT TOP 1 ""U_FVRT"" FROM ""@TNX_LKMTR""")

                        If rs.RecordCount > 0 Then

                            Dim VatURL As String =
                            rs.Fields.Item("U_FVRT").Value.ToString().Trim()

                            If VatURL <> "" Then

                                'Open browser
                                Process.Start(VatURL)

                            Else

                                objMain.objApplication.StatusBar.SetText(
                                "VAT URL is empty in Link Master",
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                            End If

                        Else

                            objMain.objApplication.StatusBar.SetText(
                            "Link Master data not found",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        End If

                    Catch ex As Exception

                        objMain.objApplication.StatusBar.SetText(
                        "URL Open Error : " & ex.Message,
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    End Try

                End If
            End If
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK AndAlso pVal.BeforeAction = False Then
                If pVal.ItemUID = "MXT_3" And pVal.ColUID = "TPA" And pVal.BeforeAction = False Then
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("MXT_3").Specific


                    If pVal.Row > 0 AndAlso pVal.Row <= objMatrix.VisualRowCount Then
                        Dim fullPath As String = objMatrix.Columns.Item("TPA").Cells.Item(pVal.Row).Specific.Value
                        If Not String.IsNullOrEmpty(fullPath) AndAlso fullPath.Contains("\") Then
                            Dim indexLoc As Integer = fullPath.LastIndexOf("\")
                            Dim filename As String = fullPath.Substring(indexLoc + 1)
                            objMatrix.Columns.Item("FN").Cells.Item(pVal.Row).Specific.Value = filename
                            objMatrix.Columns.Item("ATD").Cells.Item(pVal.Row).Specific.Value = DateTime.Now.ToString("yyyyMMdd")
                            objForm.Items.Item("btn_Del").Enabled = True
                        End If
                    End If
                End If
            End If
            'sreeja
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "1" AndAlso pVal.BeforeAction = False AndAlso pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Me.SetDefault(objForm.UniqueID)
                        objForm.Items.Item("Item_5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    End If
                    If pVal.ItemUID = "1" AndAlso pVal.BeforeAction = True AndAlso pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Try
                            If Not ValidateVATDates(objForm) Then
                                BubbleEvent = False
                                Exit Sub
                            End If

                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VATRP")

                            Dim dblTVRP As Double = 0

                            Try

                                Dim TVRPValue As String = oDBs_Head.GetValue("U_NVAT", 0).ToString().Trim()

                                If TVRPValue = "" Then

                                    objMain.objApplication.StatusBar.SetText("Net VAT Due value should not be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                                    BubbleEvent = False
                                    Exit Sub

                                End If

                                Double.TryParse(TVRPValue, dblTVRP)

                                If dblTVRP <= 0 Then

                                    objMain.objApplication.StatusBar.SetText("Net VAT Due value should be greater than Zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                                    BubbleEvent = False
                                    Exit Sub

                                End If

                            Catch ex As Exception

                                objMain.objApplication.StatusBar.SetText("NVAT Validation Error : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                                BubbleEvent = False
                                Exit Sub

                            End Try

                            If oDBs_Head Is Nothing Then

                                Exit Try
                            End If

                            Dim currentAppId As String = ""
                            Try
                                currentAppId = oDBs_Head.GetValue("U_APPI", oDBs_Head.Offset).Trim()
                            Catch
                                currentAppId = ""
                            End Try

                            If String.IsNullOrEmpty(currentAppId) Then
                                AutoDocentryNumber(objForm.UniqueID)
                                Try
                                    currentAppId = oDBs_Head.GetValue("U_APPI", oDBs_Head.Offset).Trim()
                                Catch
                                    currentAppId = ""
                                End Try
                            End If

                            Dim DocType As String = "VPA"
                            Dim TableName As String = "@TNX_VATRP"
                            Dim AppIDField As String = "U_APPI"
                            Dim AppStatField As String = "U_FTY"

                            Dim rsapp As SAPbobsCOM.Recordset = Nothing
                            Dim approvalExists As Boolean = False

                            Try
                                rsapp = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                Dim q As String =
                                "SELECT T0.""Code"" As ""TemplateID"",T1.""U_Name"" As ""Originator"", " &
                                "S1.""Code"" As ""Stage"",S2.""U_UKey"" As ""Authorizers"", " &
                                "S1.""U_NAP"" As ""NoOfAuth"",S1.""U_NRJ"" As ""NoOfRej"", " &
                                "T0.""CreateDate"", " &
                                "(Case When T3.""LineId""='1' Then 'S' Else 'O' End) As ""Status"" " &
                                "FROM ""@SBO_APPHDR"" T0 " &
                                "INNER JOIN ""@SBO_APPREQ"" T1 ON T1.""Code""=T0.""Code"" " &
                                "INNER JOIN ""@SBO_APPDOC"" T2 ON T2.""Code"" = T0.""Code"" " &
                                "INNER JOIN ""@SBO_APPAUT"" T3 ON T3.""Code"" = T0.""Code"" " &
                                "INNER JOIN ""@SBO_AST"" S1 ON T3.""U_M3_1""=S1.""Code"" " &
                                "INNER JOIN ""@SBO_AST_C0"" S2 ON S1.""Code""=S2.""Code"" " &
                                "WHERE ""U_Active"" = 'Y' " &
                                "AND T1.""U_Name"" = '" & objMain.objCompany.UserName & "' " &
                                "AND T2.""U_" & DocType & """ = 'Y' " &
                                "AND IFNULL(S2.""U_UKey"",'')<>'' " &
                                "ORDER BY T3.""LineId"""

                                rsapp.DoQuery(q)

                                If rsapp.RecordCount > 0 Then
                                    approvalExists = True
                                End If

                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText("Approval check error: " & ex.Message)
                            Finally
                                If rsapp IsNot Nothing Then
                                    Marshal.ReleaseComObject(rsapp)
                                    rsapp = Nothing
                                End If
                            End Try

                            If approvalExists Then
                                ' Me.ApprovalTrigger(DocType, currentAppId, TableName, AppIDField, AppStatField, "VAT Report")
                            Else
                                Try
                                    oDBs_Head.SetValue("U_FTY", oDBs_Head.Offset, "A")

                                Catch ex As Exception

                                End Try

                                objMain.objApplication.StatusBar.SetText("Document will be auto-approved (no approval template).", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                            End If

                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText("ItemEvent (Before Add) error: " & ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                        End Try


                        Try
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VATRP")
                        Catch ex As Exception
                            oDBs_Head = Nothing
                        End Try

                    End If
                    If pVal.ItemUID = "Item_5" And pVal.BeforeAction = False And SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Try
                            Dim AppID As String = objForm.Items.Item("APPI").Specific.value
                            Me.ApprovalTrigger("VPA", AppID, "@TNX_VATRP", "U_APPI", "U_FTY", "VAT Report")
                        Catch ex As Exception
                        End Try

                    End If
            End Select
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED _
AndAlso pVal.BeforeAction = False Then

                If pVal.ItemUID = "MXT_1" AndAlso pVal.ColUID = "ACN" Then

                    Try

                        objForm = objMain.objApplication.Forms.Item(FormUID)

                        objMatrix = CType(objForm.Items.Item("MXT_1").Specific,
                              SAPbouiCOM.Matrix)

                        Dim AccountCode As String =
            CType(objMatrix.Columns.Item("ACN").Cells.Item(pVal.Row).Specific,
                  SAPbouiCOM.EditText).Value

                        If AccountCode <> "" Then

                            objMain.objApplication.OpenForm(
                    SAPbouiCOM.BoFormObjectEnum.fo_GLAccounts,
                    "",
                    AccountCode)

                        End If

                    Catch ex As Exception

                        objMain.objApplication.StatusBar.SetText(
            "Linked Button Error : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    End Try

                End If

            End If
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED _
AndAlso pVal.BeforeAction = False Then

                If pVal.ItemUID = "MXT_2" AndAlso pVal.ColUID = "ARN" Then

                    Try

                        objForm = objMain.objApplication.Forms.Item(FormUID)

                        objMatrix = CType(objForm.Items.Item("MXT_2").Specific,
                              SAPbouiCOM.Matrix)

                        Dim AccountCode As String =
            CType(objMatrix.Columns.Item("ARN").Cells.Item(pVal.Row).Specific,
                  SAPbouiCOM.EditText).Value

                        If AccountCode <> "" Then

                            objMain.objApplication.OpenForm(
                    SAPbouiCOM.BoFormObjectEnum.fo_GLAccounts,
                    "",
                    AccountCode)

                        End If

                    Catch ex As Exception

                        objMain.objApplication.StatusBar.SetText(
            "Linked Button Error : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    End Try

                End If

            End If


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("ItemEvent Error: " & ex.Message)
        End Try

    End Sub


    Private Function GetEditValue(ByVal oForm As SAPbouiCOM.Form,
                              ByVal ItemUID As String) As String
        Try
            Return CType(oForm.Items.Item(ItemUID).Specific,
                     SAPbouiCOM.EditText).Value.Trim()
        Catch
            Return ""
        End Try
    End Function
    Private Function ConvertSAPDateToSQL(ByVal sapDate As String) As String
        Try
            If sapDate = "" Then Return ""

            'SAP B1 date format: yyyyMMdd
            If sapDate.Length = 8 Then
                Return sapDate.Substring(0, 4) & "-" &
                   sapDate.Substring(4, 2) & "-" &
                   sapDate.Substring(6, 2)
            End If

            Return sapDate

        Catch
            Return sapDate
        End Try
    End Function
    Private Sub LoadVATReportLines(ByVal FormUID As String,
                               ByVal FromDate As String,
                               ByVal ToDate As String)

        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            oForm = objMain.objApplication.Forms.Item(FormUID)
            oForm.Freeze(True)

            Dim oMatrixSales As SAPbouiCOM.Matrix =
            CType(oForm.Items.Item("MXT_1").Specific, SAPbouiCOM.Matrix)

            Dim oMatrixPurchase As SAPbouiCOM.Matrix =
            CType(oForm.Items.Item("MXT_2").Specific, SAPbouiCOM.Matrix)

            Dim oDBH As SAPbouiCOM.DBDataSource =
            oForm.DataSources.DBDataSources.Item("@TNX_VATRP")

            Dim oDBSales As SAPbouiCOM.DBDataSource =
            oForm.DataSources.DBDataSources.Item("@TNX_VATCTM_C1")

            Dim oDBPurchase As SAPbouiCOM.DBDataSource =
            oForm.DataSources.DBDataSources.Item("@TNX_VATRP_C0")

            ClearMatrix(oMatrixSales, oDBSales)
            ClearMatrix(oMatrixPurchase, oDBPurchase)
            'Clear old totals
            oDBH.SetValue("U_TVD", 0, "0")
            oDBH.SetValue("U_TVRP", 0, "0")
            oDBH.SetValue("U_NVAT", 0, "0")

            Dim rs As SAPbobsCOM.Recordset =
            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim sql As String =
            "CALL ""TNX_VAT_RETURN_RPT"" " &
            "(TO_DATE('" & FromDate & "','YYYY-MM-DD'), " &
            "TO_DATE('" & ToDate & "','YYYY-MM-DD'))"

            rs.DoQuery(sql)

            If rs.RecordCount = 0 Then
                objMain.objApplication.StatusBar.SetText("No VAT report data found")
                Exit Sub
            End If

            Dim rowSales As Integer = 0
            Dim rowPurchase As Integer = 0

            While Not rs.EoF

                Dim BoxNo As String = GetRSValue(rs, "BoxNo")
                Dim AccNo As String = GetRSValue(rs, "AccountNo")
                Dim Desc As String = GetRSValue(rs, "Description")
                Dim Amt As String = ToNumber(GetRSValue(rs, "Amount_AED"))
                Dim Vat As String = ToNumber(GetRSValue(rs, "VAT_AED"))
                Dim Adj As String = ToNumber(GetRSValue(rs, "Adjustment_AED"))

                If BoxNo = "12" Then oDBH.SetValue("U_TVD", 0, Vat)
                If BoxNo = "13" Then oDBH.SetValue("U_TVRP", 0, Vat)
                If BoxNo = "14" Then oDBH.SetValue("U_NVAT", 0, Vat)

                If IsOutputVATBox(BoxNo) Then

                    oDBSales.InsertRecord(rowSales)
                    oDBSales.SetValue("LineId", rowSales, (rowSales + 1).ToString())
                    oDBSales.SetValue("U_ACN", rowSales, AccNo)
                    oDBSales.SetValue("U_TRN", rowSales, Desc)
                    oDBSales.SetValue("U_AMT", rowSales, Amt)
                    oDBSales.SetValue("U_VATA", rowSales, Vat)
                    oDBSales.SetValue("U_AST", rowSales, Adj)

                    rowSales += 1

                ElseIf IsInputVATBox(BoxNo) Then

                    oDBPurchase.InsertRecord(rowPurchase)
                    oDBPurchase.SetValue("LineId", rowPurchase, (rowPurchase + 1).ToString())
                    oDBPurchase.SetValue("U_ARN", rowPurchase, AccNo)
                    oDBPurchase.SetValue("U_VATE", rowPurchase, Desc)
                    oDBPurchase.SetValue("U_AUT", rowPurchase, Amt)
                    oDBPurchase.SetValue("U_RVAT", rowPurchase, Vat)
                    oDBPurchase.SetValue("U_AVAT", rowPurchase, Adj)

                    rowPurchase += 1

                End If

                rs.MoveNext()

            End While

            oForm.PaneLevel = 1
            oMatrixSales.LoadFromDataSource()
            oMatrixSales.AutoResizeColumns()

            oForm.PaneLevel = 2
            oMatrixPurchase.LoadFromDataSource()
            oMatrixPurchase.AutoResizeColumns()

            oForm.PaneLevel = 1

            objMain.objApplication.StatusBar.SetText(
            "VAT matrix data loaded successfully",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "LoadVATReportLines Error: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        Finally
            If oForm IsNot Nothing Then oForm.Freeze(False)
        End Try

    End Sub
    Private Sub ClearMatrix(ByVal oMatrix As SAPbouiCOM.Matrix,
                        ByVal oDB As SAPbouiCOM.DBDataSource)

        oMatrix.FlushToDataSource()

        For i As Integer = oDB.Size - 1 To 0 Step -1
            oDB.RemoveRecord(i)
        Next

        oMatrix.Clear()

    End Sub
    Private Sub LoadVATHeaderOnly(ByVal FormUID As String)

        Dim oForm As SAPbouiCOM.Form = Nothing

        Try

            oForm = objMain.objApplication.Forms.Item(FormUID)

            oForm.Freeze(True)

            Dim oDBH As SAPbouiCOM.DBDataSource =
        oForm.DataSources.DBDataSources.Item("@TNX_VATRP")

            If oDBH.Size = 0 Then
                oDBH.InsertRecord(0)
            End If

            oDBH.Offset = 0

            Dim rs As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(
        SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim sql As String =
        "SELECT " &
        "CURRENT_DATE AS ""SubmissionDate"", " &
        "IFNULL(""TaxIdNum"",'') AS ""TRN"", " &
        "IFNULL(""CompnyName"",'') AS ""TaxablePersonNameEnglish"", " &
        "IFNULL(""PrintHeadr"",'') AS ""TaxablePersonNameArabic"", " &
        "IFNULL(""CompnyAddr"",'') AS ""TaxablePersonAddress"" " &
        "FROM OADM"

            rs.DoQuery(sql)

            If rs.RecordCount > 0 Then

                Try
                    oDBH.SetValue(
                "U_SD",
                0,
                ToSAPDate(GetRSValue(rs, "SubmissionDate")))
                Catch
                End Try

                Try
                    oDBH.SetValue(
                "U_TRNM",
                0,
                GetRSValue(rs, "TRN"))
                Catch
                End Try

                Try
                    oDBH.SetValue(
                "U_TPNE",
                0,
                GetRSValue(rs, "TaxablePersonNameEnglish"))
                Catch
                End Try

                Try
                    oDBH.SetValue(
                "U_TPNA",
                0,
                GetRSValue(rs, "TaxablePersonNameArabic"))
                Catch
                End Try

                Try
                    oDBH.SetValue(
                "U_TPA",
                0,
                GetRSValue(rs, "TaxablePersonAddress"))
                Catch
                End Try

            End If

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
        "LoadVATHeaderOnly Error: " & ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        Finally

            Try
                If oForm IsNot Nothing Then
                    oForm.Freeze(False)
                End If
            Catch
            End Try

        End Try

    End Sub

#End Region
    Private Sub SetExportButtonByMode(ByVal oForm As SAPbouiCOM.Form)

        Try

            Dim oDBHead As SAPbouiCOM.DBDataSource =
            oForm.DataSources.DBDataSources.Item("@TNX_VATRP")

            Dim ApprovalStatus As String = ""

            Try
                ApprovalStatus = oDBHead.GetValue("U_FTY", oDBHead.Offset).Trim()
            Catch
                ApprovalStatus = ""
            End Try

            'Enable only when Approval Status = Approved
            If ApprovalStatus = "A" Then

                oForm.Items.Item("Item_11").Enabled = True
                oForm.Items.Item("Item_8").Enabled = True


            Else

                oForm.Items.Item("Item_11").Enabled = False
                oForm.Items.Item("Item_8").Enabled = False

            End If

            'Submit button control
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                oForm.Items.Item("Item_5").Enabled = False

            End If


            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                oForm.Items.Item("Item_10").Enabled = False

            ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                oForm.Items.Item("Item_10").Enabled = True

            End If
        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
            "SetExportButtonByMode Error : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
    Private Function IsOutputVATBox(ByVal BoxNo As String) As Boolean
        Select Case BoxNo
            Case "1a", "1b", "1c", "1d", "1e", "1f", "1g",
                 "2", "3", "4", "5", "6", "7", "8"
                Return True
            Case Else
                Return False
        End Select
    End Function
    Private Function IsInputVATBox(ByVal BoxNo As String) As Boolean
        Select Case BoxNo
            Case "9", "10", "11", "12A"
                Return True
            Case Else
                Return False
        End Select
    End Function
    Private Sub RenumberMatrix(ByVal oMatrix As SAPbouiCOM.Matrix)

        Try
            For i As Integer = 1 To oMatrix.VisualRowCount
                oMatrix.Columns.Item(0).Cells.Item(i).Specific.String = i.ToString()
            Next
        Catch
        End Try

    End Sub
    Private Function GetRSValue(ByVal rs As SAPbobsCOM.Recordset,
                            ByVal FieldName As String) As String
        Try
            If rs.Fields.Item(FieldName).Value Is Nothing Then Return ""
            Return rs.Fields.Item(FieldName).Value.ToString().Trim()
        Catch
            Return ""
        End Try
    End Function
    Private Function ToSAPDate(ByVal value As String) As String
        Try
            If value = "" Then Return ""

            Dim dt As DateTime
            If DateTime.TryParse(value, dt) Then
                Return dt.ToString("yyyyMMdd")
            End If

            Return ""
        Catch
            Return ""
        End Try
    End Function
    Private Function ToNumber(ByVal value As String) As String
        Try
            If value = "" Then Return "0"

            Dim d As Double = 0
            Double.TryParse(value, d)

            Return d.ToString()
        Catch
            Return "0"
        End Try
    End Function
    Private Sub DeleteBlankRows(ByVal oMatrix As SAPbouiCOM.Matrix,
                            ByVal CheckColumnIndex As Integer)

        Try
            For i As Integer = oMatrix.VisualRowCount To 1 Step -1

                Dim value As String = ""

                Try
                    value = CType(oMatrix.Columns.Item(CheckColumnIndex).Cells.Item(i).Specific,
                                  SAPbouiCOM.EditText).Value.Trim()
                Catch
                    value = ""
                End Try

                If value = "" Then
                    oMatrix.DeleteRow(i)
                End If

            Next
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub
    Private Function ValidateVATDates(ByVal oForm As SAPbouiCOM.Form) As Boolean

        Try

            Dim FromDate As String = GetEditValue(oForm, "VRPY")
            Dim ToDate As String = GetEditValue(oForm, "VATTO")

            'From and To Date mandatory
            If FromDate = "" OrElse ToDate = "" Then

                objMain.objApplication.StatusBar.SetText("Please enter VAT Return Period From and To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                Return False

            End If

            Dim dtFrom As DateTime
            Dim dtTo As DateTime

            'Validate From Date
            If Not DateTime.TryParseExact(FromDate,
                                      "yyyyMMdd",
                                      System.Globalization.CultureInfo.InvariantCulture,
                                      Globalization.DateTimeStyles.None,
                                      dtFrom) Then

                objMain.objApplication.StatusBar.SetText(
            "Invalid From Date",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Return False

            End If

            'Validate To Date
            If Not DateTime.TryParseExact(ToDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, dtTo) Then

                objMain.objApplication.StatusBar.SetText("Invalid To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Return False

            End If

            If dtFrom > dtTo Then

                objMain.objApplication.StatusBar.SetText("From Date should not be greater than To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Return False

            End If

            If dtFrom = dtTo Then

                objMain.objApplication.StatusBar.SetText("From Date and To Date should not be same", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Return False

            End If

            'Quarter validation
            'Dim ExpectedToDate As DateTime = New DateTime(dtFrom.Year, dtFrom.Month, 1).AddMonths(3).AddDays(-1)

            'If dtTo <> ExpectedToDate Then

            '    objMain.objApplication.StatusBar.SetText("Please select valid quarter period. Example : 01/01/2026 To 31/03/2026", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            '    Return False

            'End If

            Dim rs As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            'Dim sql As String = "SELECT ""DocNum"" FROM ""@TNX_VATRP"" " & "WHERE " & "((TO_DATE('" & dtFrom.ToString("yyyy-MM-dd") & "') BETWEEN ""U_VRPY"" AND ""U_VATTO"") " & "OR " & "(TO_DATE('" & dtTo.ToString("yyyy-MM-dd") & "') BETWEEN ""U_VRPY"" AND ""U_VATTO""))"
            Dim sql As String =
"SELECT ""DocNum"" " &
"FROM ""@TNX_VATRP"" " &
"WHERE IFNULL(""Canceled"", 'N') <> 'Y' " &
"AND IFNULL(""U_FTY"", '') <> 'N' " &
"AND ( " &
"      (TO_DATE('" & dtFrom.ToString("yyyy-MM-dd") & "') BETWEEN ""U_VRPY"" AND ""U_VATTO"") " &
"   OR (TO_DATE('" & dtTo.ToString("yyyy-MM-dd") & "') BETWEEN ""U_VRPY"" AND ""U_VATTO"") " &
"    )"
            rs.DoQuery(sql)

            If rs.RecordCount > 0 Then

                objMain.objApplication.StatusBar.SetText("Selected VAT period already exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Marshal.ReleaseComObject(rs)

                Return False

            End If

            Marshal.ReleaseComObject(rs)

            Return True

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText("Date Validation Error : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            Return False

        End Try

    End Function
    Private Sub ExportVATToExcel(ByVal FromDate As String, ByVal ToDate As String)

        Dim rs As SAPbobsCOM.Recordset = Nothing

        Try
            rs = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
                   SAPbobsCOM.Recordset)

            Dim sql As String =
            "CALL ""TNX_VAT_RETURN_RPT"" " &
            "(TO_DATE('" & FromDate & "','YYYY-MM-DD'), " &
            "TO_DATE('" & ToDate & "','YYYY-MM-DD'))"

            rs.DoQuery(sql)

            If rs.RecordCount = 0 Then
                objMain.objApplication.StatusBar.SetText(
                "No data found for Excel export",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Exit Sub
            End If

            Dim folderPath As String = "C:\Temp"

            If Not Directory.Exists(folderPath) Then
                Directory.CreateDirectory(folderPath)
            End If

            Dim filePath As String =
            folderPath & "\VAT_Report_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".csv"

            Dim sb As New StringBuilder()

            'Headers
            For i As Integer = 0 To rs.Fields.Count - 1
                sb.Append("""" & rs.Fields.Item(i).Name.Replace("""", """""") & """")
                If i < rs.Fields.Count - 1 Then sb.Append(",")
            Next
            sb.AppendLine()

            'Rows
            While Not rs.EoF

                For i As Integer = 0 To rs.Fields.Count - 1

                    Dim value As String = ""

                    If rs.Fields.Item(i).Value IsNot Nothing Then
                        value = Convert.ToString(rs.Fields.Item(i).Value)
                    End If

                    value = value.Replace("""", """""")

                    sb.Append("""" & value & """")

                    If i < rs.Fields.Count - 1 Then sb.Append(",")

                Next

                sb.AppendLine()
                rs.MoveNext()

            End While

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8)

            If File.Exists(filePath) Then

                Dim excelPath As String = ""

                If File.Exists("C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE") Then
                    excelPath = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"

                ElseIf File.Exists("C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE") Then
                    excelPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"

                ElseIf File.Exists("C:\Program Files\Microsoft Office\Office16\EXCEL.EXE") Then
                    excelPath = "C:\Program Files\Microsoft Office\Office16\EXCEL.EXE"

                ElseIf File.Exists("C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE") Then
                    excelPath = "C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE"
                End If

                If excelPath <> "" Then
                    Process.Start(excelPath, """" & filePath & """")
                Else
                    Dim psi As New ProcessStartInfo()
                    psi.FileName = filePath
                    psi.UseShellExecute = True
                    Process.Start(psi)
                End If

                objMain.objApplication.MessageBox(
                "File created successfully:" & vbCrLf & filePath)

            Else
                objMain.objApplication.StatusBar.SetText(
                "Excel file not created",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
            "Excel Export Error: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        Finally
            If rs IsNot Nothing Then
                Marshal.ReleaseComObject(rs)
                rs = Nothing
            End If
        End Try

    End Sub
End Class

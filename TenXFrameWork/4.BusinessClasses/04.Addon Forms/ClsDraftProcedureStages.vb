Imports System.ServiceModel

Public Class ClsDraftProcedureStages

#Region "       Declaration             "
    Public objform As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details, oDS, oDS1 As SAPbouiCOM.DBDataSource
    Dim oItem As SAPbouiCOM.Item
    Dim objComboBox As SAPbouiCOM.ComboBox
    Public rs, rs1, rs2, oRsDraft As SAPbobsCOM.Recordset
    '' Dim objMatrix, objMatrix1 As SAPbouiCOM.Matrix
    Dim CurDiscountPer As Double
    Dim CurDiscountAmt As Double
    Dim GlobalDocnum As String
    Dim Flag As String = "No"
    Public rsUpdate, rsUpdate1 As SAPbobsCOM.Recordset
    Public Query, Query1, ManagerId, BranchId, PMonth, PYear, DraftMessage As String
    Public TableV, AppIDFieldV, AppStatFieldV, AppIDV, OriginatorV, TemplateIDV, DocTypeV As String
    ' Public PIform As SBO_Form
    Public PIform As SAPbouiCOM.Form
    Public objMatrix, objMatrix2 As SAPbouiCOM.Matrix
    Public i As Integer
    Public oMessageService As SAPbobsCOM.MessagesService
    Public oCmpSrv As SAPbobsCOM.CompanyService

#End Region
    Sub CreateForm(ByVal BaseForm As SAPbouiCOM.Form, ByVal DocType As String, ByVal AppID As String, ByVal Originator As String, ByVal oRsDraft As SAPbobsCOM.Recordset, ByVal Table As String, ByVal AppIDField As String, ByVal AppStatField As String, ByVal Document As String)
        Try
            objMain.objUtilities.LoadForm("DraftStage.xml", "SBO_Draft", ResourceType.Embeded)
            objform = objMain.objApplication.Forms.GetForm("SBO_Draft", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objform.Freeze(True)

            oDS = objform.DataSources.DBDataSources.Add("@SBO_DD")
            oDS1 = objform.DataSources.DBDataSources.Add("@SBO_DD1")

            objform.DataSources.UserDataSources.Add("ds1", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 100)
            objform.DataSources.UserDataSources.Add("ds2", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 100)
            objform.DataSources.UserDataSources.Add("ds3", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 100)
            objMatrix2 = objform.Items.Item("M1").Specific
            Dim oCol As SAPbouiCOM.Column = objMatrix2.Columns.Item("M1_C1")
            oCol.DataBind.SetBound(True, "", "ds1")
            Dim oCol1 As SAPbouiCOM.Column = objMatrix2.Columns.Item("M1_C2")
            oCol1.DataBind.SetBound(True, "", "ds2")
            Dim oCol2 As SAPbouiCOM.Column = objMatrix2.Columns.Item("M1_C3")
            oCol2.DataBind.SetBound(True, "", "ds3")
            '' Me.FormColumn(enControlName.cst_Matrix_M1, enControlName.cst_M1_col_NoofUsers).DataBind.SetBound(True, "", "ds3")
            'Me.FormColumn(enControlName.cst_Matrix_M1, enControlName.cst_M1_Col_Template).DataBind.SetBound(True, "", "ds1")
            'Me.FormColumn(enControlName.cst_Matrix_M1, enControlName.cst_M1_col_Remarks).DataBind.SetBound(True, "", "ds2")
            objMatrix = objform.Items.Item("Mtx").Specific


            TableV = Table
            AppIDFieldV = AppIDField
            AppStatFieldV = AppStatField
            AppIDV = AppID
            OriginatorV = Originator
            DocTypeV = DocType
            PIform = BaseForm
            If oRsDraft.RecordCount > 0 Then
                oRsDraft.MoveFirst()
                Dim T As String = oRsDraft.Fields.Item("TemplateID").Value
                objform.Items.Item("docnum").Specific.Value = AppID
                objform.Items.Item("tempid").Specific.Value = T
                objform.Items.Item("userid").Specific.Value = Originator
                objform.Items.Item("docdate").Specific.Value = DateTime.Now.ToString("yyyyMMdd")

                objform.Items.Item("Doc").Specific.Value = DocType
                objform.Items.Item("Appcnt").Specific.Value = "1"
                Dim objcheck As SAPbouiCOM.CheckBox = objform.Items.Item("Isdraft").Specific
                objcheck.Checked = True
                Dim objcombo As SAPbouiCOM.ComboBox = objform.Items.Item("statusN").Specific
                objcombo.Select("S", SAPbouiCOM.BoSearchKey.psk_ByValue)
                objform.Items.Item("Table").Specific.Value = Table
                objform.Items.Item("AppID").Specific.Value = AppIDField
                objform.Items.Item("AppStat").Specific.Value = AppStatField
                objform.Items.Item("Doc").Specific.Value = Document
                TemplateIDV = oRsDraft.Fields.Item("TemplateID").Value
                'Me.FormCellEdit(enControlName.cst_Matrix_M1, enControlName.cst_M1_Col_Template, Me.FormMatrix(enControlName.cst_Matrix_M1).RowCount).Value = oRsDraft.Fields.Item("TemplateID").Value
                'Me.FormCellEdit(enControlName.cst_Matrix_M1, enControlName.cst_M1_col_NoofUsers, Me.FormMatrix(enControlName.cst_Matrix_M1).RowCount).Value = 1
                'Me.FormCellEdit(enControlName.cst_Matrix_M1, enControlName.cst_M1_col_Remarks, Me.FormMatrix(enControlName.cst_Matrix_M1).RowCount).Value = ""

                'For j = 1 To oRsDraft.RecordCount
                objMatrix2.AddRow()

                Dim rowIndex As Integer = objMatrix2.VisualRowCount

                objMatrix2.Columns.Item("M1_C1").Cells.Item(rowIndex).Specific.Value = TemplateIDV
                objMatrix2.Columns.Item("M1_C2").Cells.Item(rowIndex).Specific.Value = ""
                objMatrix2.Columns.Item("M1_C3").Cells.Item(rowIndex).Specific.Value = 1
                ' Next
                'objMatrix.Clear()
                'oDS1.Clear()
                'objMatrix.FlushToDataSource()

                oDS = objform.DataSources.DBDataSources.Add("@SBO_DD")
                oDS1 = objform.DataSources.DBDataSources.Add("@SBO_DD1")
                objMatrix = objform.Items.Item("Mtx").Specific

                oRsDraft.MoveFirst()
                For i = 1 To oRsDraft.RecordCount

                    objMatrix.AddRow()
                    oDS1.SetValue("LineId", oDS1.Offset, i) ''objMatrix.VisualRowCount)
                    oDS1.SetValue("U_TempId", oDS1.Offset, oRsDraft.Fields.Item("TemplateID").Value)
                    oDS1.SetValue("U_Stage", oDS1.Offset, oRsDraft.Fields.Item("Stage").Value)
                    oDS1.SetValue("U_Userid", oDS1.Offset, oRsDraft.Fields.Item("Authorizers").Value)
                    oDS1.SetValue("U_statusN", oDS1.Offset, oRsDraft.Fields.Item("Status").Value)
                    oDS1.SetValue("U_cdate", oDS1.Offset, DateTime.Now().ToString("yyyyMMdd"))
                    oDS1.SetValue("U_ctime", oDS1.Offset, DateTime.Now().ToString("HHmm"))
                    oDS1.SetValue("U_Udate", oDS1.Offset, "")
                    oDS1.SetValue("U_Utime", oDS1.Offset, "")
                    oDS1.SetValue("U_Remarks", oDS1.Offset, "")
                    objMatrix.SetLineData(objMatrix.VisualRowCount)
                    oRsDraft.MoveNext()
                Next
                objMatrix.AutoResizeColumns()
            End If
            ' Me.SendMessageAlert()

            objform.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            objform.Freeze(False)
        End Try
    End Sub
    Public Sub SendMessageAlert()
        Dim oMessage As SAPbobsCOM.Message
        Dim pMessageDataColumns As SAPbobsCOM.MessageDataColumns '  MessageDataColumns
        Dim pMessageDataColumn As SAPbobsCOM.MessageDataColumn
        Dim oRecipientCollection As SAPbobsCOM.RecipientCollection
        Dim i As Integer
        Dim str, str1 As String
        Dim oRsRecipients As SAPbobsCOM.Recordset
        oRsRecipients = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        str = "SELECT T0.""Code"" As ""TemplateID"",T1.""U_Name"" As ""Originator"",S1.""Code"" As ""Stage"",S2.""U_UKey"" As ""Authorizers"" " &
                               " ,S1.""U_NAP"" As ""NoOfAuth"",S1.""U_NRJ"" As ""NoOfRej"",T0.""CreateDate"" FROM ""@SBO_APPHDR"" T0 " &
                               " INNER JOIN ""@SBO_APPREQ"" T1 ON T1.""Code""=T0.""Code"" " &
                               " INNER JOIN ""@SBO_APPDOC"" T2 ON T2.""Code"" = T0.""Code"" " &
                               " INNER JOIN ""@SBO_APPAUT"" T3 ON T3.""Code"" = T0.""Code"" " &
                               " INNER JOIN ""@SBO_AST"" S1 ON T3.""U_M3_1""=S1.""Code"" " &
                               " INNER JOIN ""@SBO_AST_C0"" S2 ON S1.""Code""=S2.""Code"" " &
                               " WHERE ""U_Active"" = 'Y' AND T1.""U_Name"" = '" & objMain.objCompany.UserName & "'  AND T2.""U_" & DocTypeV & """  = 'Y' " &
                               " AND IFNULL(S2.""U_UKey"",'')<>'' and T3.""LineId""='1' and T0.""Code""='" & TemplateIDV & "';"

        oRsRecipients.DoQuery(str)
        Dim StageV As String = oRsRecipients.Fields.Item("Stage").Value
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
            pMessageDataColumn.MessageDataLines.Item(0).Value = StageV

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

    Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False Then
                        Me.SendMessageAlert()
                        rsUpdate = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        '' If AppStatFieldV = "" Then
                        'Query = "UPDATE """ & TableV & """ SET """ & AppStatFieldV & """ = 'S' where ""U_AppID""='" & AppIDV & "'"
                        ''Else
                        ''    Query = "UPDATE """ & TableV & """ SET """ & AppStatFieldV & """ = 'S' where ""CardCode""='" & AppIDV & "'"
                        ''End If
                        Query = "UPDATE """ & TableV & """ SET """ & AppStatFieldV & """ = 'S' where """ & AppIDFieldV & """='" & AppIDV & "'"
                        rsUpdate.DoQuery(Query)



                        rsUpdate1 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Query1 = "UPDATE ""@SBO_DD"" SET ""U_statusN"" = 'S' where ""U_docnum""='" & AppIDV & "'"
                        rsUpdate1.DoQuery(Query1)

                        objform.Close()
                        PIform.Close() 'base form

                        'For i As Integer = objMain.objApplication.Forms.Count - 1 To 0 Step -1
                        '        Dim frm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(i)

                        '        If frm.TypeEx = "PCRT" Then
                        '            frm.Close()

                        '        ElseIf frm.TypeEx = "S_Rebate" Then
                        '            frm.Close()

                        '        ElseIf frm.TypeEx = "ITEMUP" Then
                        '            frm.Close()

                        '        ElseIf frm.TypeEx = "ME_ITMLIST" Then
                        '            frm.Close()

                        '        ElseIf frm.TypeEx = "134" Then
                        '            frm.Close()

                        '        End If
                        '    Next




                    End If


                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    Try
                        objMatrix = objform.Items.Item("M1").Specific
                        DraftMessage = objMatrix.Columns.Item("M1_C2").Cells.Item(1).Specific.Value
                    Catch ex As Exception
                        Exit Sub
                    End Try
            End Select
        Catch ex As Exception
            objform.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

End Class

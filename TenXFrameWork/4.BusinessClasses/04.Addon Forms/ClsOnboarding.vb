Public Class ClsOnboarding

    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix

    Public rs, RsNum As SAPbobsCOM.Recordset


    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("OnBoarding.xml", "ONBP", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ONBP", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_ONBP")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_ONBP_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific
            Dim CompanyDetailsQry As String = "Select * From ""@TNX_ONBP"" "
            Dim oRsDetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsDetQuery.DoQuery(CompanyDetailsQry)
            If oRsDetQuery.RecordCount > 0 And oRsDetQuery.Fields.Item("Code").Value = "1" Then
                objMain.objApplication.ActivateMenuItem("1290")
            Else
                Me.SetDefault(objForm.UniqueID)
            End If

            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_ONBP")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_ONBP_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific

            oDBs_Head.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"))
            Me.SetNewLine(objForm.UniqueID)

            Dim CompanyDetailsQry As String = "Select T0.""CompnyName"" As ""Organization Name"", T0.""CompnyAddr"" As ""Address"", T0.""Country"" As ""Country"", T0.""E_Mail"" As ""Email"" " &
                " , T0.""FreeZoneNo"" As ""VAT"", T0.""TaxIdNum"" As ""TAXID"", T0.""RevOffice"" As ""Common Name"", T0.""AliasName"" As ""BusinessCategory"" " &
                " , T0.""Phone1"" As ""Mobile"", T1.""ZipCode"",T1.""Street"",T1.""City"" As ""District"",T1.""StreetNo"",T1.""Building"" FROM ""OADM"" T0,ADM1 T1 "
            Dim oRsDetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsDetQuery.DoQuery(CompanyDetailsQry)
            If oRsDetQuery.RecordCount > 0 Then
                oDBs_Head.SetValue("U_CName", 0, oRsDetQuery.Fields.Item("Organization Name").Value)
                oDBs_Head.SetValue("U_Email", 0, oRsDetQuery.Fields.Item("Email").Value)
                oDBs_Head.SetValue("U_City", 0, oRsDetQuery.Fields.Item("Email").Value)
                oDBs_Head.SetValue("U_PCode", 0, oRsDetQuery.Fields.Item("ZipCode").Value)
                oDBs_Head.SetValue("U_Sheet", 0, oRsDetQuery.Fields.Item("Street").Value)
                oDBs_Head.SetValue("U_Address", 0, oRsDetQuery.Fields.Item("Address").Value)
                oDBs_Head.SetValue("U_ANum", 0, oRsDetQuery.Fields.Item("VAT").Value)
                oDBs_Head.SetValue("U_BCat", 0, oRsDetQuery.Fields.Item("BusinessCategory").Value)
                oDBs_Head.SetValue("U_Country", 0, oRsDetQuery.Fields.Item("Country").Value)
                oDBs_Head.SetValue("U_IdNum", 0, oRsDetQuery.Fields.Item("TAXID").Value)
                oDBs_Head.SetValue("U_District", 0, oRsDetQuery.Fields.Item("District").Value)
                oDBs_Head.SetValue("U_SName", 0, oRsDetQuery.Fields.Item("StreetNo").Value)
                oDBs_Head.SetValue("U_BNum", 0, oRsDetQuery.Fields.Item("Building").Value)

            End If
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "2", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objMatrix.AutoResizeColumns()

            objForm.Freeze(False)
            oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMain.objUtilities.getMaxCode("@TNX_ONBP"))
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
            If pVal.MenuUID = "On_Process" And pVal.BeforeAction = False Then
                Me.CreateForm()
            End If
        Catch ex As Exception
                objForm.Freeze(False)
                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub
        Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
            Try
                Select Case pVal.EventType






            End Select
            Catch ex As Exception
                objForm.Freeze(False)
                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Sub SetNewLine(ByVal FormUID As String)
            Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_ONBP")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_ONBP_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific

            objMatrix.AddRow()
            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_CCode", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_PMode", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_IActive", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_ILog", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Options", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Remarks", oDBs_Details.Offset, "")

            objMatrix.SetLineData(objMatrix.VisualRowCount)
                objMatrix.AutoResizeColumns()
            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText(ex.Message)
            End Try
        End Sub






End Class

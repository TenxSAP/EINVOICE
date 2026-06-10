Public Class LinkMaster

#Region "Decliration"

    Public objForm As SAPbouiCOM.Form
        Public oDBs_Content As SAPbouiCOM.DBDataSource
        Public Prgbar As SAPbouiCOM.ProgressBar
        Dim objutilities As Utilities

#End Region

        Sub CreateForm()

            Try

            objMain.objUtilities.LoadForm(
            "Linkmaster.xml",
            "frm_LKMTR",
            ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm(
            "frm_LKMTR",
            objMain.objApplication.Forms.ActiveForm.TypeCount)

                objForm.Freeze(True)

                objutilities = New Utilities

                oDBs_Content = objForm.DataSources.DBDataSources.Item("@TNX_LKMTR")

                Dim rs As SAPbobsCOM.Recordset

                rs = objMain.objCompany.GetBusinessObject(
             SAPbobsCOM.BoObjectTypes.BoRecordset)

                rs.DoQuery(
        "SELECT TOP 1 ""Code"" FROM ""@TNX_LKMTR""")

                ' =====================================
                ' FIRST TIME
                ' =====================================
                If rs.RecordCount = 0 Then

                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE

                    CType(objForm.Items.Item("Code").Specific,
                   SAPbouiCOM.EditText).Value = "1"

                    'CType(objForm.Items.Item("txtName").Specific,
                    '   SAPbouiCOM.EditText).Value = "1"

                Else

                    Me.LoadExistingRecord()

                End If

                objForm.Freeze(False)

                objMain.objApplication.StatusBar.SetText(
            "Successfully initialized, please proceed..",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Catch ex As Exception

                objForm.Freeze(False)

                objMain.objApplication.StatusBar.SetText(
            ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            End Try

        End Sub
        Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)

            Try
                If pVal.MenuUID = "LKMT" And pVal.BeforeAction = False Then
                    Me.CreateForm()
                End If
                If pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                    'Me.SetDefault(objForm.UniqueID)
                    Me.LoadExistingRecord()

                End If
                'If pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                '    Me.SetNewLineContent(objForm.UniqueID)
                'End If

            Catch ex As Exception

                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            End Try

        End Sub

        Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)

            Try

                Select Case pVal.EventType


                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                        'If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True Then
                        '    ' Me.SetDefault(objForm.UniqueID)
                        '    Me.LoadExistingRecord()
                        'End If

                End Select

            Catch ex As Exception

                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            End Try

        End Sub

        Sub LoadExistingRecord()

            Try

                Dim rs As SAPbobsCOM.Recordset

                rs = objMain.objCompany.GetBusinessObject(
                 SAPbobsCOM.BoObjectTypes.BoRecordset)

                rs.DoQuery(
            "SELECT TOP 1 * FROM ""@TNX_LKMTR""")

                If rs.RecordCount > 0 Then

                    Dim strCode As String

                    strCode = rs.Fields.Item("Code").Value.ToString()

                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE

                    CType(objForm.Items.Item("Code").Specific,
       SAPbouiCOM.EditText).Value = strCode

                    objForm.Items.Item("1").Click()

                    'objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

                End If

            Catch ex As Exception

                objMain.objApplication.StatusBar.SetText(ex.Message)

            End Try

        End Sub
    End Class


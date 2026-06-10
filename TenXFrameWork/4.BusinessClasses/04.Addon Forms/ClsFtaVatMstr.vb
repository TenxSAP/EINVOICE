Imports System
Imports System.Security.AccessControl

Public Class ClsFtaVatMstr

#Region "Decliration"

    Public objForm As SAPbouiCOM.Form
        Dim oDBs_Head As SAPbouiCOM.DBDataSource
        Dim oDBs_Details As SAPbouiCOM.DBDataSource
        Dim objMatrix, objMatrix1 As SAPbouiCOM.Matrix
        Dim objComboBox As SAPbouiCOM.ComboBox
        Dim str, str1 As String
        Public rs, RsNum As SAPbobsCOM.Recordset
        Dim LostFocusFlag As Boolean = False
        Dim oGrid As SAPbouiCOM.Grid
        Dim oDt As SAPbouiCOM.DataTable
        Dim objutilities As Utilities

#End Region

        Sub CreateForm()
            Try
            objMain.objUtilities.LoadForm("FtavatMstr.xml", "frm_FTAVM", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("frm_FTAVM", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objForm.Freeze(True)
                objutilities = New Utilities
                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_FTAVAT")
                objMatrix = objForm.Items.Item("CMtx").Specific
                Dim rs As SAPbobsCOM.Recordset

                rs = objMain.objCompany.GetBusinessObject(
     SAPbobsCOM.BoObjectTypes.BoRecordset)

                rs.DoQuery(
"SELECT TOP 1 ""Code"" FROM ""@TNX_FTAVAT""")

                If rs.RecordCount = 0 Then

                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE

                    Me.SetNewLine(objForm.UniqueID)

                Else

                    Me.MatrixLoad()

                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

                End If

                objForm.Freeze(False)
                objForm.EnableMenu("1292", True)

                objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Catch ex As Exception
                objForm.Freeze(False)
                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub
        Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                If pVal.MenuUID = "FTAV" And pVal.BeforeAction = False Then
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


                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "1" _
    And pVal.BeforeAction = False _
    And pVal.ActionSuccess = True Then

                            ' Reload saved data
                            Me.MatrixLoad()

                            ' Switch to OK Mode
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

                        End If



                End Select
            Catch ex As Exception
                objForm.Freeze(False)
                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub


        Sub SetNewLine(ByVal FormUID As String)

            Try

                objForm = objMain.objApplication.Forms.Item(FormUID)

                objMatrix = objForm.Items.Item("CMtx").Specific

                If objMatrix.VisualRowCount = 0 Then
                    objMatrix.AddRow()
                End If

                objMatrix.Columns.Item("DocEntry") _
                 .Cells.Item(1).Specific.Value = 1

                objMatrix.AutoResizeColumns()

            Catch ex As Exception

                objMain.objApplication.StatusBar.SetText(ex.Message)

            End Try

        End Sub
        Public Sub MatrixLoad()

            Try

                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_FTAVAT")

                objMatrix = objForm.Items.Item("CMtx").Specific

                Dim rs1 As String

                rs1 =
            "SELECT ""Code"", ""U_CMR"", ""U_SO"", " &
            """U_GRPO"", ""U_DCNF"", ""U_CMN"" " &
            "FROM ""@TNX_FTAVAT"""

                Dim ors1 As SAPbobsCOM.Recordset

                ors1 = objMain.objCompany.GetBusinessObject(
            SAPbobsCOM.BoObjectTypes.BoRecordset)

                ors1.DoQuery(rs1)

                If ors1.RecordCount > 0 Then

                    objMatrix.Clear()

                    For j As Integer = 1 To ors1.RecordCount

                        objMatrix.AddRow()

                        oDBs_Head.Offset = j - 1

                        oDBs_Head.SetValue("DocEntry",
                                       oDBs_Head.Offset,
                                       j)

                        oDBs_Head.SetValue("Code",
                                       oDBs_Head.Offset,
                                       ors1.Fields.Item("Code").Value)

                        oDBs_Head.SetValue("U_CMR",
                                       oDBs_Head.Offset,
                                       ors1.Fields.Item("U_CMR").Value)

                        oDBs_Head.SetValue("U_SO",
                                       oDBs_Head.Offset,
                                       ors1.Fields.Item("U_SO").Value)

                        oDBs_Head.SetValue("U_GRPO",
                                       oDBs_Head.Offset,
                                       ors1.Fields.Item("U_GRPO").Value)

                        oDBs_Head.SetValue("U_DCNF",
                                       oDBs_Head.Offset,
                                       CType(
                                       ors1.Fields.Item("U_DCNF").Value,
                                       DateTime).ToString("yyyyMMdd"))

                        oDBs_Head.SetValue("U_CMN",
                                       oDBs_Head.Offset,
                                       ors1.Fields.Item("U_CMN").Value)

                        objMatrix.SetLineData(j)

                        ors1.MoveNext()

                    Next

                    objForm.Mode =
                SAPbouiCOM.BoFormMode.fm_OK_MODE

                Else

                    Me.SetNewLine(objForm.UniqueID)

                End If

                objMatrix.FlushToDataSource()

                objMatrix.AutoResizeColumns()

            Catch ex As Exception

                objMain.objApplication.StatusBar.SetText(ex.Message)

            End Try

        End Sub



    End Class

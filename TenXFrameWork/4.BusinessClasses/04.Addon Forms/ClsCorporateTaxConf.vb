
Imports SAPbobsCOM
Imports System
Imports System.IO
Imports System.Text
Imports System.Net.Http
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates

Imports ZXing
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Xml.XmlReader
Imports System.Xml.XmlWriter

Imports System.Xml.Serialization
Imports Zatca.EInvoice.SDK
Imports Zatca.EInvoice.SDK.Contracts.Models
Imports Newtonsoft.Json
Imports TenXFrameWork.Zatca_CSID.Models
Imports TenXFrameWork.Zatca.Models
Imports uuids
Imports Org.BouncyCastle.Math.EC
Imports ECCurve = System.Security.Cryptography.ECCurve
Imports FormatException = System.FormatException
Imports System.Security.Cryptography.Xml
Imports Org.BouncyCastle.Crypto
Imports Org.BouncyCastle.Security
Imports Org.BouncyCastle.Crypto.Signers
Imports Org.BouncyCastle.Crypto.Parameters
Imports Org.BouncyCastle.OpenSsl
Imports TenXFrameWork.FlickINVModel
Imports System.Net.Http.Headers

Public Class ClsCorporateTaxConf
#Region "Declaration"
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
            objMain.objUtilities.LoadForm("CorporateTaxConfiguration.xml", "CTAXC", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("CTAXC", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            objutilities = New Utilities
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCNF")
            objMatrix = objForm.Items.Item("MXT_1").Specific

            objForm.Freeze(False)
            Me.MatrixLoad()
            objForm.EnableMenu("1292", True)
            ' Me.GlAccountCode()
            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "CTAXC" And pVal.BeforeAction = False Then
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                objMatrix = objForm.Items.Item("MXT_1").Specific
                Me.SetNewLine(objForm.UniqueID)
            End If
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCNF")
                    ' oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PALLET_C0")
                    objMatrix = objForm.Items.Item("MXT_1").Specific
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects

                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If oCFL.UniqueID = "CFL_1" Then
                            oDBs_Head.SetValue("U_EAccount", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))
                            objMatrix.SetLineData(pVal.Row)
                        End If
                        If oCFL.UniqueID = "CFL_0" Then
                            oDBs_Head.SetValue("U_LAccount", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))
                            objMatrix.SetLineData(pVal.Row)
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                    If pVal.BeforeAction = False Then

                        objForm = objMain.objApplication.Forms.Item(FormUID)
                        objMatrix = CType(objForm.Items.Item("MXT_1").Specific, SAPbouiCOM.Matrix)

                        If pVal.ItemUID = "MXT_1" AndAlso
                           (pVal.ColUID = "LAccount" OrElse pVal.ColUID = "EAccount") Then

                            Dim AcctCode As String =
                                CType(objMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific,
                                SAPbouiCOM.EditText).Value.Trim()

                            If AcctCode <> "" Then
                                objMain.objApplication.OpenForm(
                                    SAPbouiCOM.BoFormObjectEnum.fo_GLAccounts,
                                    "",
                                    AcctCode)
                            End If

                        End If

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
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCNF")
            objMatrix = objForm.Items.Item("MXT_1").Specific
            objMatrix.AddRow()
            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.VisualRowCount)
            oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMatrix.VisualRowCount)
            oDBs_Head.SetValue("U_MnProfit", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_MxProfit", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_TaxPrc", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_LAccount", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_EAccount", oDBs_Head.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
            ' objMatrix.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub MatrixLoad()
        Try
            oDBs_Head = objForm.DataSources.DBDataSources.Add("@TNX_CTAXCNF")
            Dim objsectMat As SAPbouiCOM.Matrix = objForm.Items.Item("MXT_1").Specific

            Dim rs1 As String = "SELECT ""DocEntry"",""Code"", ""U_MnProfit"", ""U_MxProfit"", ""U_TaxPrc"", ""U_LAccount"", ""U_EAccount"" " &
"FROM ""@TNX_CTAXCNF"" ORDER BY ""Code"" "
            Dim ors1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ors1.DoQuery(rs1)
            If ors1.RecordCount > 0 Then
                objsectMat.Clear()
                For j As Integer = 1 To ors1.RecordCount
                    objsectMat.AddRow()
                    oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, ors1.Fields.Item("DocEntry").Value)
                    oDBs_Head.SetValue("Code", oDBs_Head.Offset, ors1.Fields.Item("Code").Value)
                    oDBs_Head.SetValue("U_MnProfit", oDBs_Head.Offset, ors1.Fields.Item("U_MnProfit").Value)
                    oDBs_Head.SetValue("U_MxProfit", oDBs_Head.Offset, ors1.Fields.Item("U_MxProfit").Value)
                    oDBs_Head.SetValue("U_TaxPrc", oDBs_Head.Offset, ors1.Fields.Item("U_TaxPrc").Value)
                    oDBs_Head.SetValue("U_LAccount", oDBs_Head.Offset, ors1.Fields.Item("U_LAccount").Value)
                    oDBs_Head.SetValue("U_EAccount", oDBs_Head.Offset, ors1.Fields.Item("U_EAccount").Value)
                    objsectMat.SetLineData(objsectMat.VisualRowCount)
                    ors1.MoveNext()
                Next
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            Else
                Me.SetNewLine(objForm.UniqueID)
            End If
            objMatrix.FlushToDataSource()

        Catch ex As Exception
            oDBs_Head.Freeze(False)
        End Try
    End Sub


End Class

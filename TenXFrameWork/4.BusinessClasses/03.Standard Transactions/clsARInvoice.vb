Imports System
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
Imports System.Security.Cryptography.Xml
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Xml.XmlReader
Imports System.Xml.XmlWriter
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Org.BouncyCastle.Crypto
Imports Org.BouncyCastle.Crypto.Parameters
Imports Org.BouncyCastle.Crypto.Signers
Imports Org.BouncyCastle.Math.EC
Imports Org.BouncyCastle.OpenSsl
Imports Org.BouncyCastle.Security
Imports SAPbobsCOM
Imports TenXFrameWork.FlickINVModel
Imports TenXFrameWork.Zatca.Models
Imports TenXFrameWork.Zatca_CSID.Models
Imports uuids
Imports Zatca.EInvoice.SDK
Imports Zatca.EInvoice.SDK.Contracts.Models
Imports ZXing
Imports ECCurve = System.Security.Cryptography.ECCurve
Imports FormatException = System.FormatException

Public Class clsARInvoice

#Region "        Declaration        "
    Dim objForm, AltForm, objMIForm As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim SplitCombo As SAPbouiCOM.ComboBox
    Dim SplitFlgCombo As SAPbouiCOM.ComboBox
    '  Public TLV As TLV

    Private EditText0 As SAPbouiCOM.EditText

    Private EditText1 As SAPbouiCOM.EditText

#End Region

#Region "Declaration"
    Dim oitem As SAPbouiCOM.Item
#End Region

    Public Function GetFormattedTime(t As String) As List(Of String)
        Dim lstTime As New List(Of String)()

        If t.Length < 6 Then
            Dim t2 = "0" & t
            lstTime.Add(Convert.ToString(t2(0)) & Convert.ToString(t2(1)))
            lstTime.Add(Convert.ToString(t2(2)) & Convert.ToString(t2(3)))
            lstTime.Add(Convert.ToString(t2(4)) & Convert.ToString(t2(5)))
        Else
            Dim t1 = t
            lstTime.Add(Convert.ToString(t1(0)) & Convert.ToString(t1(1)))
            lstTime.Add(Convert.ToString(t1(2)) & Convert.ToString(t1(3)))
            lstTime.Add(Convert.ToString(t1(4)) & Convert.ToString(t1(5)))
        End If

        Return lstTime
    End Function
    Private Sub ClearEditText(ByVal oForm As SAPbouiCOM.Form, ByVal itemId As String)

        Try
            Dim oItem As SAPbouiCOM.Item = oForm.Items.Item(itemId)

            If TypeOf oItem.Specific Is SAPbouiCOM.EditText Then
                Dim oEdit As SAPbouiCOM.EditText = oItem.Specific

                ' Clear UI
                oEdit.Value = ""

                ' Clear binding also (IMPORTANT)
                If oEdit.DataBind IsNot Nothing Then
                    If oEdit.DataBind.TableName <> "" Then
                        oForm.DataSources.DBDataSources.Item(oEdit.DataBind.TableName).SetValue(oEdit.DataBind.Alias, 0, "")
                    ElseIf oEdit.DataBind.UserDataSource <> "" Then
                        oForm.DataSources.UserDataSources.Item(oEdit.DataBind.UserDataSource).ValueEx = ""
                    End If
                End If
            End If

        Catch
        End Try

    End Sub
    Private Sub ClearFields(ByVal oForm As SAPbouiCOM.Form)

        Try
            CType(oForm.Items.Item("TXTPSA").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("TXTAM").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("TXTTAX").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("TXTUUID").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("TXTDATE").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("TXTPAY").Specific, SAPbouiCOM.EditText).Value = ""
        Catch
        End Try

    End Sub
    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent,
                     ByRef BubbleEvent As Boolean)

        Try

            If pVal.BeforeAction = False Then


                If pVal.MenuUID = "1282" Then

                    Dim objForm As SAPbouiCOM.Form
                    objForm = objMain.objApplication.Forms.ActiveForm

                    If objForm.TypeEx = "133" Then

                        objForm.Items.Item("Split").Enabled = False

                        objForm.Items.Item("Ainvoice").Enabled = False


                    End If

                End If

            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix = objForm.Items.Item("38").Specific


                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "ARN" And pVal.BeforeAction = False Then

                        objForm = objMain.objApplication.Forms.Item(FormUID)

                        objForm.PaneLevel = 44

                        'Do not load data in Add Mode
                        If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            LoadEInvoiceData(objForm)
                        Else
                            ClearFields(objForm)
                        End If

                        'Hide Response Label
                        Try
                            objForm.Items.Item("LPAY").Visible = False
                        Catch ex As Exception
                        End Try

                        'Hide Response Textbox
                        Try
                            objForm.Items.Item("TXTPAY").Visible = False
                        Catch ex As Exception
                        End Try

                    End If
                    'If pVal.ItemUID = "ARN" And pVal.BeforeAction = False Then

                    '    objForm.PaneLevel = "44"

                    '    ' ❌ DO NOT load data in Add Mode
                    '    If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    '        LoadEInvoiceData(objForm)
                    '    Else
                    '        ' ✅ Clear fields in Add Mode
                    '        ClearFields(objForm)
                    '    End If
                    '    If objForm.Items.Exists("LPAY") Then
                    '        objForm.Items.Item("LPAY").Visible = False
                    '    End If

                    '    If objForm.Items.Exists("TXTPAY") Then
                    '        objForm.Items.Item("TXTPAY").Visible = False
                    '    End If

                    'End If


                    If pVal.ItemUID = "Ainvoice" And pVal.BeforeAction = False Then

                        Dim oPay As New ClsPayLoad


                        oPay.CreateForm()

                        oPay.LoadFieldValues(FormUID)

                        Exit Sub

                    End If


                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix = objForm.Items.Item("38").Specific
                    If pVal.ItemUID = "Split" And pVal.BeforeAction = True And objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Dim sErrMsg As String = Nothing
                        Dim lErrCode As Integer = 0

                        Try

                            Dim recinvoice As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(300), SAPbobsCOM.Recordset)
                            Dim DocNum As String = objForm.Items.Item("8").Specific.Value.ToString()
                            Dim Series As String = objForm.Items.Item("88").Specific.Value.ToString()
                            Dim Qry As String = "SELECT ""DocEntry"" FROM OINV T0 WHERE T0.""DocNum"" = '" & DocNum & "' and T0.""Series""='" & Series & "'"
                            recinvoice.DoQuery(Qry)

                            Dim DocEntry As String = recinvoice.Fields.Item(0).Value().ToString()
                            Me.E_INVPosting(DocEntry)


                            Exit Sub
                        Catch exception As System.Exception
                            Dim ex As System.Exception = exception
                            ErrorLog(ex.Message)
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        End Try

                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                    If pVal.BeforeAction = False Then

                        objForm = objMain.objApplication.Forms.Item(FormUID)
                        objMatrix = objForm.Items.Item("38").Specific

                        'Add custom controls
                        Me.AddItems(objForm.UniqueID)

                        'Hide LPAY
                        Try
                            objForm.Items.Item("LPAY").Visible = False
                        Catch ex As Exception
                        End Try

                        'Hide TXTPAY
                        Try
                            objForm.Items.Item("TXTPAY").Visible = False
                        Catch ex As Exception
                        End Try

                        'Disable buttons
                        Try
                            objForm.Items.Item("Split").Enabled = False
                        Catch ex As Exception
                        End Try

                        Try
                            objForm.Items.Item("Ainvoice").Enabled = False
                        Catch ex As Exception
                        End Try

                    End If
                    'Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    '    objForm = objMain.objApplication.Forms.Item(FormUID)
                    '    objMatrix = objForm.Items.Item("38").Specific
                    '    If pVal.BeforeAction = False Then

                    '        Me.AddItems(objForm.UniqueID)
                    '        If objForm.Items.Exists("LPAY") Then
                    '            objForm.Items.Item("LPAY").Visible = False
                    '        End If

                    '        If objForm.Items.Exists("TXTPAY") Then
                    '            objForm.Items.Item("TXTPAY").Visible = False
                    '        End If
                    '        objForm.Items.Item("Split").Enabled = False
                    '        objForm.Items.Item("Ainvoice").Enabled = False
                    '    End If
            End Select
        Catch ex As Exception
            ErrorLog(ex.Message)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Sub FormDataEvent1(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo,
                         ByRef BubbleEvent As Boolean)

        Try

            Select Case BusinessObjectInfo.EventType

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD


                    If BusinessObjectInfo.BeforeAction = False And
                   BusinessObjectInfo.ActionSuccess = True Then

                        objForm = objMain.objApplication.Forms.GetForm("133",
                    objMain.objApplication.Forms.ActiveForm.TypeCount)


                        Dim query As String =
                "SELECT ""U_STATUS1"" FROM OINV " &
                "WHERE ""DocNum"" = '" & objForm.Items.Item("8").Specific.Value & "' " &
                "AND ""Series"" = '" & objForm.Items.Item("88").Specific.Value & "'"

                        Dim oRs As SAPbobsCOM.Recordset =
                objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        oRs.DoQuery(query)

                        If oRs.RecordCount > 0 Then

                            Dim status As String =
                    oRs.Fields.Item("U_STATUS1").Value.ToString().Trim().ToUpper()

                            If status = "SUCCESS" Then

                                'Posting button disable
                                objForm.Items.Item("Split").Enabled = False

                                'Reconciliation button enable
                                objForm.Items.Item("Ainvoice").Enabled = True
                                objForm.Items.Item("LPAY").Visible = False
                                objForm.Items.Item("TXTPAY").Visible = False
                            End If
                            If status = "" Then

                                'Posting button disable
                                objForm.Items.Item("Split").Enabled = True

                                'Reconciliation button enable
                                objForm.Items.Item("Ainvoice").Enabled = False
                                objForm.Items.Item("LPAY").Visible = False
                                objForm.Items.Item("TXTPAY").Visible = False
                            End If

                        End If

                    End If

            End Select

        Catch ex As Exception

            ErrorLog(ex.Message)

            objMain.objApplication.StatusBar.SetText(ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try

            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then

                        Try
                            objForm = objMain.objApplication.Forms.GetForm("133", objMain.objApplication.Forms.ActiveForm.TypeCount)

                            Dim GetInv As String = " Select Max(""DocEntry"")  From OINV"
                            Dim oRsGetInv As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetInv.DoQuery(GetInv)

                            If oRsGetInv.RecordCount > 0 Then
                                Dim DocEntry As String = oRsGetInv.Fields.Item(0).Value
                                Me.E_INVPosting(DocEntry)
                            End If
                        Catch ex As Exception
                            ErrorLog(ex.Message)
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then

                        objForm = objMain.objApplication.Forms.GetForm("133", objMain.objApplication.Forms.ActiveForm.TypeCount)


                        LoadEInvoiceData(objForm)

                    End If

                    'Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    '    If BusinessObjectInfo.BeforeAction = False And
                    '   BusinessObjectInfo.ActionSuccess = True Then

                    '        objForm = objMain.objApplication.Forms.GetForm("133",
                    '    objMain.objApplication.Forms.ActiveForm.TypeCount)


                    '        Dim query As String =
                    '"SELECT ""U_STATUS1"" FROM OINV " &
                    '"WHERE ""DocNum"" = '" & objForm.Items.Item("8").Specific.Value & "' " &
                    '"AND ""Series"" = '" & objForm.Items.Item("88").Specific.Value & "'"

                    '        Dim oRs As SAPbobsCOM.Recordset =
                    'objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    '        oRs.DoQuery(query)

                    '        If oRs.RecordCount > 0 Then

                    '            Dim status As String =
                    '    oRs.Fields.Item("U_STATUS1").Value.ToString().Trim().ToUpper()

                    '            If status = "SUCCESS" Then

                    '                'Posting button disable
                    '                objForm.Items.Item("Split").Enabled = False

                    '                'Reconciliation button enable
                    '                objForm.Items.Item("Ainvoice").Enabled = True

                    '            End If
                    '            If status = "" Then

                    '                'Posting button disable
                    '                objForm.Items.Item("Split").Enabled = True

                    '                'Reconciliation button enable
                    '                objForm.Items.Item("Ainvoice").Enabled = False

                    '            End If

                    '        End If

                    '    End If

                    'If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                    '    Try
                    '        objForm = objMain.objApplication.Forms.GetForm("133", objMain.objApplication.Forms.ActiveForm.TypeCount)

                    '        Dim GetInv As String = " Select ""U_STATUS1""  From OINV Where ""DocNum""='" & objForm.Items.Item("8").Specific.Value & "' AND ""Series""='" & objForm.Items.Item("88").Specific.Value & "'"
                    '        Dim oRsGetInv As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '        oRsGetInv.DoQuery(GetInv)
                    '        'If oRsGetInv.RecordCount > 0 Then
                    '        '    Dim status As String = oRsGetInv.Fields.Item("U_STATUS1").Value.ToString().Trim()

                    '        '    ' If field has value we are hiding button
                    '        '    If status <> "" Then
                    '        '        objForm.Items.Item("Split").Enabled = False
                    '        '    End If
                    '        '    'If oRsGetInv.Fields.Item(0).Value = "success" Then
                    '        '    '    objForm.Items.Item("Split").Enabled = False
                    '        '    'End If
                    '        'End If
                    '    Catch ex As Exception
                    '        ErrorLog(ex.Message)
                    '        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    End Try
                    'End If

            End Select
        Catch ex As Exception
            ErrorLog(ex.Message)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub




    Private Sub AddItems(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim objMatrix As SAPbouiCOM.Matrix
            objMatrix = objForm.Items.Item("38").Specific

            oitem = objForm.Items.Add("Split", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oitem.Top = objForm.Items.Item("2").Top
            oitem.Left = objForm.Items.Item("2").Left + objForm.Items.Item("2").Width + 5
            oitem.Width = objForm.Items.Item("2").Width + 40
            oitem.Height = objForm.Items.Item("2").Height
            oitem.FromPane = objForm.Items.Item("2").FromPane
            oitem.ToPane = objForm.Items.Item("2").ToPane
            Dim btn As SAPbouiCOM.Button = objForm.Items.Item("Split").Specific
            btn.Caption = "E Invoice Posting"

            oitem = objForm.Items.Add("Ainvoice", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oitem.Top = objForm.Items.Item("Split").Top - 25
            oitem.Left = objForm.Items.Item("Split").Left + 110
            oitem.Width = objForm.Items.Item("Split").Width
            oitem.Height = objForm.Items.Item("Split").Height
            oitem.FromPane = objForm.Items.Item("Split").FromPane
            oitem.ToPane = objForm.Items.Item("Split").ToPane
            Dim btn1 As SAPbouiCOM.Button = objForm.Items.Item("Ainvoice").Specific
            btn1.Caption = "E Invoice Preview"

            oitem = objForm.Items.Add("ARN", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oitem.Top = objForm.Items.Item("138").Top
            oitem.Left = objForm.Items.Item("138").Left
            oitem.Width = objForm.Items.Item("138").Width + 100
            oitem.Height = objForm.Items.Item("138").Height
            oitem.AffectsFormMode = False
            Dim fld3 As SAPbouiCOM.Folder = objForm.Items.Item("ARN").Specific
            fld3.Caption = "E Invoice"
            fld3.GroupWith("138")
            fld3.Pane = "44"
            oitem = objForm.Items.Add("PSA", SAPbouiCOM.BoFormItemTypes.it_STATIC)

            oitem.Top = objForm.Items.Item("38").Top + 1
            oitem.Left = objForm.Items.Item("38").Left + 10
            oitem.Width = 120
            oitem.Height = 15

            oitem.FromPane = 44
            oitem.ToPane = 44

            Dim lbl As SAPbouiCOM.StaticText
            lbl = objForm.Items.Item("PSA").Specific
            lbl.Caption = "Posting Status"

            oitem = objForm.Items.Add("TXTPSA", SAPbouiCOM.BoFormItemTypes.it_EDIT)

            oitem.Top = objForm.Items.Item("PSA").Top + 1
            oitem.Left = objForm.Items.Item("PSA").Left + 90
            oitem.Width = 150
            oitem.Height = 15

            oitem.FromPane = 44
            oitem.ToPane = 44

            Dim txt As SAPbouiCOM.EditText
            txt = objForm.Items.Item("TXTPSA").Specific
            '---------------------------
            ' Total Amount Label
            '---------------------------
            oitem = objForm.Items.Add("TAM", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oitem.Top = objForm.Items.Item("PSA").Top + 22
            oitem.Left = objForm.Items.Item("PSA").Left
            oitem.Width = objForm.Items.Item("PSA").Width
            oitem.Height = objForm.Items.Item("PSA").Height
            oitem.FromPane = 44
            oitem.ToPane = 44

            Dim lbl1 As SAPbouiCOM.StaticText
            lbl1 = objForm.Items.Item("TAM").Specific
            lbl1.Caption = "Total Amount"

            '---------------------------
            ' Total Amount Textbox
            '---------------------------
            oitem = objForm.Items.Add("TXTAM", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oitem.Top = objForm.Items.Item("TAM").Top
            oitem.Left = objForm.Items.Item("TXTPSA").Left
            oitem.Width = objForm.Items.Item("TXTPSA").Width
            oitem.Height = objForm.Items.Item("TXTPSA").Height
            oitem.FromPane = 44
            oitem.ToPane = 44
            Dim txt1 As SAPbouiCOM.EditText
            txt1 = objForm.Items.Item("TXTAM").Specific
            '=========================================
            ' UUID Label
            '=========================================
            oitem = objForm.Items.Add("LUUID", SAPbouiCOM.BoFormItemTypes.it_STATIC)

            oitem.Top = objForm.Items.Item("TAM").Top + 22
            oitem.Left = objForm.Items.Item("PSA").Left
            oitem.Width = objForm.Items.Item("PSA").Width
            oitem.Height = objForm.Items.Item("PSA").Height
            oitem.FromPane = 44
            oitem.ToPane = 44

            Dim lbl2 As SAPbouiCOM.StaticText
            lbl2 = objForm.Items.Item("LUUID").Specific
            lbl2.Caption = "UUID"

            '=========================================
            ' UUID Edit Box
            '=========================================
            oitem = objForm.Items.Add("TXTUUID", SAPbouiCOM.BoFormItemTypes.it_EDIT)

            oitem.Top = objForm.Items.Item("LUUID").Top
            oitem.Left = objForm.Items.Item("TXTPSA").Left
            oitem.Width = objForm.Items.Item("TXTPSA").Width
            oitem.Height = objForm.Items.Item("TXTPSA").Height
            oitem.FromPane = 44
            oitem.ToPane = 44

            Dim txt2 As SAPbouiCOM.EditText
            txt2 = objForm.Items.Item("TXTUUID").Specific


            '=========================================
            ' Tax Amount Label
            '=========================================
            oitem = objForm.Items.Add("LTAX", SAPbouiCOM.BoFormItemTypes.it_STATIC)

            oitem.Top = objForm.Items.Item("LUUID").Top + 22
            oitem.Left = objForm.Items.Item("PSA").Left
            oitem.Width = objForm.Items.Item("PSA").Width
            oitem.Height = objForm.Items.Item("PSA").Height
            oitem.FromPane = 44
            oitem.ToPane = 44

            Dim lbl3 As SAPbouiCOM.StaticText
            lbl3 = objForm.Items.Item("LTAX").Specific
            lbl3.Caption = "Tax Amount"

            '=========================================
            ' Tax Amount Edit Box
            '=========================================
            oitem = objForm.Items.Add("TXTTAX", SAPbouiCOM.BoFormItemTypes.it_EDIT)

            oitem.Top = objForm.Items.Item("LTAX").Top
            oitem.Left = objForm.Items.Item("TXTPSA").Left
            oitem.Width = objForm.Items.Item("TXTPSA").Width
            oitem.Height = objForm.Items.Item("TXTPSA").Height
            oitem.FromPane = 44
            oitem.ToPane = 44

            Dim txt3 As SAPbouiCOM.EditText
            txt3 = objForm.Items.Item("TXTTAX").Specific


            '=========================================
            ' Date Label
            '=========================================
            oitem = objForm.Items.Add("LDATE", SAPbouiCOM.BoFormItemTypes.it_STATIC)

            oitem.Top = objForm.Items.Item("LTAX").Top + 22
            oitem.Left = objForm.Items.Item("PSA").Left
            oitem.Width = objForm.Items.Item("PSA").Width
            oitem.Height = objForm.Items.Item("PSA").Height
            oitem.FromPane = 44
            oitem.ToPane = 44

            Dim lbl4 As SAPbouiCOM.StaticText
            lbl4 = objForm.Items.Item("LDATE").Specific
            lbl4.Caption = "Date"


            oitem = objForm.Items.Add("TXTDATE", SAPbouiCOM.BoFormItemTypes.it_EDIT)

            oitem.Top = objForm.Items.Item("LDATE").Top
            oitem.Left = objForm.Items.Item("TXTPSA").Left
            oitem.Width = objForm.Items.Item("TXTPSA").Width
            oitem.Height = objForm.Items.Item("TXTPSA").Height
            oitem.FromPane = 44
            oitem.ToPane = 44

            Dim txt4 As SAPbouiCOM.EditText
            txt4 = objForm.Items.Item("TXTDATE").Specific

            oitem = objForm.Items.Add("LPAY", SAPbouiCOM.BoFormItemTypes.it_STATIC)

            oitem.Top = objForm.Items.Item("LDATE").Top + 22
            oitem.Left = objForm.Items.Item("PSA").Left
            oitem.Width = objForm.Items.Item("PSA").Width
            oitem.Height = objForm.Items.Item("PSA").Height
            oitem.FromPane = 44
            oitem.ToPane = 44

            Dim lbl5 As SAPbouiCOM.StaticText
            lbl5 = objForm.Items.Item("LPAY").Specific
            lbl5.Caption = "Response"


            oitem = objForm.Items.Add("TXTPAY", SAPbouiCOM.BoFormItemTypes.it_EDIT)

            oitem.Top = objForm.Items.Item("LPAY").Top
            oitem.Left = objForm.Items.Item("TXTPSA").Left

            'Extended Width
            oitem.Width = objForm.Items.Item("TXTPSA").Width + 220

            'Extended Height
            oitem.Height = objForm.Items.Item("TXTPSA").Height + 45

            oitem.FromPane = 44
            oitem.ToPane = 44
            objForm.Items.Item("LPAY").Visible = False
            objForm.Items.Item("TXTPAY").Visible = False

            Dim txt5 As SAPbouiCOM.EditText
            txt5 = objForm.Items.Item("TXTPAY").Specific


            'Posting Status
            objForm.Items.Item("TXTPSA").SetAutoManagedAttribute(
SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1,
SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            'Total Amount
            objForm.Items.Item("TXTAM").SetAutoManagedAttribute(
SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1,
SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            'UUID
            objForm.Items.Item("TXTUUID").SetAutoManagedAttribute(
SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1,
SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            'Tax Amount
            objForm.Items.Item("TXTTAX").SetAutoManagedAttribute(
SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1,
SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            'Date
            objForm.Items.Item("TXTDATE").SetAutoManagedAttribute(
SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1,
SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            'Response
            objForm.Items.Item("TXTPAY").SetAutoManagedAttribute(
SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1,
SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objMain.objUtilities.AddLabel(objForm.UniqueID, "Sts", objForm.Items.Item("86").Top + objForm.Items.Item("86").Height + 5, objForm.Items.Item("86").Left, objForm.Items.Item("86").Width, "E Invoice Status", "86")
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "U_STATUS1", objForm.Items.Item("46").Top + objForm.Items.Item("46").Height + 5, objForm.Items.Item("46").Left, objForm.Items.Item("46").Width, "OINV", "U_STATUS1", "46", objForm.Items.Item("46").Height)

            objForm.Items.Item("U_STATUS1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("U_STATUS1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("U_STATUS1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("U_STATUS1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)


        Catch ex As Exception
            ErrorLog(ex.Message)

            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub E_INVPosting(ByVal INVEntry As String)
        Try
            Dim rs As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            Dim sql As String = "SELECT T1.""Code"" FROM OADM T0  INNER JOIN OCRY T1 ON T0.""Country"" = T1.""Code"""

            rs.DoQuery(sql)
            If rs.RecordCount = 0 Then
                Throw New Exception("No configuration found in E Invoice configuration")
            End If

            Dim Type As String = rs.Fields.Item("Code").Value.ToString()

            If Type = "SA" Then
                Saudi_InvoicePosting(INVEntry)
            ElseIf Type = "AE" Then
                UAE_InvoicePosting(INVEntry)

            Else
                Throw New Exception("Invalid TYPE in @TNX_INVF (1=B2B, 2=B2C)")
            End If

        Catch ex As Exception
            ErrorLog(ex.Message)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub Saudi_InvoicePosting(ByVal INVEntry As String)
        Try
            Dim rs As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            Dim sql As String = "SELECT TOP 1 ""U_TYPE"" FROM ""@TNX_INVF"""

            rs.DoQuery(sql)
            If rs.RecordCount = 0 Then
                Throw New Exception("No configuration found in E Invoice Configuration")
            End If

            Dim Type As String = rs.Fields.Item("U_TYPE").Value.ToString()

            If Type = "2" Then
                E_INVPosting_B2C(INVEntry)
            ElseIf Type = "1" Then
                E_INVPosting_B2B(INVEntry)

            Else
                Throw New Exception("Invalid TYPE in @TNX_INVF (1=B2B, 2=B2C)")
            End If

        Catch ex As Exception
            ErrorLog(ex.Message)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub E_INVPosting_B2C(ByVal INVEntry As String)
        Try

            Dim invQry As String = "Call ""TNX_ZATCA_OINV""('" & INVEntry & "')"
            Dim rsInvoice As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            rsInvoice.DoQuery(invQry)

            Dim dtInvoiceDt As Date = rsInvoice.Fields.Item("invoice_date").Value
            Dim IssueDate As String = dtInvoiceDt.ToString("yyyy-MM-dd")
            Dim IssueTime As String = rsInvoice.Fields.Item("DocTime").Value
            Dim InvDocNum As String = rsInvoice.Fields.Item("invoice_number").Value
            Dim QR As String = rsInvoice.Fields.Item("QRCodeSrc").Value
            Dim InvRef = rsInvoice.Fields.Item("NumAtCard").Value
            Dim CusCode As String = rsInvoice.Fields.Item("CusCode").Value
            Dim CustName As String = rsInvoice.Fields.Item("name").Value
            Dim notes As String = rsInvoice.Fields.Item("notes").Value
            Dim DocEntry As Integer = rsInvoice.Fields.Item("DocEntry").Value
            Dim DocType As String = rsInvoice.Fields.Item("DocType").Value
            Dim InvType As String = rsInvoice.Fields.Item("InvType").Value
            Dim Phone As String = rsInvoice.Fields.Item("phone").Value
            Dim Email As String = rsInvoice.Fields.Item("email").Value
            Dim Country As String = rsInvoice.Fields.Item("country").Value
            Dim CountryCode As String = rsInvoice.Fields.Item("country_code").Value
            Dim PaymentTerms As String = rsInvoice.Fields.Item("payment_terms").Value
            Dim ItemCode As String = rsInvoice.Fields.Item("item_id").Value
            Dim ItemName As String = rsInvoice.Fields.Item("item_name").Value
            Dim Quantity As Decimal = Convert.ToDecimal(rsInvoice.Fields.Item("quantity").Value)
            Dim UnitRate As Decimal = Convert.ToDecimal(rsInvoice.Fields.Item("unit_rate").Value)
            Dim LineTotal As Decimal = Convert.ToDecimal(rsInvoice.Fields.Item("LineTotal").Value)
            Dim VatCategory As String = rsInvoice.Fields.Item("vat_category_code").Value
            Dim VatRate As Decimal = Convert.ToDecimal(rsInvoice.Fields.Item("vat_rate").Value)
            Dim Uom As String = rsInvoice.Fields.Item("unit_of_measure").Value.ToString()
            Dim PaymentMeans As String = rsInvoice.Fields.Item("payment_means").Value.ToString()
            Dim LogoUrl As String = rsInvoice.Fields.Item("logo_url").Value.ToString()
            Dim Type As String = rsInvoice.Fields.Item("type").Value.ToString()
            Dim ItemPriceDiscount As Decimal = Convert.ToDecimal(rsInvoice.Fields.Item("item_price_discount").Value)
            Dim PriceAllowanceIndicator As String = rsInvoice.Fields.Item("price_allowance_indicator").Value.ToString()
            Dim VatExemptionCode As String = rsInvoice.Fields.Item("vat_exemption_code").Value.ToString()
            Dim VatExemptionReason As String = rsInvoice.Fields.Item("vat_exemption_reason").Value.ToString()
            Dim ServiceCode As String = rsInvoice.Fields.Item("service_code").Value.ToString()
            Dim HeaderDiscount As Decimal = Convert.ToDecimal(rsInvoice.Fields.Item("header_discount").Value)
            Dim InvoiceCurrency As String = rsInvoice.Fields.Item("invoice_currency").Value.ToString()
            Dim BusinessKind As String = rsInvoice.Fields.Item("business_kind").Value.ToString()
            Dim InvoiceTransactionCode As String = rsInvoice.Fields.Item("invoice_transaction_code").Value.ToString()

            ' Dim Street As String = rsInvoice.Fields.Item("Street").Value.ToString()

            'Dim DiscountApplicable As String = rsInvoice.Fields.Item("is_discount_applicable").Value.ToString()

            Dim oInv As New FlickINVModel.E_InvoiceRequest()
            oInv.type = Type
            oInv.invoice_date = IssueDate
            oInv.invoice_number = InvDocNum
            oInv.logo_url = LogoUrl
            oInv.payment_terms = PaymentTerms
            oInv.notes = notes
            oInv.payment_means = PaymentMeans
            oInv.header_discount = HeaderDiscount
            oInv.invoice_currency = InvoiceCurrency
            oInv.business_kind = BusinessKind
            oInv.invoice_transaction_code = InvoiceTransactionCode


            'Dim oParty As Party_Details = New Party_Details
            Dim oBilledTo As New FlickINVModel.BilledTo()
            oBilledTo.name = CustName
            oBilledTo.phone = Phone
            oBilledTo.country = Country
            oBilledTo.country_code = CountryCode
            oBilledTo.email = Email


            oInv.billed_to = oBilledTo
            Dim oItems As New List(Of ItemDetail)

            rsInvoice.MoveFirst()
            While Not rsInvoice.EoF

                Dim oItem As New ItemDetail()

                oItem.item_id = ItemCode
                oItem.item_name = ItemName
                oItem.quantity = Quantity
                oItem.unit_rate = UnitRate
                oItem.unit_of_measure = Uom
                oItem.vat_category_code = VatCategory
                oItem.vat_rate = VatRate
                oItem.item_price_discount = ItemPriceDiscount
                oItem.price_allowance_indicator = PriceAllowanceIndicator
                oItem.vat_exemption_code = VatExemptionCode
                oItem.vat_exemption_reason = VatExemptionReason
                oItem.service_code = ServiceCode
                oItem.is_discount_applicable = (oItem.item_price_discount > 0)

                oItems.Add(oItem)

                rsInvoice.MoveNext()
            End While

            oInv.item_details = oItems
            Dim jsonString As String = JsonConvert.SerializeObject(oInv, Newtonsoft.Json.Formatting.Indented)
            Dim rs As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            Dim sql As String = "SELECT ""U_IPPRL"", ""U_BPID"",""U_IQRCG"",""U_APIK"" FROM ""@TNX_INVF"" ORDER BY ""Code"" LIMIT 1"

            rs.DoQuery(sql)


            Dim baseUrl As String = rs.Fields.Item("U_IPPRL").Value.ToString.Trim
            Dim bpid As String = rs.Fields.Item("U_BPID").Value.ToString.Trim
            Dim QRGenerationUrl As String = rs.Fields.Item("U_IQRCG").Value.ToString.Trim
            Dim apiKey As String = rs.Fields.Item("U_APIK").Value.ToString.Trim


            If String.IsNullOrEmpty(baseUrl) Then
                Throw New Exception("U_IPPRL" & " is empty in TNX_INVF")
            End If

            ' Change this if API expects query string instead
            Dim finalUrl As String = baseUrl & "/" & bpid

            Dim apiResult = MakeRequestAsync(finalUrl, jsonString, apiKey).GetAwaiter().GetResult()

            UpdateInvoiceAfterResponse(DocEntry, apiResult)

            LoadAndUpdateInvoiceZatca(DocEntry, QRGenerationUrl, apiKey)

            Dim Test As String = ""

        Catch ex As Exception
            ErrorLog(ex.Message)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub E_INVPosting_B2B(ByVal INVEntry As String)
        Try
            Dim invQry As String = "Call ""TNX_ZATCA_B2B_OINV""('" & INVEntry & "')"
            Dim rsInvoice As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            rsInvoice.DoQuery(invQry)

            Dim dtInvoiceDt As Date = rsInvoice.Fields.Item("invoice_date").Value
            Dim IssueDate As String = dtInvoiceDt.ToString("yyyy-MM-dd")
            Dim dtDueDt As Date = rsInvoice.Fields.Item("due_date").Value
            Dim duedate As String = dtInvoiceDt.ToString("yyyy-MM-dd")
            Dim IssueTime As String = rsInvoice.Fields.Item("DocTime").Value
            Dim InvDocNum As String = rsInvoice.Fields.Item("invoice_number").Value
            Dim QR As String = rsInvoice.Fields.Item("QRCodeSrc").Value
            Dim InvRef = rsInvoice.Fields.Item("NumAtCard").Value
            Dim CusCode As String = rsInvoice.Fields.Item("CusCode").Value
            Dim CustName As String = rsInvoice.Fields.Item("name").Value
            Dim notes As String = rsInvoice.Fields.Item("notes").Value
            Dim DocEntry As Integer = rsInvoice.Fields.Item("DocEntry").Value
            Dim DocType As String = rsInvoice.Fields.Item("DocType").Value
            Dim InvType As String = rsInvoice.Fields.Item("InvType").Value
            Dim Phone As String = rsInvoice.Fields.Item("phone").Value
            Dim Email As String = rsInvoice.Fields.Item("email").Value
            Dim Country As String = rsInvoice.Fields.Item("country").Value
            Dim CountryCode As String = rsInvoice.Fields.Item("country_code").Value
            Dim PaymentTerms As String = rsInvoice.Fields.Item("payment_terms").Value
            Dim ItemCode As String = rsInvoice.Fields.Item("item_id").Value
            Dim ItemName As String = rsInvoice.Fields.Item("item_name").Value
            Dim Quantity As Decimal = Convert.ToDecimal(rsInvoice.Fields.Item("quantity").Value)
            Dim UnitRate As Decimal = Convert.ToDecimal(rsInvoice.Fields.Item("unit_rate").Value)
            Dim LineTotal As Decimal = Convert.ToDecimal(rsInvoice.Fields.Item("LineTotal").Value)
            Dim VatCategory As String = rsInvoice.Fields.Item("vat_category_code").Value
            Dim VatRate As Decimal = Convert.ToDecimal(rsInvoice.Fields.Item("vat_rate").Value)
            Dim Uom As String = rsInvoice.Fields.Item("unit_of_measure").Value.ToString()
            Dim PaymentMeans As String = rsInvoice.Fields.Item("payment_means").Value.ToString()
            Dim LogoUrl As String = rsInvoice.Fields.Item("logo_url").Value.ToString()
            Dim Type As String = rsInvoice.Fields.Item("type").Value.ToString()
            Dim ItemPriceDiscount As Decimal = Convert.ToDecimal(rsInvoice.Fields.Item("item_price_discount").Value)
            Dim PriceAllowanceIndicator As String = rsInvoice.Fields.Item("price_allowance_indicator").Value.ToString()
            Dim VatExemptionCode As String = rsInvoice.Fields.Item("vat_exemption_code").Value.ToString()
            Dim VatExemptionReason As String = rsInvoice.Fields.Item("vat_exemption_reason").Value.ToString()
            Dim ServiceCode As String = rsInvoice.Fields.Item("service_code").Value.ToString()
            Dim HeaderDiscount As Decimal = Convert.ToDecimal(rsInvoice.Fields.Item("header_discount").Value)
            Dim InvoiceCurrency As String = rsInvoice.Fields.Item("invoice_currency").Value.ToString()
            Dim BusinessKind As String = rsInvoice.Fields.Item("business_kind").Value.ToString()
            Dim InvoiceTransactionCode As String = rsInvoice.Fields.Item("invoice_transaction_code").Value.ToString()



            Dim name As String = rsInvoice.Fields.Item("name").Value.ToString()
            Dim company_name As String = rsInvoice.Fields.Item("company_name").Value.ToString()
            Dim logo_url As String = rsInvoice.Fields.Item("logo_url").Value.ToString()
            Dim address1 As String = rsInvoice.Fields.Item("address1").Value.ToString()
            Dim city As String = rsInvoice.Fields.Item("city").Value.ToString()
            Dim district As String = rsInvoice.Fields.Item("district").Value.ToString()
            Dim postal_code As String = rsInvoice.Fields.Item("postal_code").Value.ToString()
            Dim company_reg_number As String = rsInvoice.Fields.Item("company_reg_number").Value.ToString()
            Dim customer_number As String = rsInvoice.Fields.Item("customer_number").Value.ToString()

            Dim gst_no As String = rsInvoice.Fields.Item("gst_no").Value.ToString()
            Dim grp_vat_no As String = rsInvoice.Fields.Item("grp_vat_no").Value.ToString()
            Dim other_id As String = rsInvoice.Fields.Item("other_id").Value.ToString()
            Dim other_id_type As String = rsInvoice.Fields.Item("other_id_type").Value.ToString()

            Dim building_number As String = rsInvoice.Fields.Item("building_number").Value.ToString()
            Dim country_code As String = rsInvoice.Fields.Item("country_code").Value.ToString()
            Dim neighbourhood As String = rsInvoice.Fields.Item("neighbourhood").Value.ToString()

            Dim company_name_local As String = rsInvoice.Fields.Item("company_name_local").Value.ToString()
            Dim iso_certificate_url As String = rsInvoice.Fields.Item("iso_certificate_url").Value.ToString()

            Dim profile_type As String = rsInvoice.Fields.Item("profile_type").Value.ToString()
            Dim is_saudi_vat_registered As Boolean = Convert.ToBoolean(rsInvoice.Fields.Item("is_saudi_vat_registered").Value)

            Dim payment_card_last_four_digits As String = rsInvoice.Fields.Item("payment_card_last_four_digits").Value.ToString()
            Dim address_in_local_language As String = rsInvoice.Fields.Item("address_in_local_language").Value.ToString()
            ' Dim Street As String = rsInvoice.Fields.Item("Street").Value.ToString()

            'Dim DiscountApplicable As String = rsInvoice.Fields.Item("is_discount_applicable").Value.ToString()

            Dim oInv As New FlickINVModel.E_InvoiceRequestB2B()
            oInv.type = Type
            oInv.invoice_date = IssueDate
            oInv.due_date = duedate
            oInv.invoice_number = InvDocNum
            oInv.logo_url = LogoUrl
            oInv.payment_terms = PaymentTerms
            oInv.notes = notes
            oInv.payment_means = PaymentMeans
            oInv.header_discount = HeaderDiscount
            oInv.invoice_currency = InvoiceCurrency
            oInv.business_kind = BusinessKind
            oInv.invoice_transaction_code = InvoiceTransactionCode


            'Dim oParty As Party_Details = New Party_Details
            Dim oBilledTo As New FlickINVModel.BilledToB2B()

            oBilledTo.name = name
            oBilledTo.company_name = company_name
            oBilledTo.logo_url = logo_url
            oBilledTo.email = Email
            oBilledTo.phone = Phone

            oBilledTo.address1 = address1
            oBilledTo.city = city
            oBilledTo.district = district
            oBilledTo.postal_code = postal_code

            oBilledTo.country = Country
            oBilledTo.company_reg_number = company_reg_number
            oBilledTo.customer_number = customer_number

            oBilledTo.gst_no = gst_no
            oBilledTo.grp_vat_no = grp_vat_no
            oBilledTo.other_id = other_id
            oBilledTo.other_id_type = other_id_type

            oBilledTo.building_number = building_number
            oBilledTo.country_code = country_code
            oBilledTo.neighbourhood = neighbourhood

            oBilledTo.company_name_local = company_name_local
            oBilledTo.iso_certificate_url = iso_certificate_url

            oBilledTo.profile_type = profile_type
            oBilledTo.is_saudi_vat_registered = is_saudi_vat_registered

            oBilledTo.payment_card_last_four_digits = payment_card_last_four_digits
            oBilledTo.address_in_local_language = address_in_local_language

            oInv.billed_to = oBilledTo
            Dim oItems As New List(Of ItemDetailB2B)

            rsInvoice.MoveFirst()
            While Not rsInvoice.EoF

                Dim oItem As New ItemDetailB2B()

                oItem.item_id = ItemCode
                oItem.item_name = ItemName
                oItem.quantity = Quantity
                oItem.unit_rate = UnitRate
                oItem.unit_of_measure = Uom
                oItem.vat_category_code = VatCategory
                oItem.vat_rate = VatRate
                oItem.item_price_discount = ItemPriceDiscount
                oItem.price_allowance_indicator = PriceAllowanceIndicator
                oItem.vat_exemption_code = VatExemptionCode
                oItem.vat_exemption_reason = VatExemptionReason
                oItem.service_code = ServiceCode
                oItem.is_discount_applicable = (oItem.item_price_discount > 0)

                oItems.Add(oItem)

                rsInvoice.MoveNext()
            End While

            oInv.item_details = oItems
            Dim jsonString As String = JsonConvert.SerializeObject(oInv, Newtonsoft.Json.Formatting.Indented)
            Dim rs As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            Dim sql As String = "SELECT ""U_IPURLB"", ""U_BPIDB"",""U_IQRCGB"",""U_APIKB"" FROM ""@TNX_INVF"" ORDER BY ""Code"" LIMIT 1"

            rs.DoQuery(sql)


            Dim baseUrl As String = rs.Fields.Item("U_IPURLB").Value.ToString.Trim
            Dim bpid As String = rs.Fields.Item("U_BPIDB").Value.ToString.Trim
            Dim QRGenerationUrl As String = rs.Fields.Item("U_IQRCGB").Value.ToString.Trim
            Dim apiKey As String = rs.Fields.Item("U_APIKB").Value.ToString.Trim


            If String.IsNullOrEmpty(baseUrl) Then
                Throw New Exception("U_IPPRL" & " is empty in TNX_INVF")
            End If

            ' 🔗 Merge BPID with URL
            ' Change this if API expects query string instead
            Dim finalUrl As String = baseUrl & "/" & bpid

            Dim apiResult = MakeRequestAsync(finalUrl, jsonString, apiKey).GetAwaiter().GetResult()

            UpdateInvoiceAfterResponse(DocEntry, apiResult)

            LoadAndUpdateInvoiceZatca(DocEntry, QRGenerationUrl, apiKey)


            Dim Test As String = ""

        Catch ex As Exception
            ErrorLog(ex.Message)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub UpdateInvoiceAfterResponse(CRNEntry As String, apiResult As ApiCallResult)

        Dim oInvoice As SAPbobsCOM.Documents = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices), SAPbobsCOM.Documents)

        oInvoice.GetByKey(Integer.Parse(CRNEntry))

        Dim parsedJson As JToken = JToken.Parse(apiResult.RawResponse)
        Dim formattedJson As String = parsedJson.ToString(Newtonsoft.Json.Formatting.Indented)

        oInvoice.UserFields.Fields.Item("U_RESPONSE").Value = formattedJson


        oInvoice.UserFields.Fields.Item("U_MSG").Value = apiResult.Message

        oInvoice.Update()
        If apiResult.IsSuccess AndAlso apiResult.ParsedResponse IsNot Nothing Then

            If apiResult.ParsedResponse.zatca_response IsNot Nothing AndAlso apiResult.ParsedResponse.zatca_response.error = False Then
                oInvoice.UserFields.Fields.Item("U_STATUS1").Value = apiResult.ParsedResponse.zatca_response.message
            Else
                oInvoice.UserFields.Fields.Item("U_STATUS1").Value = "ZATCA_FAILED"
            End If

            If apiResult.ParsedResponse.details IsNot Nothing Then
                oInvoice.UserFields.Fields.Item("U_UUID").Value = apiResult.ParsedResponse.details.document_id
            End If
            '  Dim status As String = oInvoice.UserFields.Fields.Item("U_STATUS1").Value.ToString().Trim().ToUpper()


            oInvoice.Update()

            objMain.objApplication.StatusBar.SetText("ZATCA Invoice Post Success", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Else
            ErrorLog(apiResult.Message)
            objMain.objApplication.StatusBar.SetText("ZATCA Error: " & apiResult.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If

    End Sub

    Private Sub LoadEInvoiceData(ByVal oForm As SAPbouiCOM.Form)

        Try
            Dim oRs As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim docNum As String = oForm.Items.Item("8").Specific.Value.ToString().Trim()
            Dim series As String = oForm.Items.Item("88").Specific.Value.ToString().Trim()

            Dim query As String = "
        SELECT 
            T0.""U_STATUS1"",
            IFNULL(T0.""DocTotal"",0) AS ""DocTotal"",
            IFNULL(T0.""VatSum"",0) AS ""VatSum"",
            T0.""U_UUID"",
            T0.""DocDate"",
            T0.""U_RESPONSE""
        FROM OINV T0
        WHERE T0.""DocNum"" = '" & docNum & "'
        AND T0.""Series"" = '" & series & "'"

            oRs.DoQuery(query)

            If oRs.RecordCount > 0 Then

                '👉 STOP UI REFRESH
                oForm.Freeze(True)

                oRs.MoveFirst()

                oForm.Items.Item("TXTPSA").Specific.Value =
            oRs.Fields.Item("U_STATUS1").Value.ToString()

                oForm.Items.Item("TXTAM").Specific.Value =
            Convert.ToDecimal(oRs.Fields.Item("DocTotal").Value).ToString("F2")

                oForm.Items.Item("TXTTAX").Specific.Value =
            Convert.ToDecimal(oRs.Fields.Item("VatSum").Value).ToString("F2")

                oForm.Items.Item("TXTUUID").Specific.Value =
            oRs.Fields.Item("U_UUID").Value.ToString()

                If Not IsDBNull(oRs.Fields.Item("DocDate").Value) Then
                    Dim dt As DateTime =
                Convert.ToDateTime(oRs.Fields.Item("DocDate").Value)

                    oForm.Items.Item("TXTDATE").Specific.Value =
                dt.ToString("yyyy-MM-dd")
                End If

                oForm.Items.Item("TXTPAY").Specific.Value =
            oRs.Fields.Item("U_RESPONSE").Value.ToString()

                '👉 RESUME UI REFRESH
                oForm.Freeze(False)

            End If

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText("Load Error: " & ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Public Sub LoadAndUpdateInvoiceZatca(ByVal INVEntry As Integer, ByVal QRGenerationUrl As String, apiKey As String)
        Try
            ' load Invoice
            Dim oInvoice As SAPbobsCOM.Documents = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices), SAPbobsCOM.Documents)

            If Not oInvoice.GetByKey(INVEntry) Then
                objMain.objApplication.StatusBar.SetText("Invoice not found", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            ' Read UUID from OINV
            Dim sUUID As String = oInvoice.UserFields.Fields.Item("U_UUID").Value.ToString().Trim()

            If String.IsNullOrEmpty(sUUID) Then
                objMain.objApplication.StatusBar.SetText("UUID not available in invoice", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            '  Build GET URL
            Dim getUrl As String = QRGenerationUrl & "/" & Uri.EscapeDataString(sUUID)

            ' 4️⃣ Call GET API
            Dim response As GetApiResponse = GetRequestAsync(getUrl, apiKey).GetAwaiter().GetResult()

            If response Is Nothing Then
                objMain.objApplication.StatusBar.SetText("Empty response from ZATCA", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            ' Update Invoice Fields

            ' Save XML → OINV.XMLPath
            If Not String.IsNullOrEmpty(response.xml) Then
                oInvoice.UserFields.Fields.Item("U_XML").Value = response.xml
            End If

            ' Save QR → SAP QR
            If Not String.IsNullOrEmpty(response.qr) Then
                oInvoice.CreateQRCodeFrom = response.qr
                ' Optional backup
                oInvoice.UserFields.Fields.Item("U_QRCODE").Value = response.qr
            End If
            ' 6 Update Invoice
            Dim lRetCode As Integer = oInvoice.Update()

            If lRetCode = 0 Then
                objMain.objApplication.StatusBar.SetText("ZATCA XML & QR updated successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

        Catch ex As Exception
            ErrorLog(ex.Message)
            objMain.objApplication.StatusBar.SetText("ZATCA Update Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    ' Define a class for the JSON response

    Public Async Function MakeRequestAsync(url As String, jsonData As String, apiKey As String) As Task(Of ApiCallResult)

        Using client As New HttpClient()

            client.DefaultRequestHeaders.Clear()
            client.DefaultRequestHeaders.Add("api-key", apiKey)

            Dim content As New StringContent(jsonData, Encoding.UTF8, "application/json")

            Dim httpResponse As HttpResponseMessage = Await client.PostAsync(url, content)

            Dim rawJson As String = Await httpResponse.Content.ReadAsStringAsync()

            Dim msg As String = "API error occurred"
            Dim parsed As ApiResponse = Nothing
            ' 🔹 Extract "message" from JSON (success or error)
            Try
                parsed = JsonConvert.DeserializeObject(Of ApiResponse)(rawJson)
                If parsed IsNot Nothing AndAlso Not String.IsNullOrEmpty(parsed.message) Then
                    msg = parsed.message
                End If
            Catch
                msg = rawJson
            End Try

            Return New ApiCallResult With {
            .RawResponse = rawJson,
            .Message = msg,
            .IsSuccess = httpResponse.IsSuccessStatusCode,
            .ParsedResponse = parsed
        }

        End Using

    End Function


    Public Async Function GetRequestAsync(url As String, apiKey As String) As Task(Of GetApiResponse)
        Using client As New HttpClient()

            client.DefaultRequestHeaders.Add("api-key", apiKey)

            Dim response As HttpResponseMessage = Await client.GetAsync(url)

            response.EnsureSuccessStatusCode()

            Dim jsonResponse As String = Await response.Content.ReadAsStringAsync()
            Dim result As GetApiResponse = JsonConvert.DeserializeObject(Of GetApiResponse)(jsonResponse)

            Return result
        End Using
    End Function
    Private Sub ErrorLog(ByVal errorMsg As String)

        Try
            Dim oTbl As SAPbobsCOM.UserTable = objMain.objCompany.UserTables.Item("TNX_ERRORLOGS")
            ' Get Next Code (MAX + 1)
            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRs.DoQuery("SELECT IFNULL(MAX(CAST(""Code"" AS INTEGER)), 0) + 1 AS ""NextCode"" FROM ""@TNX_ERRORLOGS""")

            Dim nextCode As String = oRs.Fields.Item("NextCode").Value.ToString()

            ' Mandatory Keys

            oTbl.Code = nextCode
            oTbl.Name = nextCode
            ' Date & Time
            oTbl.UserFields.Fields.Item("U_DATE").Value = Date.Now
            oTbl.UserFields.Fields.Item("U_TIME").Value = Date.Now.ToString("HHmmss")

            ' Error Message
            oTbl.UserFields.Fields.Item("U_ERROR_MSG").Value =
            errorMsg.Substring(0, Math.Min(254, errorMsg.Length))

            oTbl.Add()

        Catch
        End Try

    End Sub
    Private Sub UpdateInvoiceDateTime(ByVal postStatus As String,
                                  ByVal oSource As SAPbouiCOM.Form)

        Try

            If UCase(postStatus).Contains("SUCCESS") _
        Or UCase(postStatus).Contains("FAILED") Then

                Dim oInv As SAPbobsCOM.Documents
                Dim docEntry As Integer
                Dim ret As Integer

                docEntry = CInt(oSource.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).Trim())

                oInv = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

                If oInv.GetByKey(docEntry) Then

                    oInv.UserFields.Fields.Item("U_DATETIME").Value = Now

                    ret = oInv.Update()

                    If ret = 0 Then
                        objMain.objApplication.StatusBar.SetText(
                    "Date Time Updated",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    End If

                End If

            End If

        Catch ex As Exception

        End Try

    End Sub
    Sub UAE_InvoicePosting(ByVal INVEntry As String)

        Try
            ' 🔷 Get API Config
            Dim result As SAPbobsCOM.Recordset =
            DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            Dim sql As String = "SELECT ""U_AIPPRL"", ""U_ALIQRCG"",""U_ALCMPRL"",""U_AIUUID"",""U_ACMQRC"" FROM ""@TNX_INVF"" ORDER BY ""Code"" LIMIT 1"
            result.DoQuery(sql)


            Dim baseUrl As String = result.Fields.Item("U_AIPPRL").Value.ToString.Trim
            Dim username As String = result.Fields.Item("U_ALCMPRL").Value.ToString.Trim
            Dim postUrl As String = result.Fields.Item("U_ALIQRCG").Value.ToString.Trim
            Dim password As String = result.Fields.Item("U_ACMQRC").Value.ToString.Trim
            Dim getUrl As String = result.Fields.Item("U_AIUUID").Value.ToString.Trim
            Dim token As String = GetAccessToken(baseUrl, username, password)

            ' 🔷 Get Invoice Data
            Dim invQry As String = "Call ""TNX_E_OINV""('" & INVEntry & "')"
            Dim rs As SAPbobsCOM.Recordset =
            DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            rs.DoQuery(invQry)

            Dim oInv As New InvoiceModel()

            ' ================= HEADER =================
            Dim ReferenceNumber As String = rs.Fields.Item("ReferenceNumber").Value.ToString()
            Dim FinancialYear As String = rs.Fields.Item("FinancialYear").Value.ToString()
            Dim DocEntry As String = rs.Fields.Item("DocEntry").Value.ToString()


            oInv.ReferenceNumber = rs.Fields.Item("ReferenceNumber").Value.ToString()
            oInv.FinancialYear = rs.Fields.Item("FinancialYear").Value.ToString()
            oInv.edelivery = rs.Fields.Item("edelivery").Value.ToString()
            oInv.InvoiceID = rs.Fields.Item("InvoiceID").Value.ToString()
            oInv.IssueDate = Convert.ToDateTime(rs.Fields.Item("IssueDate").Value).ToString("yyyy-MM-dd")
            oInv.DueDate = Convert.ToDateTime(rs.Fields.Item("DueDate").Value).ToString("yyyy-MM-dd")
            oInv.InvoiceTypeCode = rs.Fields.Item("InvoiceTypeCode").Value.ToString()
            oInv.Note = rs.Fields.Item("Note").Value.ToString()
            oInv.DocumentCurrencyCode = rs.Fields.Item("DocumentCurrencyCode").Value.ToString()
            oInv.VATCurrencyCode = rs.Fields.Item("VATCurrencyCode").Value.ToString()
            oInv.AccountingCost = rs.Fields.Item("AccountingCost").Value.ToString()
            oInv.BuyerReference = rs.Fields.Item("BuyerReference").Value.ToString()
            oInv.VATtaxpointdate = Convert.ToDateTime(rs.Fields.Item("VATtaxpointdate").Value).ToString("yyyy-MM-dd")
            oInv.PrincipleID = rs.Fields.Item("PrincipleID").Value.ToString()
            oInv.BeneficiaryID = rs.Fields.Item("BeneficiaryID").Value.ToString()
            oInv.Currencyexchangerate = rs.Fields.Item("Currencyexchangerate").Value.ToString()
            oInv.Purchaseorderreference = rs.Fields.Item("Purchaseorderreference").Value.ToString()
            oInv.Salesorderreference = rs.Fields.Item("Salesorderreference").Value.ToString()
            oInv.Despatchadvicereference = rs.Fields.Item("Despatchadvicereference").Value.ToString()
            oInv.Receivingadvicereference = rs.Fields.Item("Receivingadvicereference").Value.ToString()
            oInv.Customsreferencenumber = rs.Fields.Item("Customsreferencenumber").Value.ToString()
            oInv.FrequencyofBilling = rs.Fields.Item("FrequencyofBilling").Value.ToString()

            ' ================= INVOICE PERIOD =================
            oInv.InvoicePeriod = New InvoicePeriodModel With {
    .StartDate = Convert.ToDateTime(rs.Fields.Item("StartDate").Value).ToString("yyyy-MM-dd"),
    .EndDate = Convert.ToDateTime(rs.Fields.Item("EndDate").Value).ToString("yyyy-MM-dd")
}

            ' ================= BILLING =================
            oInv.BillingReference = New List(Of BillingReferenceModel) From {
    New BillingReferenceModel With {
        .ID = rs.Fields.Item("RefInvoiceID").Value.ToString(),
        .IssueDate = Convert.ToDateTime(rs.Fields.Item("RefIssueDate").Value).ToString("yyyy-MM-dd"),
        .Creditnotereasoncode = rs.Fields.Item("CreditReason").Value.ToString()
    }
}

            ' ================= ADDITIONAL DOC =================
            oInv.AdditionalDocumentDetails = New List(Of AdditionalDocumentModel) From {
    New AdditionalDocumentModel With {
        .AdditionalDocumentReferenceID = rs.Fields.Item("AdditionalDocumentReferenceID").Value.ToString(),
        .SchemeID = rs.Fields.Item("SchemeID").Value.ToString(),
        .DocumentTypecode = rs.Fields.Item("DocumentTypecode").Value.ToString(),
        .DocumentDescription = rs.Fields.Item("DocumentDescription").Value.ToString(),
        .EmbeddedDocumentBinaryObject = rs.Fields.Item("EmbeddedDocumentBinaryObject").Value.ToString(),
        .mimeCode = rs.Fields.Item("mimeCode").Value.ToString(),
        .FileName = rs.Fields.Item("FileName").Value.ToString(),
        .ExternalReferenceID = rs.Fields.Item("ExternalReferenceID").Value.ToString()
    }
}

            ' ================= SUPPLIER =================
            oInv.AccountingSupplierParty = New AccountingSupplierPartyModel With {
    .SellerCode = rs.Fields.Item("SellerCode").Value.ToString(),
    .EndpointID = New SupplierEndpointModel With {
        .schemeID = rs.Fields.Item("SupplierSchemeID").Value.ToString(),
        .SellerElectronicAddress = rs.Fields.Item("SellerElectronicAddress").Value.ToString()
    },
    .PartyIdentification = New List(Of SupplierPartyIdentificationModel) From {
        New SupplierPartyIdentificationModel With {
            .SellerLegalSchemeIdentifier = rs.Fields.Item("SellerLegalSchemeIdentifier").Value.ToString(),
            .SellerLegalRegistrationType = rs.Fields.Item("SellerLegalRegistrationType").Value.ToString(),
            .SellerCommercial_Tradelicense = rs.Fields.Item("SellerTradeLicense").Value.ToString(),
            .SellerEmiratesID = rs.Fields.Item("SellerEmiratesID").Value.ToString(),
            .SellerPassport = rs.Fields.Item("SellerPassport").Value.ToString(),
            .SellerCabinetDecision = rs.Fields.Item("SellerCabinetDecision").Value.ToString(),
            .SellerAuthorityname = rs.Fields.Item("SellerAuthorityname").Value.ToString(),
            .SellerPassportIssuingCountrycode = rs.Fields.Item("SellerPassportCountry").Value.ToString()
        }
    },
    .PostalAddress = New SupplierAddressModel With {
        .StreetName = rs.Fields.Item("SupplierStreet").Value.ToString(),
        .AdditionalStreetName = rs.Fields.Item("SupplierStreet2").Value.ToString(),
        .CityName = rs.Fields.Item("SupplierCity").Value.ToString(),
        .CitySubdivisionName = rs.Fields.Item("SupplierArea").Value.ToString(),
        .PostalZone = rs.Fields.Item("SupplierZip").Value.ToString(),
        .Country = New SupplierCountryModel With {
            .IdentificationCode = rs.Fields.Item("SupplierCountry").Value.ToString()
        }
    },
    .TaxScheme = New SupplierTaxModel With {
        .SellerTaxidentifier = rs.Fields.Item("SupplierVAT").Value.ToString(),
        .ID = rs.Fields.Item("SupplierTaxScheme").Value.ToString()
    },
    .PartyLegalEntity = New SupplierLegalModel With {
        .RegistrationName = rs.Fields.Item("SupplierName").Value.ToString(),
        .CompanyID = rs.Fields.Item("SupplierCompanyID").Value.ToString()
    },
    .Contact = New SupplierContactModel With {
        .Name = rs.Fields.Item("SupplierContactName").Value.ToString(),
        .Telephone = rs.Fields.Item("SupplierPhone").Value.ToString(),
        .EmailID = rs.Fields.Item("SupplierEmail").Value.ToString()
    }
}

            ' ================= CUSTOMER =================
            oInv.AccountingCustomerParty = New AccountingCustomerPartyModel With {
                .BuyerCode = rs.Fields.Item("BuyerCode").Value.ToString(),
                .EndpointID = New CustomerEndpointModel With {
                    .SchemeID = rs.Fields.Item("CustomerSchemeID").Value.ToString(),
                    .BuyerElectronicAddress = rs.Fields.Item("CustomerVAT").Value.ToString()
                },
                .PartyIdentification = New CustomerPartyIdentificationModel With {
                    .ID = rs.Fields.Item("CustomerVAT").Value.ToString(),
                    .BuyerSchemeidentifier = rs.Fields.Item("CustomerSchemeID").Value.ToString()
                },
                .PostalAddress = New CustomerAddressModel With {
                    .StreetName = rs.Fields.Item("CustStreet").Value.ToString(),
                    .AdditionalStreetName = rs.Fields.Item("CustStreet2").Value.ToString(),
                    .CityName = rs.Fields.Item("CustCity").Value.ToString(),
                    .CitySubdivisionName = rs.Fields.Item("CustArea").Value.ToString(),
                    .PostalZone = rs.Fields.Item("CustZip").Value.ToString(),
                    .Country = New CountryModel With {
                        .IdentificationCode = rs.Fields.Item("CustCountry").Value.ToString()
                    }
                },
                 .PartyTaxScheme = New CustomerTaxModel With {
                    .CompanyID = rs.Fields.Item("CustomerVAT").Value.ToString(),
                    .TaxScheme = New TaxSchemeModel With {
                        .ID = "VAT",
                        .Buyerlegalregistrationidentifiertype = rs.Fields.Item("Buyerlegalregistrationidentifiertype").Value.ToString(),
                        .BuyerCommercialorTradelicense = rs.Fields.Item("BuyerCommercialorTradelicense").Value.ToString(),
                        .BuyerEmiratesID = rs.Fields.Item("BuyerEmiratesID").Value.ToString(),
                        .BuyerPassport = rs.Fields.Item("BuyerPassport").Value.ToString(),
                        .BuyerCabinetDecision = rs.Fields.Item("BuyerCabinetDecision").Value.ToString(),
                        .BuyerAuthorityName = rs.Fields.Item("BuyerAuthorityName").Value.ToString(),
                        .BuyerPassportIssuingCountrycode = rs.Fields.Item("BuyerPassportIssuingCountrycode").Value.ToString()
                    },
                    .PartyLegalEntity = New CustomerLegalModel With {
                        .RegistrationName = rs.Fields.Item("CustomerName").Value.ToString(),
                        .CompanyID = New CustomerCompanyModel With {
                            .schemeID = rs.Fields.Item("CustomerSchemeID").Value.ToString(),
                            .RegistrationIdentifier = rs.Fields.Item("CustomerVAT").Value.ToString()
                        }
                    },
                    .Contact = New CustomerContactModel With {
                        .Name = rs.Fields.Item("ContactPerson").Value.ToString(),
                        .Telephone = rs.Fields.Item("Phone").Value.ToString(),
                        .ElectronicEmail = rs.Fields.Item("Email").Value.ToString()
                    }
                }
            }

            ' ================= DELIVERY =================
            oInv.Delivery = New DeliveryModel With {
                .ActualDeliveryDate = Convert.ToDateTime(rs.Fields.Item("DeliveryDate").Value).ToString("yyyy-MM-dd"),
                .DeliveryLocation = New DeliveryLocationModel With {
                    .Name = rs.Fields.Item("DeliveryName").Value.ToString(),
                    .ID = New DeliveryIDModel With {
                        .schemeID = rs.Fields.Item("DeliverySchemeID").Value.ToString(),
                        .DeliveryIdentifier = rs.Fields.Item("DeliveryCode").Value.ToString()
                    },
                    .PostalAddress = New DeliveryAddressModel With {
                        .StreetName = rs.Fields.Item("DeliveryStreet").Value.ToString(),
                        .AdditionalStreetName = rs.Fields.Item("DeliveryStreet2").Value.ToString(),
                        .CityName = rs.Fields.Item("DeliveryCity").Value.ToString(),
                        .PostalZone = rs.Fields.Item("DeliveryZip").Value.ToString(),
                        .Country = New DeliveryCountryModel With {
                            .IdentificationCode = rs.Fields.Item("DeliveryCountry").Value.ToString()
                        }
                    }
                },
                .DeliveryParty = New DeliveryPartyModel With {
                    .PartyName = rs.Fields.Item("DeliveryName").Value.ToString()
                }
            }

            ' ================= ALLOWANCE =================
            oInv.AllowanceCharge = New List(Of AllowanceChargeModel) From {
                New AllowanceChargeModel With {
                    .AllowanceID = rs.Fields.Item("AllowanceID").Value.ToString(),
                    .ChargeIndicator = Convert.ToBoolean(rs.Fields.Item("ChargeIndicator").Value),
                    .AllowanceChargeReasonCode = rs.Fields.Item("AllowanceReason").Value.ToString(),
                    .MultiplierFactorNumeric = rs.Fields.Item("Multiplier").Value.ToString(),
                    .Amount = New AmountModel With {
                        .AllowanceAmount = rs.Fields.Item("AllowanceAmount").Value.ToString()
                        },
                    .BaseAmount = New BaseAmountModel With {
                        .AllowanceDiscountBaseAmount = rs.Fields.Item("BaseAmount").Value.ToString()
                    },
                    .TaxCategory = New TaxCategoryModel With {
                    .ID = rs.Fields.Item("TaxCategoryID").Value.ToString(),      ' e.g. S
                    .Percent = rs.Fields.Item("TaxPercent").Value.ToString()    ' e.g. 10                   
                  }
                }
            }

            ' ================= PAYMENT =================
            oInv.PaymentMeans = New PaymentMeansModel With {
                .PayeeName = rs.Fields.Item("PayeeName").Value.ToString(),
                .PaymentInstructions = rs.Fields.Item("PaymentInstructions").Value.ToString(),
                .PaymentDueDate = Convert.ToDateTime(rs.Fields.Item("PaymentDueDate").Value).ToString("yyyy-MM-dd"),
                .PaymentMeansCode = rs.Fields.Item("PaymentMeansCode").Value.ToString(),
                .PayeeFinancialAccount = New FinancialAccountModel With {
                    .ID = rs.Fields.Item("AccountID").Value.ToString(),
                    .Name = rs.Fields.Item("AccountName").Value.ToString(),
                    .FinancialBranch = New BranchModel With {
                        .ID = rs.Fields.Item("BranchID").Value.ToString()
                    }
                }
            }

            ' ================= PAYMENT TERMS =================
            oInv.PaymentTerms = New PaymentTermsModel With {
                .Note = rs.Fields.Item("PaymentTerms").Value.ToString(),
                .PayeeAccountID = rs.Fields.Item("PayeeAccountID").Value.ToString()
            }


            oInv.PaymentTerms = New PaymentTermsModel With {
            .Note = rs.Fields.Item("PaymentTerms").Value.ToString()
        }

            ' ================= TOTAL =================
            oInv.LegalMonetaryTotal = New TotalModel With {
                .LineExtensionAmount = rs.Fields.Item("LineExtensionAmount").Value.ToString(),
                .TaxExclusiveAmount = rs.Fields.Item("TaxExclusiveAmount").Value.ToString(),
                .TaxInclusiveAmount = rs.Fields.Item("TaxInclusiveAmount").Value.ToString(),
                .ChargeTotalAmount = rs.Fields.Item("ChargeTotalAmount").Value.ToString(),
                .PrepaidAmount = rs.Fields.Item("PrepaidAmount").Value.ToString(),
                .PayableRoundingAmount = rs.Fields.Item("PayableRoundingAmount").Value.ToString(),
                .PayableAmount = rs.Fields.Item("PayableAmount").Value.ToString(),
                .TotalDiscountAmount = rs.Fields.Item("TotalDiscountAmount").Value.ToString()
            }

            ' ================= LINES =================
            ' ================= LINES =================
            Dim oLines As New List(Of InvoiceLineModel)

            If Not rs.EoF Then
                rs.MoveFirst()

                While Not rs.EoF

                    ' ===== Create Line =====
                    Dim line As New InvoiceLineModel

                    ' ===== Basic Details =====
                    line.InvoiceLineIdentifier = rs.Fields.Item("LineID").Value.ToString()

                    ' ===== Quantity =====
                    line.InvoicedQuantity = New InvoicedQuantityModel With {
            .unitCode = rs.Fields.Item("UomCode").Value.ToString(),
            .Quantity = rs.Fields.Item("Quantity").Value.ToString()
        }

                    ' ===== Line Amount =====
                    line.LineExtensionAmount = New LineExtensionAmountModel With {
            .TaxExclusiveAmount = rs.Fields.Item("LineTotal").Value.ToString()
        }

                    ' ===== Line Item =====
                    Dim oItem As New LineItemModel

                    oItem.Description = rs.Fields.Item("Dscription").Value.ToString()
                    oItem.Name = rs.Fields.Item("ItemName").Value.ToString()
                    oItem.Typeofgoodsorservices = rs.Fields.Item("GoodsType").Value.ToString()
                    oItem.ServiceAccountingcode = rs.Fields.Item("ServiceCode").Value.ToString()
                    oItem.ServiceAccountingSchemeidentifier = rs.Fields.Item("ServiceAccountingSchemeidentifier").Value.ToString()
                    oItem.Itemtype = rs.Fields.Item("ItemType").Value.ToString()

                    ' ===== Standard Item =====
                    oItem.StandardItemIdentification = New StandardItemIdentificationModel With {
            .schemeID = rs.Fields.Item("StandardSchemeID").Value.ToString(),
            .SchemeIDCode = rs.Fields.Item("StandardSchemeCode").Value.ToString()
        }

                    ' ===== Commodity =====
                    oItem.CommodityClassification = New CommodityClassificationModel With {
            .listID = rs.Fields.Item("CommodityListID").Value.ToString(),
            .Itemcode = rs.Fields.Item("CommodityCode").Value.ToString()
        }

                    ' ===== Tax =====
                    oItem.ClassifiedTaxCategory = New TaxCategoryModel With {
            .ID = rs.Fields.Item("TaxCode").Value.ToString(),
            .Percent = rs.Fields.Item("TaxPercent").Value.ToString(),
            .TaxSchemeId = rs.Fields.Item("TaxSchemeId").Value.ToString(),
            .TaxExemptionReasonCode = rs.Fields.Item("TaxExemptionReasonCode").Value.ToString(),
            .TaxExemptionReason = rs.Fields.Item("TaxExemptionReason").Value.ToString()
        }

                    ' ===== Allowance / Charge =====
                    Dim oAllowance As New LineAllowanceChargeModel

                    oAllowance.LineAllowanceID = rs.Fields.Item("LineID").Value.ToString()
                    oAllowance.ChargeIndicator = rs.Fields.Item("LineChargeIndicator").Value.ToString()
                    oAllowance.AllowanceChargeReasonCode = rs.Fields.Item("AllowanceReason").Value.ToString()
                    oAllowance.MultiplierFactorNumeric = rs.Fields.Item("LineMultiplier").Value.ToString()

                    oAllowance.Amount = New LineAllowanceAmountModel With {
            .CurrencyID = rs.Fields.Item("LineCurrencyID").Value.ToString(),
            .AllowanceAmount = rs.Fields.Item("LineAllowanceAmount").Value.ToString()
        }

                    oAllowance.BaseAmount = New LineAllowanceBaseModel With {
            .CurrencyID = rs.Fields.Item("LineBaseCurrencyID").Value.ToString(),
            .AllowanceChargeBaseAmount = rs.Fields.Item("LineBaseAmount").Value.ToString()
        }

                    oItem.LineAllowanceCharge = New List(Of LineAllowanceChargeModel)
                    oItem.LineAllowanceCharge.Add(oAllowance)

                    ' ===== Price =====
                    oItem.Price = New PriceModel With {
            .BaseQuantity = rs.Fields.Item("BaseQuantity").Value.ToString(),
            .BaseQuantityUOM = rs.Fields.Item("BaseQuantityUOM").Value.ToString(),
            .PriceAmount = New PriceAmountModel With {
                .Amount = rs.Fields.Item("Price").Value.ToString()
            },
            .AlwChg = New AlwChgModel With {
                .Amt = rs.Fields.Item("AlwChgAmt").Value.ToString(),
                .BaseAmt = rs.Fields.Item("AlwChgBaseAmt").Value.ToString()
            }
        }

                    ' ===== Assign Item to Line =====
                    line.LineItem = oItem

                    ' ===== Add Line =====
                    oLines.Add(line)

                    rs.MoveNext()
                End While
            End If



            oInv.InvoiceLine = oLines

            ' ================= JSON =================
            Dim jsonString As String = JsonConvert.SerializeObject(oInv, Newtonsoft.Json.Formatting.Indented)
            Dim rsUpd As SAPbobsCOM.Recordset
            rsUpd = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            Dim safeJson As String = jsonString.Replace("'", "''")

            Dim sqlUpd As String =
            "UPDATE ""OINV"" SET ""U_ARQT"" = '" & safeJson & "' WHERE ""DocEntry"" = '" & DocEntry & "'"

            rsUpd.DoQuery(sqlUpd)
            Dim response As String = PostInvoiceData(postUrl, token, jsonString)
            System.Threading.Thread.Sleep(5000) ' 5 seconds
            Dim resultUUID As String = GetInvoiceData(getUrl, token, ReferenceNumber, "380", FinancialYear)

            ProcessInvoiceResponse(resultUUID, DocEntry)



        Catch ex As Exception
            ErrorLog(ex.Message)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub

    Private Sub ProcessInvoiceResponse(ByVal resultUUID As String, ByVal DocEntry As String)

        Try
            ' 🔹 Deserialize
            Dim obj As InvoiceResponse = Newtonsoft.Json.JsonConvert.DeserializeObject(Of InvoiceResponse)(resultUUID)

            ' 🔹 Format JSON
            Dim jsonFormatted As String = Newtonsoft.Json.JsonConvert.SerializeObject(
            Newtonsoft.Json.JsonConvert.DeserializeObject(resultUUID),
            Newtonsoft.Json.Formatting.Indented)

            ' 🔹 Get SAP Invoice Object
            Dim oInvoice As SAPbobsCOM.Documents = CType(
            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices),
            SAPbobsCOM.Documents)

            If oInvoice.GetByKey(Integer.Parse(DocEntry)) Then

                If obj IsNot Nothing AndAlso obj.Invoice IsNot Nothing Then

                    ' ================= SUCCESS =================
                    If Not String.IsNullOrEmpty(obj.Invoice.UUID) Then

                        oInvoice.UserFields.Fields.Item("U_UUID").Value = obj.Invoice.UUID
                        oInvoice.UserFields.Fields.Item("U_STATUS1").Value = "SUCCESS"
                        oInvoice.UserFields.Fields.Item("U_EIDate").Value = Date.Today
                        oInvoice.UserFields.Fields.Item("U_EITime").Value = Format(Now, "HHmmss")
                        ' 🔹 Clear Error Safely
                        Dim errField = oInvoice.UserFields.Fields.Item("U_ERRORMSG")
                        errField.Value = " "
                        errField.Value = ""

                        oInvoice.UserFields.Fields.Item("U_RESPONSE").Value = jsonFormatted
                        Dim ret As Integer = oInvoice.Update()

                        If ret = 0 Then
                            objMain.objApplication.StatusBar.SetText(
                                "EInvoice Submitted Successfully. UUID: " & obj.Invoice.UUID,
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        Else
                            Dim errMsg As String = objMain.objCompany.GetLastErrorDescription()
                            objMain.objApplication.StatusBar.SetText(
                                "SAP Update Error: " & errMsg,
                                SAPbouiCOM.BoMessageTime.bmt_Long,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If

                        ' ================= FAILED =================
                    Else
                        oInvoice.UserFields.Fields.Item("U_STATUS1").Value = "FAILED"
                        oInvoice.UserFields.Fields.Item("U_EIDate").Value = Date.Today
                        oInvoice.UserFields.Fields.Item("U_EITime").Value = Format(Now, "HHmmss")
                        oInvoice.UserFields.Fields.Item("U_ERRORMSG").Value = obj.Invoice.ErrorMessage
                        oInvoice.UserFields.Fields.Item("U_RESPONSE").Value = jsonFormatted
                        Dim ret As Integer = oInvoice.Update()

                        If ret = 0 Then
                            objMain.objApplication.StatusBar.SetText("E Invoice Failed: " & obj.Invoice.ErrorMessage,)

                        Else
                            Dim errMsg As String = objMain.objCompany.GetLastErrorDescription()
                            objMain.objApplication.StatusBar.SetText(
                                "SAP Update Error: " & errMsg,
                                SAPbouiCOM.BoMessageTime.bmt_Long,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    End If



                Else
                    objMain.objApplication.StatusBar.SetText("Invalid API Response")
                End If

            Else
                objMain.objApplication.StatusBar.SetText("Invoice Not Found in SAP")
            End If

        Catch ex As Exception
            ErrorLog(ex.Message)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub  'post 
    Private Function PostInvoiceData(ByVal postUrl As String,
                                 ByVal accessToken As String,
                                 ByVal jsonBody As String) As String
        Try
            ' 🔹 Create Request
            Dim request As HttpWebRequest = CType(WebRequest.Create(postUrl), HttpWebRequest)
            request.Method = "POST"
            request.ContentType = "application/json"

            ' 🔹 Add Bearer Token
            request.Headers.Add("Authorization", "Bearer " & accessToken)

            ' 🔹 Write JSON Body
            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(jsonBody)
            request.ContentLength = byteArray.Length

            Using dataStream As Stream = request.GetRequestStream()
                dataStream.Write(byteArray, 0, byteArray.Length)
            End Using

            ' 🔹 Get Response
            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)

            Dim responseString As String = ""
            Using reader As New StreamReader(response.GetResponseStream())
                responseString = reader.ReadToEnd()
            End Using

            Return responseString

        Catch ex As Exception
            Throw New Exception("Post API Error: " & ex.Message)
        End Try
    End Function
    'access token
    Private Function GetAccessToken(ByVal baseUrl As String, ByVal username As String, ByVal password As String) As String
        Try
            ' 🔹 API URL
            Dim tokenUrl As String = baseUrl

            ' 🔹 Create Request
            Dim request As HttpWebRequest = CType(WebRequest.Create(tokenUrl), HttpWebRequest)
            request.Method = "POST"
            request.ContentType = "application/x-www-form-urlencoded"

            ' 🔹 Basic Auth Header
            Dim authValue As String = username & ":" & password
            Dim authBytes As Byte() = Encoding.UTF8.GetBytes(authValue)
            Dim base64Auth As String = Convert.ToBase64String(authBytes)
            request.Headers.Add("Authorization", "Basic " & base64Auth)

            ' 🔹 Grant Type (usually still required)
            Dim postData As String = "grant_type=client_credentials"

            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(postData)
            request.ContentLength = byteArray.Length

            Using dataStream As Stream = request.GetRequestStream()
                dataStream.Write(byteArray, 0, byteArray.Length)
            End Using

            ' 🔹 Get Response
            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)

            Dim responseString As String = ""
            Using reader As New StreamReader(response.GetResponseStream())
                responseString = reader.ReadToEnd()
            End Using

            ' 🔹 Parse JSON
            Dim json As JObject = JObject.Parse(responseString)

            Dim accessToken As String = json("access_token").ToString()
            Dim tokenType As String = json("token_type").ToString()
            Dim expiresIn As String = json("expires_in").ToString()
            Dim orgId As String = json("orgId").ToString()

            Return accessToken

        Catch ex As Exception
            Throw New Exception("Error while getting token: " & ex.Message)
        End Try
    End Function
    Private Function GetInvoiceData(ByVal getUrl As String,
                                ByVal accessToken As String,
                                ByVal referenceNumber As String,
                                ByVal invoiceTypeCode As String,
                                ByVal financialYear As String) As String
        Try
            ' 🔹 Build URL with Query Params
            Dim fullUrl As String = getUrl & "?referencenumber=" & referenceNumber &
                               "&InvoiceTypeCode=" & invoiceTypeCode &
                               "&financialyear=" & financialYear

            ' 🔹 Create Request
            Dim request As HttpWebRequest = CType(WebRequest.Create(fullUrl), HttpWebRequest)
            request.Method = "GET"
            request.ContentType = "application/json"

            ' 🔹 Add Bearer Token
            request.Headers.Add("Authorization", "Bearer " & accessToken)

            ' 🔹 Get Response
            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)

            Dim responseString As String = ""
            Using reader As New StreamReader(response.GetResponseStream())
                responseString = reader.ReadToEnd()
            End Using

            Return responseString

        Catch ex As Exception
            Throw New Exception("GET API Error: " & ex.Message)
        End Try
    End Function


    Public Class ApiResponse
        Public Property message As String
        Public Property details As ApiDetails
        Public Property zatca_response As ZatcaResponse
    End Class

    Public Class ApiDetails
        Public Property document_id As String
        Public Property doc_type As String
        Public Property view_type As String
    End Class

    Public Class ZatcaResponse
        Public Property message As String
        Public Property [error] As Boolean
    End Class
    Public Class GetApiResponse
        Public Property xml As String
        Public Property qr As String
        Public Property message As String
    End Class
    Public Class ApiCallResult
        Public Property RawResponse As String
        Public Property Message As String
        Public Property IsSuccess As Boolean
        Public Property ParsedResponse As ApiResponse   ' ⭐ ADD THIS

    End Class




End Class

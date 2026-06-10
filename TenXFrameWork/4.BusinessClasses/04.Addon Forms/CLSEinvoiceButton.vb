'Imports SAPbouiCOM
'Imports SAPbobsCOM

'Public Class CLSEinvoiceButton

'#Region "Declaration"

'    Public objForm As SAPbouiCOM.Form
'    Private oItem As SAPbouiCOM.Item
'    Private oButton As SAPbouiCOM.Button

'#End Region

'#Region "Main Item Event"

'    Public Sub ItemEvent(ByVal FormUID As String,
'                         ByRef pVal As SAPbouiCOM.ItemEvent,
'                         ByRef BubbleEvent As Boolean)

'        Try

'            'A/R Invoice Standard Form
'            If pVal.FormTypeEx = "133" Then

'                Select Case pVal.EventType

'                    'After form load
'                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

'                        If pVal.BeforeAction = False Then
'                            AddEInvoiceButton(FormUID)
'                        End If

'                    'Button Click
'                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

'                        If pVal.BeforeAction = False _
'                        AndAlso pVal.ItemUID = "btnEInv" Then

'                            objForm = objMain.objApplication.Forms.Item(FormUID)

'                            Dim DocNum As String = ""
'                            DocNum = CType(objForm.Items.Item("8").Specific,
'                                           SAPbouiCOM.EditText).Value

'                            objMain.objApplication.StatusBar.SetText(
'                            "E-Invoice Reconciliation Clicked : " & DocNum,
'                            SAPbouiCOM.BoMessageTime.bmt_Short,
'                            SAPbouiCOM.BoStatusBarMessageType.smt_Success)

'                            '=================================================
'                            'Write your reconciliation logic here
'                            '=================================================

'                        End If

'                End Select

'            End If

'        Catch ex As Exception

'            objMain.objApplication.StatusBar.SetText(
'            "ItemEvent Error : " & ex.Message,
'            SAPbouiCOM.BoMessageTime.bmt_Short,
'            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

'        End Try

'    End Sub

'#End Region

'#Region "Add Button"

'    Private Sub AddEInvoiceButton(ByVal FormUID As String)

'        Try

'            objForm = objMain.objApplication.Forms.Item(FormUID)

'            'Prevent duplicate button
'            If ItemExists(objForm, "btnEInv") = False Then

'                objForm.Freeze(True)

'                oItem = objForm.Items.Add("btnEInv",
'                        SAPbouiCOM.BoFormItemTypes.it_BUTTON)

'                'Use Cancel button (UID = 2) as reference
'                oItem.Top = objForm.Items.Item("Split").Top
'                oItem.Left = objForm.Items.Item("Split").Left - 170
'                oItem.Width = 160
'                oItem.Height = objForm.Items.Item("Split").Height

'                oItem.FromPane = 0
'                oItem.ToPane = 0

'                oButton = CType(oItem.Specific, SAPbouiCOM.Button)
'                oButton.Caption = "E Invoice Recon"

'                objForm.Freeze(False)

'            End If

'        Catch ex As Exception

'            Try
'                objForm.Freeze(False)
'            Catch
'            End Try

'            objMain.objApplication.StatusBar.SetText(
'            "Add Button Error : " & ex.Message,
'            SAPbouiCOM.BoMessageTime.bmt_Short,
'            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

'        End Try

'    End Sub

'#End Region

'#Region "Check Item Exists"

'    Private Function ItemExists(ByVal oForm As SAPbouiCOM.Form,
'                                ByVal ItemUID As String) As Boolean

'        Try
'            Dim x As SAPbouiCOM.Item
'            x = oForm.Items.Item(ItemUID)
'            Return True
'        Catch
'            Return False
'        End Try

'    End Function

'#End Region

'End Class

Imports System.Net
Imports System.Reflection.Emit
Imports System.Runtime.CompilerServices.RuntimeHelpers
Imports SAPbobsCOM

Public Class CLSEinvoiceButton

#Region "Declaration"
    Public objForm, objGRPOForm As SAPbouiCOM.Form
    Dim objMatrix, objGRPOMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head, oDBs_details As SAPbouiCOM.DBDataSource
    Dim BaseUom As String
    Dim CalcualtionBtnFlag As Boolean = False
    Dim strQry As String = ""
    Dim OrsSTRQRY, Ors, rs1, Orsc, Orsca, OrsPO, Orsme, Orsmm, Orsmd, Orsmg, Orsta, Orsds As SAPbobsCOM.Recordset
    Public Itemcode2 As New List(Of String)
    Dim LoadButtonFlag As Boolean = False
    Dim oItem As SAPbouiCOM.Item
    Dim oComboBox As SAPbouiCOM.ComboBox
    Dim oRecordset As SAPbobsCOM.Recordset
    Dim oButton As SAPbouiCOM.Button



#End Region



    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)

        Try
            If pVal.FormTypeEx = "133" Then   ' A/R Invoice

                Select Case pVal.EventType

                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                        If pVal.BeforeAction = False Then
                            objForm = objMain.objApplication.Forms.Item(FormUID)
                            AddEInvoiceButton(FormUID)
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED


                End Select

            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub
    'Private Sub AddEInvoiceButton(ByVal FormUID As String)

    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)

    '        Dim refItem As SAPbouiCOM.Item
    '        refItem = objForm.Items.Item("10000329")   'Copy From

    '        'Add new button
    '        oItem = objForm.Items.Add("btnBulkDel", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
    '        oItem.Top = objForm.Items.Item("42").Top + objForm.Items.Item("42").Height + 140
    '        oItem.Left = objForm.Items.Item("10000330").Left - 90
    '        oItem.Width = objForm.Items.Item("10000329").Width
    '        oItem.Height = objForm.Items.Item("10000329").Height
    '        oButton = oItem.Specific
    '        oButton.Caption = "Bulk Delivery"


    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try

    'End Sub


    Private Sub AddEInvoiceButton(ByVal FormUID As String)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim refItem As SAPbouiCOM.Item
            refItem = objForm.Items.Item("99")   'Copy From

            'Add new button
            oItem = objForm.Items.Add("btnEInv", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Top = objForm.Items.Item("105").Top + objForm.Items.Item("105").Height + 140
            oItem.Left = objForm.Items.Item("99").Left + 20
            oItem.Width = objForm.Items.Item("99").Width
            oItem.Height = objForm.Items.Item("99").Height
            oButton = oItem.Specific
            oButton.Caption = "E Invoice"
            'oButton.Caption = "E Invoice Reconciliation"

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub


End Class

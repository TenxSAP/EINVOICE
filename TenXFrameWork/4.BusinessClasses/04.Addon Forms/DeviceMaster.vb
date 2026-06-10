Imports Microsoft.VisualBasic
Imports SAPbouiCOM
Imports System.Text
Imports SBO_ADDONBASE
Imports SBO_ENUMS
Imports SBO_TABLES
Imports System.Threading
Imports System.IO
Imports System.Security.Cryptography
Public Class DeviceMaster
    Private Structure enControlName
        Const Master_Mxt As String = "Mtx"
        Const Master_Mxt_Col_DeviceMasterRowId As String = "#"
        Const Master_Mxt_Col_DeviceMasterCode As String = "Code"
        Const Master_Mxt_Col_DeviceMasterName As String = "Name"
        Public EnumNew As String
    End Structure
#Region "  Declaration  "
    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim objMatrix, objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim str, str1 As String
    Public rs, RsNum As SAPbobsCOM.Recordset
    Dim i, j As Int32
    Dim LostFocusFlag As Boolean = False
    Dim objutilities As Utilities
    'Private enSAPToolBarMenuUID As Object
#End Region
    Public Structure enSAPToolBarMenuUID
        Const Master_Mxt As String = "Mtx"
        Const cst_Find As String = "1281"
        Const cst_Add As String = "1282"
        Const cst_AddRow As String = "1292"
        Const cst_DeleteRow As String = "1293"
        Const cst_LastRecord As String = "1291"
        Const Cst_Cancel As String = "1284"
        Const cst_Close As String = "1286"
        'Const Master_Mxt_Col_DeviceMasterRowId As String = "#"
        'Const Master_Mxt_Col_DeviceMasterCode As String = "Code"
        'Const Master_Mxt_Col_DeviceMasterName As String = "Name"
        Public EnumNew As String
    End Structure

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("Device.xml", "DEVICE", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("DEVICE", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            objutilities = New Utilities
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_DEVICE")
            objMatrix = objForm.Items.Item("Mtx").Specific
            MasterMatrixLoad()
            'Me.SetNewLine(objForm.UniqueID)
            objMatrix.Columns.Item("Sts").Visible = False
            objMatrix.Columns.Item("DKey").Visible = False
            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("1293", True)
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub MasterMatrixLoad()
        Try
            objForm.EnableMenu(enSAPToolBarMenuUID.cst_Find, False)
            objForm.EnableMenu(enSAPToolBarMenuUID.cst_Add, False)
            objForm.EnableMenu(enSAPToolBarMenuUID.cst_AddRow, True)
            objForm.EnableMenu(enSAPToolBarMenuUID.cst_DeleteRow, True)
            oDBs_Head.Query()
            For i As Integer = 0 To oDBs_Head.Size - 1
                oDBs_Head.Offset = i
                j = 0
                j = j + 1
                Me.FormMatrix(enControlName.Master_Mxt).AddRow(j)
            Next i
            Me.FormMatrix(enControlName.Master_Mxt).FlushToDataSource()
            Me.FormMatrix(enControlName.Master_Mxt).LoadFromDataSource()
            Me.FormMatrix(enControlName.Master_Mxt).AutoResizeColumns()
            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
        Catch ex As Exception
            objForm.Freeze(False)
            MsgBox(ex.Message)
        End Try
    End Sub
    'Sub SetNewLine(ByVal FormUID As String)
    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_DEVICE")
    '        objMatrix = objForm.Items.Item("Mtx").Specific
    '        objMatrix.AddRow()
    '        oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.VisualRowCount)
    '        oDBs_Head.SetValue("Code", oDBs_Head.Offset, "")
    '        oDBs_Head.SetValue("Name", oDBs_Head.Offset, "")
    '        oDBs_Head.SetValue("U_MAC", oDBs_Head.Offset, "")
    '        oDBs_Head.SetValue("U_Remarks", oDBs_Head.Offset, "")
    '        objMatrix.SetLineData(objMatrix.VisualRowCount)
    '        objMatrix.AutoResizeColumns()
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub

    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("Mtx").Specific
            Dim Device As String = objMatrix.Columns.Item("Code").Cells.Item(1).Specific.Value
            Dim Devicename As String = objMatrix.Columns.Item("Name").Cells.Item(1).Specific.Value
            If Device = "" Then
                objMain.objApplication.StatusBar.SetText("Device Code is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If Devicename = "" Then
                objMain.objApplication.StatusBar.SetText("Device No is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Public ReadOnly Property FormCellEdit(ByVal pst_MatrixUID As String, ByVal pst_ColumnUID As String, ByVal pin_Row As Int32) As SAPbouiCOM.EditText
        Get
            Return CType(objForm.Items.Item(pst_MatrixUID).Specific.Columns.Item(pst_ColumnUID).Cells.Item(pin_Row).Specific, SAPbouiCOM.EditText)
        End Get
    End Property
    Public ReadOnly Property FormMatrix(ByVal pst_MatrixUID As String) As SAPbouiCOM.Matrix
        Get
            Return CType(objForm.Items.Item(pst_MatrixUID).Specific, SAPbouiCOM.Matrix)
        End Get
    End Property


    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            'If pVal.MenuUID = "DEVICE" And pVal.BeforeAction = False Then
            '    Me.CreateForm()
            'ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then


            'End If


            If pVal.MenuUID = "DEVICE" And pVal.BeforeAction = False Then
                Me.CreateForm()
                ' Me.SetNewLine()

                'ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                'ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                '    'Me.SetNewLine(objForm.UniqueID)
                'ElseIf pVal.MenuUID = "1293" And pVal.BeforeAction = False Then
                '    objMatrix = objForm.Items.Item("Mtx").Specific
                '    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                '        If objMatrix.IsRowSelected(i) = True Then
                '            objMatrix.DeleteRow(i)
                '        End If
                '    Next
                '    For i As Integer = 1 To objMatrix.VisualRowCount
                '        objMatrix.Columns.Item("#").Cells.Item(i).Specific.Value = i
                '    Next

            End If
            'comment
            Select Case pVal.BeforeAction
                Case True
                    Select Case pVal.MenuUID
                        Case enSAPToolBarMenuUID.cst_DeleteRow
                            If Me.FormCellEdit(enControlName.Master_Mxt, enControlName.Master_Mxt_Col_DeviceMasterCode, Me.FormMatrix(enControlName.Master_Mxt).VisualRowCount).Value = "" And Me.FormCellEdit(enControlName.Master_Mxt, enControlName.Master_Mxt_Col_DeviceMasterName, Me.FormMatrix(enControlName.Master_Mxt).VisualRowCount).Value = "" Then
                                oDBs_Head.RemoveRecord(Me.FormMatrix(enControlName.Master_Mxt).VisualRowCount - 1)
                                Me.FormMatrix(enControlName.Master_Mxt).DeleteRow(Me.FormMatrix(enControlName.Master_Mxt).VisualRowCount)
                                Me.FormMatrix(enControlName.Master_Mxt).FlushToDataSource()
                                BubbleEvent = False
                            Else
                                objForm.objApplication.StatusBar.SetText("Cannot Delete Records", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                            End If
                        Case enSAPToolBarMenuUID.cst_AddRow
                            Try
                                objForm.Freeze(True)
                                oDBs_Head.InsertRecord(oDBs_Head.Size)
                                oDBs_Head.Offset = oDBs_Head.Size - 1
                                Me.FormMatrix(enControlName.Master_Mxt).AddRow(1)
                                i = Me.FormMatrix(enControlName.Master_Mxt).VisualRowCount
                                Me.FormCellEdit(enControlName.Master_Mxt, enControlName.Master_Mxt_Col_DeviceMasterCode, Me.FormMatrix(enControlName.Master_Mxt).VisualRowCount).Active = True
                                Me.FormCellEdit(enControlName.Master_Mxt, enControlName.Master_Mxt_Col_DeviceMasterRowId, Me.FormMatrix(enControlName.Master_Mxt).VisualRowCount).Value = Me.FormMatrix(enControlName.Master_Mxt).VisualRowCount
                                objForm.Freeze(False)
                            Catch ex As Exception
                                objForm.Freeze(False)
                                objMain.objApplication.StatusBar.SetText("Unable to complte Add Operation" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try
                    End Select
                Case False
                    Select Case pVal.MenuUID
                        Case enSAPToolBarMenuUID.cst_Add

                        Case enSAPToolBarMenuUID.cst_Find
                        Case enSAPToolBarMenuUID.cst_DeleteRow
                        Case enSAPToolBarMenuUID.cst_AddRow
                        Case enSAPToolBarMenuUID.cst_LastRecord
                        Case enSAPToolBarMenuUID.Cst_Cancel
                        Case enSAPToolBarMenuUID.cst_Close
                    End Select
            End Select

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.Before_Action = True Then  ' SAP Business One Actions are done here 
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD
                        If Me.FormCellEdit(enControlName.Master_Mxt, enControlName.Master_Mxt_Col_DeviceMasterName, Me.FormMatrix(enControlName.Master_Mxt).VisualRowCount).Value = "" And Me.FormCellEdit(enControlName.Master_Mxt, enControlName.Master_Mxt_Col_DeviceMasterCode, Me.FormMatrix(enControlName.Master_Mxt).VisualRowCount).Value = "" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            oDBs_Head.RemoveRecord(oDBs_Head.Size - 1)
                            Me.FormMatrix(enControlName.Master_Mxt).DeleteRow(Me.FormMatrix(enControlName.Master_Mxt).VisualRowCount)
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                                            (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            If Validation(objForm.UniqueID) = False Then
                                BubbleEvent = False
                                Exit Sub
                            Else
                            End If
                        End If
                End Select

            Else
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD
                        If Me.FormCellEdit(enControlName.Master_Mxt, enControlName.Master_Mxt_Col_DeviceMasterCode, Me.FormMatrix(enControlName.Master_Mxt).VisualRowCount).Value <> "" Then
                            Dim rs, RsNum As SAPbobsCOM.Recordset

                            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Dim RowId As Int32
                            ' RowId = Me.FormMatrix(enControlName.Master_Mxt).VisualRowCount + 1
                            rs.DoQuery("UPDATE ONNM SET ""AutoKey"" = '" & RowId & "' WHERE ""ObjectCode"" = 'SBO_DEVICEUDO';")
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
                            oDBs_Head.SetValue("Code", Me.FormMatrix(enControlName.Master_Mxt).VisualRowCount - 1, Me.FormCellEdit(enControlName.Master_Mxt, enControlName.Master_Mxt_Col_DeviceMasterCode, Me.FormMatrix(enControlName.Master_Mxt).VisualRowCount).Value)
                            oDBs_Head.SetValue("Object", Me.FormMatrix(enControlName.Master_Mxt).VisualRowCount - 1, "SBO_DEVICEUDO")
                            oDBs_Head.SetValue("DataSource", Me.FormMatrix(enControlName.Master_Mxt).VisualRowCount - 1, "I")
                            'oDS.SetValue("DocEntry", Me.FormMatrix(enControlName.Master_Mxt).VisualRowCount - 1, Me.FormMatrix(enControlName.Master_Mxt).VisualRowCount - 1)
                            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            rs.DoQuery("SELECT ""USERID"" AS ""UserSign"" FROM ""OUSR"" WHERE ""U_NAME"" = '" & objMain.objCompany.UserName & "'")
                            oDBs_Head.SetValue("UserSign", Me.FormMatrix(enControlName.Master_Mxt).VisualRowCount, rs.Fields.Item("UserSign").Value)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
                        Else
                            Me.FormMatrix(enControlName.Master_Mxt).DeleteRow(Me.FormMatrix(enControlName.Master_Mxt).VisualRowCount)
                        End If
                End Select
            End If
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


#Region " Right Click Event"
    'Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)

    '    Dim objForm As SAPbouiCOM.Form
    '    Dim oMenuItem As SAPbouiCOM.MenuItem
    '    Dim oMenus As SAPbouiCOM.Menus

    '    Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
    '    oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
    '    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
    '    objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)


    '    'objMatrix = objForm.Items.Item("Mtx").Specific
    '    Try
    '        If eventInfo.FormUID = objForm.UniqueID Then
    '            If (eventInfo.BeforeAction = True) Then
    '                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
    '                    objForm.EnableMenu("1292", True)
    '                    objForm.EnableMenu("1293", True)


    '                End If
    '            End If
    '        End If
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub
#End Region



End Class


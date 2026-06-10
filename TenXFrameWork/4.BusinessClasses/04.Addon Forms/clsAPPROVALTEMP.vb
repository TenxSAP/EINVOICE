'Public Class clsAPPROVALTEMP


'#Region "       Declaration             "
'    Public objForm, objSTRForm As SAPbouiCOM.Form
'    Dim oDBs_Head, oDBs_Details1, oDBs_Details2, oDBs_Details3 As SAPbouiCOM.DBDataSource
'    Dim objMatrix1, objMatrix2 As SAPbouiCOM.Matrix
'    Dim objComboBox As SAPbouiCOM.ComboBox
'    Dim str, str1, str2, str3, Query, str4 As String
'    Public rs, RsNum As SAPbobsCOM.Recordset
'    Dim LostFocusFlag As Boolean = False
'    Dim objutilities As Utilities
'    Dim oDt As SAPbouiCOM.DataTable
'    Dim oGrid As SAPbouiCOM.Grid
'    Dim r1, r2, r3 As RadioButton
'    Dim column, column1 As SAPbouiCOM.EditTextColumn
'    Dim TERM As String
'#End Region
'    Sub CreateForm()
'        Try
'            objMain.objUtilities.LoadForm("ApprovalTemplates.xml", "Temp", ResourceType.Embeded)
'            objForm = objMain.objApplication.Forms.GetForm("Temp", objMain.objApplication.Forms.ActiveForm.TypeCount)
'            Disable_Checkboxes()
'            objForm.Freeze(True)

'            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_APPHDR")
'            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_APPREQ")
'            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_APPDOC")
'            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@SBO_APPAUT")
'            objMatrix1 = objForm.Items.Item("M1").Specific
'            objMatrix2 = objForm.Items.Item("M3").Specific
'            objForm.EnableMenu("1292", True)

'            objForm.EnableMenu("774", True)

'            '  .cst_Txt_Name

'            'objForm.DataBrowser.BrowseBy = "T1"
'            objForm.Items.Item("flReq").AffectsFormMode = False
'            objForm.Items.Item("flDoc").AffectsFormMode = False
'            objForm.Items.Item("flAut").AffectsFormMode = False
'            objForm.Items.Item("flCon").AffectsFormMode = False
'            objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
'            ' objMatrix2.FlushToDataSource()

'            'Disable_Checkboxes()
'            Me.SetNewLine1(objForm.UniqueID)
'            Me.SetNewLine2(objForm.UniqueID)
'            Me.fillDeptCombo(objForm.UniqueID)
'            objForm.Freeze(False)


'        Catch ex As Exception
'            objMain.objApplication.StatusBar.SetText(ex.Message)
'        End Try
'    End Sub
'    Public Sub fillDeptCombo(ByVal FormUID As String)
'        Try

'            objForm = objMain.objApplication.Forms.Item(objForm.UniqueID)
'            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_APPHDR")
'            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_APPREQ")
'            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_APPDOC")
'            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@SBO_APPAUT")
'            objMatrix1 = objForm.Items.Item("M1").Specific
'            Dim DOCK As SAPbouiCOM.Column = objMatrix1.Columns.Item("M1_2")
'            Dim QURY As String = "SELECT DISTINCT ""Code"", ""Name"" FROM OUDP"
'            objMain.objUtilities.MatrixComboBoxValues(DOCK, QURY)
'            objMatrix1.Columns.Item("M1_2").DisplayDesc = True
'            objMatrix1.Columns.Item("M1_2").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
'            'TEST


'        Catch ex As Exception
'            objForm.Freeze(False)
'            objMain.objApplication.StatusBar.SetText(ex.Message)
'        End Try
'    End Sub

'    Sub SetNewLine1(ByVal FormUID As String)
'        Try
'            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_APPHDR")
'            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_APPREQ")
'            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_APPDOC")
'            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@SBO_APPAUT")
'            objMatrix1 = objForm.Items.Item("M1").Specific
'            objMatrix2 = objForm.Items.Item("M3").Specific

'            objMatrix1.AddRow()
'            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
'            oDBs_Details1.SetValue("U_Name", oDBs_Details1.Offset, "")
'            oDBs_Details1.SetValue("U_EmpId", oDBs_Details1.Offset, "")
'            oDBs_Details1.SetValue("U_Dept", oDBs_Details1.Offset, "")


'            objMatrix1.SetLineData(objMatrix1.VisualRowCount)
'            objMatrix1.AutoResizeColumns()
'        Catch ex As Exception
'            objMain.objApplication.StatusBar.SetText(ex.Message)
'        End Try
'    End Sub
'    Sub SetNewLine2(ByVal FormUID As String)
'        Try
'            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_APPHDR")
'            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_APPREQ")
'            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_APPDOC")
'            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@SBO_APPAUT")
'            objMatrix1 = objForm.Items.Item("M1").Specific
'            objMatrix2 = objForm.Items.Item("M3").Specific

'            objMatrix2.AddRow()
'            oDBs_Details3.SetValue("LineID", oDBs_Details3.Offset, objMatrix2.VisualRowCount)
'            oDBs_Details3.SetValue("U_M3_1", oDBs_Details3.Offset, "")
'            oDBs_Details3.SetValue("U_M3_2", oDBs_Details3.Offset, "")



'            objMatrix2.SetLineData(objMatrix2.VisualRowCount)
'            objMatrix2.AutoResizeColumns()
'        Catch ex As Exception
'            objMain.objApplication.StatusBar.SetText(ex.Message)
'        End Try
'    End Sub
'    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
'        Try
'            If pVal.MenuUID = "Temp" And pVal.BeforeAction = False Then
'                Me.CreateForm()
'                'Me.SetDefault(objForm.UniqueID)
'            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
'                'Me.SetDefault(objForm.UniqueID)
'            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
'                If TERM = "M1" Then
'                    Me.SetNewLine1(objForm.UniqueID)
'                ElseIf TERM = "M3" Then
'                    Me.SetNewLine2(objForm.UniqueID)
'                End If
'            ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then
'                objForm = objMain.objApplication.Forms.GetForm("Temp", objMain.objApplication.Forms.ActiveForm.TypeCount)
'                objMatrix1 = objForm.Items.Item("M1").Specific
'                objMatrix2 = objForm.Items.Item("M3").Specific


'                If TERM = "M1" Then
'                    Dim row As Integer = objMatrix1.VisualRowCount
'                    If objMatrix1.IsRowSelected(1) <> True And objMatrix1.VisualRowCount < 1 Then
'                        objMatrix1.AddRow()
'                        oDBs_Details1.SetValue("V_-1", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
'                        objMatrix1.SetLineData(objMatrix1.VisualRowCount)
'                    End If
'                    If objMatrix1.IsRowSelected(row) = True Then
'                        objMatrix1.DeleteRow(row)
'                    Else
'                        For i As Integer = 1 To objMatrix1.VisualRowCount - 1
'                            If objMatrix1.IsRowSelected(i) = True Then
'                                '  If i <> 1 Then
'                                objMatrix1.DeleteRow(i)
'                                '   End If
'                            End If
'                        Next
'                    End If
'                    For i As Integer = 1 To objMatrix1.VisualRowCount
'                        objMatrix1.Columns.Item("V_-1").Cells.Item(i).Specific.Value = i
'                        'objMatrix.FlushToDataSource()
'                    Next
'                ElseIf TERM = "M3" Then
'                    Dim row As Integer = objMatrix2.VisualRowCount
'                    If objMatrix2.IsRowSelected(1) <> True And objMatrix2.VisualRowCount < 1 Then
'                        objMatrix2.AddRow()
'                        oDBs_Details3.SetValue("V_-1", oDBs_Details3.Offset, objMatrix2.VisualRowCount)
'                        objMatrix2.SetLineData(objMatrix2.VisualRowCount)
'                    End If
'                    If objMatrix2.IsRowSelected(row) = True Then
'                        objMatrix2.DeleteRow(row)
'                    Else
'                        For i As Integer = 1 To objMatrix2.VisualRowCount - 1
'                            If objMatrix2.IsRowSelected(i) = True Then
'                                '  If i <> 1 Then
'                                objMatrix2.DeleteRow(i)
'                                '   End If
'                            End If
'                        Next
'                    End If
'                    For i As Integer = 1 To objMatrix2.VisualRowCount - 1
'                        If objMatrix2.IsRowSelected(i) = True Then
'                            objMatrix2.DeleteRow(i)
'                        End If
'                    Next
'                End If

'                For i As Integer = 1 To objMatrix2.VisualRowCount
'                    objMatrix2.Columns.Item("V_-1").Cells.Item(i).Specific.Value = i
'                Next


'                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
'                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
'                End If

'            End If




'        Catch ex As Exception
'            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
'        End Try
'    End Sub
'    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
'        Try
'            Select Case pVal.EventType
'                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

'                    If pVal.ItemUID = "Rect" And pVal.BeforeAction = True Then
'                        Disable_Checkboxes()
'                        'objForm.Items.Item("Item_11").Enabled = True
'                        objForm.Items.Item("Item_13").Enabled = True
'                        objForm.Items.Item("Item_15").Enabled = True
'                        objForm.Items.Item("Item_38").Enabled = True

'                    End If
'                    If pVal.ItemUID = "HR" And pVal.BeforeAction = True Then
'                        Disable_Checkboxes()
'                        objForm.Items.Item("Item_12").Enabled = True
'                        objForm.Items.Item("Item_18").Enabled = True
'                        objForm.Items.Item("Item_14").Enabled = True

'                    End If
'                    If pVal.ItemUID = "Payroll" And pVal.BeforeAction = True Then
'                        Disable_Checkboxes()
'                        objForm.Items.Item("Item_22").Enabled = True
'                        objForm.Items.Item("Item_23").Enabled = True
'                        objForm.Items.Item("Item_26").Enabled = True
'                        objForm.Items.Item("Item_27").Enabled = True
'                        objForm.Items.Item("Item_29").Enabled = True
'                        objForm.Items.Item("Item_30").Enabled = True
'                        objForm.Items.Item("PerEval").Enabled = True
'                        objForm.Items.Item("Encash").Enabled = True

'                    End If
'                    If pVal.ItemUID = "TmAt" And pVal.BeforeAction = True Then
'                        Disable_Checkboxes()
'                        objForm.Items.Item("Item_7").Enabled = True
'                        objForm.Items.Item("Item_21").Enabled = True
'                        '  objForm.Items.Item("Item_19").Enabled = True
'                        objForm.Items.Item("Item_24").Enabled = True
'                        objForm.Items.Item("Item_20").Enabled = True
'                        objForm.Items.Item("Item_3").Enabled = True
'                        objForm.Items.Item("Item_1").Enabled = True
'                        objForm.Items.Item("Item_28").Enabled = True
'                        objForm.Items.Item("Item_35").Enabled = True
'                        objForm.Items.Item("Item_31").Enabled = True
'                        objForm.Items.Item("Item_25").Enabled = True
'                        objForm.Items.Item("LvApp").Enabled = True


'                    End If
'                    If pVal.ItemUID = "EOSDoc" And pVal.BeforeAction = True Then
'                        Disable_Checkboxes()
'                        objForm.Items.Item("Item_5").Enabled = True
'                        objForm.Items.Item("Item_16").Enabled = True
'                        objForm.Items.Item("Item_34").Enabled = True

'                    End If
'                    If pVal.ItemUID = "DocReq" And pVal.BeforeAction = True Then
'                        Disable_Checkboxes()
'                        objForm.Items.Item("Item_8").Enabled = True
'                        objForm.Items.Item("Item_9").Enabled = True
'                        objForm.Items.Item("Item_17").Enabled = True

'                    End If

'                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
'                    Dim oCFL As SAPbouiCOM.ChooseFromList
'                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
'                    Dim CFL_Id As String
'                    CFL_Id = CFLEvent.ChooseFromListUID
'                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)

'                    Dim oDT As SAPbouiCOM.DataTable
'                    oDT = CFLEvent.SelectedObjects
'                    If oCFL.UniqueID = "CFL_1" And pVal.BeforeAction = True Then
'                    End If
'                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
'                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
'                        If oCFL.UniqueID = "ouser2" Then
'                            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
'                            oDBs_Details1.SetValue("U_Name", oDBs_Details1.Offset, oDT.GetValue("USER_CODE", 0))
'                            oDBs_Details1.SetValue("U_Dept", oDBs_Details1.Offset, oDT.GetValue("Department", 0))
'                            '  oDBs_Details1.SetValue("U_UKey", oDBs_Details1.Offset, oDT.GetValue("USER_CODE", 0))
'                            objMatrix1.SetLineData(pVal.Row)
'                        End If
'                        If oCFL.UniqueID = "CFL_EMP" Then
'                            ' oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
'                            oDBs_Details1.SetValue("U_EmpId", oDBs_Details1.Offset, oDT.GetValue("Code", 0))
'                            oDBs_Details1.SetValue("U_Dept", oDBs_Details1.Offset, oDT.GetValue("dept", 0))

'                        End If
'                        If oCFL.UniqueID = "Stages" Then
'                            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@SBO_APPAUT")
'                            Dim c As String = oDT.GetValue("Code", 0)
'                            Dim d As String = oDT.GetValue("U_SDesc", 0)
'                            oDBs_Details3.SetValue("LineId", oDBs_Details3.Offset, pVal.Row)
'                            oDBs_Details3.SetValue("U_M3_1", oDBs_Details3.Offset, c)
'                            oDBs_Details3.SetValue("U_M3_2", oDBs_Details3.Offset, d)
'                            objMatrix2.SetLineData(pVal.Row)
'                        End If
'                    End If
'                'Else


'                Case SAPbouiCOM.BoEventTypes.et_CLICK
'                    Select Case pVal.ItemUID
'                        Case "flReq"
'                            objForm.Freeze(True)
'                            objForm.PaneLevel = 1
'                            '         FloderName = "Requester"
'                            objForm.Freeze(False)
'                        Case "flDoc"
'                            objForm.Freeze(True)
'                            objForm.PaneLevel = 2
'                            'FloderName = "Documents"
'                            objForm.Freeze(False)
'                        Case "flAut"
'                            objForm.Freeze(True)
'                            objForm.PaneLevel = 3
'                            ' FloderName = "Requester"
'                            objForm.Freeze(False)
'                        Case "flCon"
'                            objForm.Freeze(True)
'                            objForm.PaneLevel = 4
'                            ' FloderName = "Conditions"
'                            objForm.Freeze(False)
'                        Case "M1"
'                            TERM = pVal.ItemUID
'                        Case "M3"
'                            TERM = pVal.ItemUID

'                    End Select
'                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
'                    If pVal.ItemUID = "T1" And pVal.BeforeAction = True Then
'                        If objForm.Items.Item("T1").Specific.Value.ToString().Length() > 10 Then
'                            objMain.objApplication.StatusBar.SetText("Name should be less then 10 characters.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
'                            BubbleEvent = False
'                        End If
'                    End If


'                Case SAPbouiCOM.BoEventTypes.et_CLICK
'                    If pVal.ItemUID = "M1" Or pVal.ItemUID = "M3" Then
'                        TERM = pVal.ItemUID

'                    End If
'                    'Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

'                    '    ' objMatrix1 = objForm.Items.Item("M1").Specific
'                    '    'objMatrix2 = objForm.Items.Item("M3").Specific

'                    '    If pVal.ItemUID = "M1" And pVal.ColUID = "M1_1" Then
'                    '        If objMatrix1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() <> "" Then
'                    '            If pVal.Row = objMatrix1.VisualRowCount Then
'                    '                Me.SetNewLine1(objForm.UniqueID)
'                    '            End If
'                    '        End If
'                    '    End If

'                    '    If pVal.ItemUID = "M3" And pVal.ColUID = "M3_1" Then
'                    '        If objMatrix2.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() <> "" Then
'                    '            If pVal.Row = objMatrix1.VisualRowCount Then
'                    '                Me.SetNewLine2(objForm.UniqueID)
'                    '            End If
'                    '        End If
'                    '    End If

'            End Select
'        Catch ex As Exception
'            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

'            objForm.Freeze(False)


'        End Try


'    End Sub


'    Function Validation(ByVal FormUID As String)

'        objForm = objMain.objApplication.Forms.Item(FormUID)
'        '  oDBs_Head = objForm.DataSources.DBDataSources.Item("@FISH_13")
'        ' oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@FISH_41")
'        ' objMatrix1 = objForm.Items.Item("Mtx").Specific

'        'If oDBs_Head.GetValue("U_NAME", oDBs_Head.Offset) = "" Then
'        '    objMain.objApplication.StatusBar.SetText("Name is missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
'        '    Return False
'        'End If

'    End Function

'    Public Function Disable_Checkboxes()
'        Try
'            'objForm.Items.Item("Item_11").Enabled = False
'            'objForm.Items.Item("Item_13").Enabled = False
'            'objForm.Items.Item("Item_15").Enabled = False
'            'objForm.Items.Item("Item_38").Enabled = False

'            'objForm.Items.Item("Item_12").Enabled = False
'            'objForm.Items.Item("Item_18").Enabled = False
'            'objForm.Items.Item("Item_14").Enabled = False

'            'objForm.Items.Item("Item_7").Enabled = False
'            'objForm.Items.Item("Item_21").Enabled = False
'            '' objForm.Items.Item("Item_19").Enabled = False
'            'objForm.Items.Item("Item_24").Enabled = False
'            'objForm.Items.Item("Item_20").Enabled = False
'            'objForm.Items.Item("Item_3").Enabled = False
'            'objForm.Items.Item("Item_1").Enabled = False
'            'objForm.Items.Item("Item_28").Enabled = False
'            '' objForm.Items.Item("Item_38").Enabled = False


'            'objForm.Items.Item("Item_22").Enabled = False
'            'objForm.Items.Item("Item_23").Enabled = False
'            'objForm.Items.Item("Item_26").Enabled = False
'            'objForm.Items.Item("Item_27").Enabled = False
'            'objForm.Items.Item("Item_29").Enabled = False
'            'objForm.Items.Item("Item_30").Enabled = False
'            'objForm.Items.Item("PerEval").Enabled = False
'            'objForm.Items.Item("Encash").Enabled = False
'            'objForm.Items.Item("Item_25").Enabled = False
'            'objForm.Items.Item("LvApp").Enabled = False



'            'objForm.Items.Item("Item_5").Enabled = False
'            'objForm.Items.Item("Item_8").Enabled = False
'            'objForm.Items.Item("Item_9").Enabled = False
'            'objForm.Items.Item("Item_16").Enabled = False
'            'objForm.Items.Item("Item_34").Enabled = False
'            'objForm.Items.Item("Item_17").Enabled = False
'            'objForm.Items.Item("Item_35").Enabled = False
'            'objForm.Items.Item("Item_31").Enabled = False

'        Catch ex As Exception
'            objMain.objApplication.StatusBar.SetText(ex.Message)
'        End Try
'        'End If
'    End Function


'#Region " Right Click Event"
'    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)

'        Dim objForm As SAPbouiCOM.Form
'        Dim oMenuItem As SAPbouiCOM.MenuItem
'        Dim oMenus As SAPbouiCOM.Menus
'        objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)
'        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
'        oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
'        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
'        oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_APPHDR")
'        oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_APPREQ")
'        oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_APPDOC")
'        oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@SBO_APPAUT")
'        objMatrix1 = objForm.Items.Item("M1").Specific
'        objMatrix2 = objForm.Items.Item("M3").Specific

'        Try
'            If eventInfo.FormUID = objForm.UniqueID Then
'                If (eventInfo.BeforeAction = True) Then
'                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then 'SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
'                        '  If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Or objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then

'                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then 'SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
'                            '  objMatrix = objForm.Items.Item("Matx").Specific
'                            objForm.EnableMenu("1292", True)
'                            objForm.EnableMenu("774", True)

'                            If eventInfo.ItemUID = "M1" And eventInfo.ColUID = "V_-1" Then
'                                Try

'                                Catch ex As Exception
'                                    objMain.objApplication.StatusBar.SetText(ex.Message)
'                                End Try
'                            ElseIf eventInfo.ItemUID = "M3" And eventInfo.ColUID = "V_-1" Then
'                                Try

'                                Catch ex As Exception
'                                    objForm.Freeze(False)

'                                    objMain.objApplication.StatusBar.SetText(ex.Message)
'                                End Try

'                            End If

'                        End If
'                    End If
'                End If
'            End If
'        Catch ex As Exception
'            objMain.objApplication.StatusBar.SetText(ex.Message)
'        End Try
'    End Sub
'#End Region

'End Class

'madhu commetted on 30/03/2026
Public Class clsAPPROVALTEMP


#Region "       Declaration             "
    Public objForm, objSTRForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1, oDBs_Details2, oDBs_Details3 As SAPbouiCOM.DBDataSource
    Dim objMatrix1, objMatrix2 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim str, str1, str2, str3, Query, str4 As String
    Public rs, RsNum As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False
    Dim objutilities As Utilities
    Dim oDt As SAPbouiCOM.DataTable
    Dim oGrid As SAPbouiCOM.Grid
    Dim r1, r2, r3 As RadioButton
    Dim objMIForm As SAPbouiCOM.Form

    Dim column, column1 As SAPbouiCOM.EditTextColumn
    Dim TERM As String
#End Region
    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("ApprovalTemplates.xml", "Temp", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("Temp", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Disable_Checkboxes()
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_APPHDR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_APPREQ")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_APPDOC")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@SBO_APPAUT")
            objMatrix1 = objForm.Items.Item("M1").Specific
            objMatrix2 = objForm.Items.Item("M3").Specific
            objForm.EnableMenu("1292", True)

            objForm.EnableMenu("774", True)

            '  .cst_Txt_Name

            'objForm.DataBrowser.BrowseBy = "T1"
            objForm.Items.Item("flReq").AffectsFormMode = False
            objForm.Items.Item("flDoc").AffectsFormMode = False
            objForm.Items.Item("flAut").AffectsFormMode = False
            objForm.Items.Item("flCon").AffectsFormMode = False
            objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            ' objMatrix2.FlushToDataSource()

            'Disable_Checkboxes()
            Me.SetNewLine1(objForm.UniqueID)
            Me.SetNewLine2(objForm.UniqueID)
            Me.fillDeptCombo(objForm.UniqueID)
            objForm.Freeze(False)


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Sub fillDeptCombo(ByVal FormUID As String)
        Try

            objForm = objMain.objApplication.Forms.Item(objForm.UniqueID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_APPHDR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_APPREQ")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_APPDOC")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@SBO_APPAUT")
            objMatrix1 = objForm.Items.Item("M1").Specific
            Dim DOCK As SAPbouiCOM.Column = objMatrix1.Columns.Item("M1_2")
            Dim QURY As String = "SELECT DISTINCT ""Code"", ""Name"" FROM OUDP"
            objMain.objUtilities.MatrixComboBoxValues(DOCK, QURY)
            objMatrix1.Columns.Item("M1_2").DisplayDesc = True
            objMatrix1.Columns.Item("M1_2").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            'TEST


        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine1(ByVal FormUID As String)
        Try
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_APPHDR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_APPREQ")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_APPDOC")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@SBO_APPAUT")
            objMatrix1 = objForm.Items.Item("M1").Specific
            objMatrix2 = objForm.Items.Item("M3").Specific

            objMatrix1.AddRow()
            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
            oDBs_Details1.SetValue("U_Name", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_EmpId", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_Dept", oDBs_Details1.Offset, "")


            objMatrix1.SetLineData(objMatrix1.VisualRowCount)
            objMatrix1.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub SetNewLine2(ByVal FormUID As String)
        Try
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_APPHDR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_APPREQ")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_APPDOC")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@SBO_APPAUT")
            objMatrix1 = objForm.Items.Item("M1").Specific
            objMatrix2 = objForm.Items.Item("M3").Specific

            objMatrix2.AddRow()
            oDBs_Details3.SetValue("LineID", oDBs_Details3.Offset, objMatrix2.VisualRowCount)
            oDBs_Details3.SetValue("U_M3_1", oDBs_Details3.Offset, "")
            oDBs_Details3.SetValue("U_M3_2", oDBs_Details3.Offset, "")



            objMatrix2.SetLineData(objMatrix2.VisualRowCount)
            objMatrix2.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "Temp" And pVal.BeforeAction = False Then
                Me.CreateForm()
                'Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                'Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                If TERM = "M1" Then
                    Me.SetNewLine1(objForm.UniqueID)
                ElseIf TERM = "M3" Then
                    Me.SetNewLine2(objForm.UniqueID)
                End If
            ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("Temp", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix1 = objForm.Items.Item("M1").Specific
                objMatrix2 = objForm.Items.Item("M3").Specific


                If TERM = "M1" Then
                    Dim row As Integer = objMatrix1.VisualRowCount
                    If objMatrix1.IsRowSelected(1) <> True And objMatrix1.VisualRowCount < 1 Then
                        objMatrix1.AddRow()
                        oDBs_Details1.SetValue("V_-1", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
                        objMatrix1.SetLineData(objMatrix1.VisualRowCount)
                    End If
                    If objMatrix1.IsRowSelected(row) = True Then
                        objMatrix1.DeleteRow(row)
                    Else
                        For i As Integer = 1 To objMatrix1.VisualRowCount - 1
                            If objMatrix1.IsRowSelected(i) = True Then
                                '  If i <> 1 Then
                                objMatrix1.DeleteRow(i)
                                '   End If
                            End If
                        Next
                    End If
                    For i As Integer = 1 To objMatrix1.VisualRowCount
                        objMatrix1.Columns.Item("V_-1").Cells.Item(i).Specific.Value = i
                        'objMatrix.FlushToDataSource()
                    Next
                ElseIf TERM = "M3" Then
                    Dim row As Integer = objMatrix2.VisualRowCount
                    If objMatrix2.IsRowSelected(1) <> True And objMatrix2.VisualRowCount < 1 Then
                        objMatrix2.AddRow()
                        oDBs_Details3.SetValue("V_-1", oDBs_Details3.Offset, objMatrix2.VisualRowCount)
                        objMatrix2.SetLineData(objMatrix2.VisualRowCount)
                    End If
                    If objMatrix2.IsRowSelected(row) = True Then
                        objMatrix2.DeleteRow(row)
                    Else
                        For i As Integer = 1 To objMatrix2.VisualRowCount - 1
                            If objMatrix2.IsRowSelected(i) = True Then
                                '  If i <> 1 Then
                                objMatrix2.DeleteRow(i)
                                '   End If
                            End If
                        Next
                    End If
                    For i As Integer = 1 To objMatrix2.VisualRowCount - 1
                        If objMatrix2.IsRowSelected(i) = True Then
                            objMatrix2.DeleteRow(i)
                        End If
                    Next
                End If

                For i As Integer = 1 To objMatrix2.VisualRowCount
                    objMatrix2.Columns.Item("V_-1").Cells.Item(i).Specific.Value = i
                Next


                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If

            End If




        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.ItemUID = "Rect" And pVal.BeforeAction = True Then
                        Disable_Checkboxes()
                        'objForm.Items.Item("Item_11").Enabled = True
                        objForm.Items.Item("Item_13").Enabled = True
                        objForm.Items.Item("Item_15").Enabled = True
                        objForm.Items.Item("Item_38").Enabled = True

                    End If
                    If pVal.ItemUID = "HR" And pVal.BeforeAction = True Then
                        Disable_Checkboxes()
                        objForm.Items.Item("Item_12").Enabled = True
                        objForm.Items.Item("Item_18").Enabled = True
                        objForm.Items.Item("Item_14").Enabled = True

                    End If
                    If pVal.ItemUID = "Payroll" And pVal.BeforeAction = True Then
                        Disable_Checkboxes()
                        objForm.Items.Item("Item_22").Enabled = True
                        objForm.Items.Item("Item_23").Enabled = True
                        objForm.Items.Item("Item_26").Enabled = True
                        objForm.Items.Item("Item_27").Enabled = True
                        objForm.Items.Item("Item_29").Enabled = True
                        objForm.Items.Item("Item_30").Enabled = True
                        objForm.Items.Item("PerEval").Enabled = True
                        objForm.Items.Item("Encash").Enabled = True

                    End If
                    If pVal.ItemUID = "TmAt" And pVal.BeforeAction = True Then
                        Disable_Checkboxes()
                        objForm.Items.Item("Item_7").Enabled = True
                        objForm.Items.Item("Item_21").Enabled = True
                        '  objForm.Items.Item("Item_19").Enabled = True
                        objForm.Items.Item("Item_24").Enabled = True
                        objForm.Items.Item("Item_20").Enabled = True
                        objForm.Items.Item("Item_3").Enabled = True
                        objForm.Items.Item("Item_1").Enabled = True
                        objForm.Items.Item("Item_28").Enabled = True
                        objForm.Items.Item("Item_35").Enabled = True
                        objForm.Items.Item("Item_31").Enabled = True
                        objForm.Items.Item("Item_25").Enabled = True
                        objForm.Items.Item("LvApp").Enabled = True


                    End If
                    If pVal.ItemUID = "EOSDoc" And pVal.BeforeAction = True Then
                        Disable_Checkboxes()
                        objForm.Items.Item("Item_5").Enabled = True
                        objForm.Items.Item("Item_16").Enabled = True
                        objForm.Items.Item("Item_34").Enabled = True

                    End If
                    If pVal.ItemUID = "DocReq" And pVal.BeforeAction = True Then
                        Disable_Checkboxes()
                        objForm.Items.Item("Item_8").Enabled = True
                        objForm.Items.Item("Item_9").Enabled = True
                        objForm.Items.Item("Item_17").Enabled = True

                    End If
                    'Akhila
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        If Me.Validation(objForm.UniqueID) = False Then BubbleEvent = False
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)

                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    If oCFL.UniqueID = "CFL_1" And pVal.BeforeAction = True Then
                    End If
                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If oCFL.UniqueID = "ouser2" Then
                            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
                            oDBs_Details1.SetValue("U_Name", oDBs_Details1.Offset, oDT.GetValue("USER_CODE", 0))
                            oDBs_Details1.SetValue("U_Dept", oDBs_Details1.Offset, oDT.GetValue("Department", 0))
                            '  oDBs_Details1.SetValue("U_UKey", oDBs_Details1.Offset, oDT.GetValue("USER_CODE", 0))
                            objMatrix1.SetLineData(pVal.Row)
                        End If
                        If oCFL.UniqueID = "CFL_EMP" Then
                            ' oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
                            oDBs_Details1.SetValue("U_EmpId", oDBs_Details1.Offset, oDT.GetValue("empID", 0))
                            oDBs_Details1.SetValue("U_Dept", oDBs_Details1.Offset, oDT.GetValue("dept", 0))
                            objMatrix1.SetLineData(pVal.Row)

                        End If
                        If oCFL.UniqueID = "Stages" Then
                            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@SBO_APPAUT")
                            Dim c As String = oDT.GetValue("Code", 0)
                            Dim d As String = oDT.GetValue("U_SDesc", 0)
                            oDBs_Details3.SetValue("LineId", oDBs_Details3.Offset, pVal.Row)
                            oDBs_Details3.SetValue("U_M3_1", oDBs_Details3.Offset, c)
                            oDBs_Details3.SetValue("U_M3_2", oDBs_Details3.Offset, d)
                            objMatrix2.SetLineData(pVal.Row)
                        End If
                    End If
                'Else


                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    Select Case pVal.ItemUID
                        Case "flReq"
                            objForm.Freeze(True)
                            objForm.PaneLevel = 1
                            '         FloderName = "Requester"
                            objForm.Freeze(False)
                        Case "flDoc"
                            objForm.Freeze(True)
                            objForm.PaneLevel = 2
                            'FloderName = "Documents"
                            objForm.Freeze(False)
                        Case "flAut"
                            objForm.Freeze(True)
                            objForm.PaneLevel = 3
                            ' FloderName = "Requester"
                            objForm.Freeze(False)
                        Case "flCon"
                            objForm.Freeze(True)
                            objForm.PaneLevel = 4
                            ' FloderName = "Conditions"
                            objForm.Freeze(False)
                        Case "M1"
                            TERM = pVal.ItemUID
                        Case "M3"
                            TERM = pVal.ItemUID

                    End Select
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    If pVal.ItemUID = "T1" And pVal.BeforeAction = True Then
                        If objForm.Items.Item("T1").Specific.Value.ToString().Length() > 10 Then
                            objMain.objApplication.StatusBar.SetText("Name should be less then 10 characters.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                        End If
                    End If


                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "M1" Or pVal.ItemUID = "M3" Then
                        TERM = pVal.ItemUID

                    End If

                    'Akhila
                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix1 = objForm.Items.Item("M1").Specific

                    If pVal.ItemUID = "M1" And pVal.ColUID = "EmpId" And pVal.Before_Action = False Then
                        objMatrix1 = objForm.Items.Item("M1").Specific

                        Dim DocumentNumber As String = objMatrix1.Columns.Item("EmpId").Cells.Item(pVal.Row).Specific.value
                        objMain.objApplication.ActivateMenuItem("3590")
                        objMIForm = objMain.objApplication.Forms.GetForm("60100", objMain.objApplication.Forms.ActiveForm.TypeCount)
                        objMIForm.Freeze(True)
                        objMIForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objMIForm.Items.Item("480002078").Specific.Value = DocumentNumber
                        objMIForm.Items.Item("2").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objMIForm.Freeze(False)

                    End If
                    If pVal.ItemUID = "M3" And pVal.ColUID = "M3_1" And pVal.Before_Action = False Then
                        objMatrix2 = objForm.Items.Item("M3").Specific

                        Dim Stagename As String = objMatrix2.Columns.Item("M3_1").Cells.Item(pVal.Row).Specific.value
                        objMain.objApplication.ActivateMenuItem("SBO_AST")
                        objMIForm = objMain.objApplication.Forms.GetForm("SBO_AST", objMain.objApplication.Forms.ActiveForm.TypeCount)
                        objMIForm.Freeze(True)
                        objMIForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objMIForm.Items.Item("SName").Specific.Value = Stagename
                        objMIForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objMIForm.Freeze(False)

                    End If





                    'Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                    '    ' objMatrix1 = objForm.Items.Item("M1").Specific
                    '    'objMatrix2 = objForm.Items.Item("M3").Specific

                    '    If pVal.ItemUID = "M1" And pVal.ColUID = "M1_1" Then
                    '        If objMatrix1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() <> "" Then
                    '            If pVal.Row = objMatrix1.VisualRowCount Then
                    '                Me.SetNewLine1(objForm.UniqueID)
                    '            End If
                    '        End If
                    '    End If

                    '    If pVal.ItemUID = "M3" And pVal.ColUID = "M3_1" Then
                    '        If objMatrix2.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() <> "" Then
                    '            If pVal.Row = objMatrix1.VisualRowCount Then
                    '                Me.SetNewLine2(objForm.UniqueID)
                    '            End If
                    '        End If
                    '    End If

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            objForm.Freeze(False)


        End Try


    End Sub


    Function Validation(ByVal FormUID As String)

        objForm = objMain.objApplication.Forms.Item(FormUID)
        oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_APPREQ")
        oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@SBO_APPAUT")

        objMatrix1 = objForm.Items.Item("M1").Specific
        objMatrix2 = objForm.Items.Item("M3").Specific

        If oDBs_Details1.GetValue("U_Name", oDBs_Details1.Offset) = "" Then
            objMain.objApplication.StatusBar.SetText(
                "On ""Originator"" tab, fill in ""Requester User ID"" column",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        If oDBs_Details3.GetValue("U_M3_1", oDBs_Details3.Offset) = "" Then
            objMain.objApplication.StatusBar.SetText(
                "On ""Stages"" tab, fill in ""Stage Name"" column.",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        Return True



        '  oDBs_Head = objForm.DataSources.DBDataSources.Item("@FISH_13")
        ' oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@FISH_41")
        ' objMatrix1 = objForm.Items.Item("Mtx").Specific

        'If oDBs_Head.GetValue("U_NAME", oDBs_Head.Offset) = "" Then
        '    objMain.objApplication.StatusBar.SetText("Name is missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '    Return False
        'End If

    End Function

    Public Function Disable_Checkboxes()
        Try
            'objForm.Items.Item("Item_11").Enabled = False
            'objForm.Items.Item("Item_13").Enabled = False
            'objForm.Items.Item("Item_15").Enabled = False
            'objForm.Items.Item("Item_38").Enabled = False

            'objForm.Items.Item("Item_12").Enabled = False
            'objForm.Items.Item("Item_18").Enabled = False
            'objForm.Items.Item("Item_14").Enabled = False

            'objForm.Items.Item("Item_7").Enabled = False
            'objForm.Items.Item("Item_21").Enabled = False
            '' objForm.Items.Item("Item_19").Enabled = False
            'objForm.Items.Item("Item_24").Enabled = False
            'objForm.Items.Item("Item_20").Enabled = False
            'objForm.Items.Item("Item_3").Enabled = False
            'objForm.Items.Item("Item_1").Enabled = False
            'objForm.Items.Item("Item_28").Enabled = False
            '' objForm.Items.Item("Item_38").Enabled = False


            'objForm.Items.Item("Item_22").Enabled = False
            'objForm.Items.Item("Item_23").Enabled = False
            'objForm.Items.Item("Item_26").Enabled = False
            'objForm.Items.Item("Item_27").Enabled = False
            'objForm.Items.Item("Item_29").Enabled = False
            'objForm.Items.Item("Item_30").Enabled = False
            'objForm.Items.Item("PerEval").Enabled = False
            'objForm.Items.Item("Encash").Enabled = False
            'objForm.Items.Item("Item_25").Enabled = False
            'objForm.Items.Item("LvApp").Enabled = False



            'objForm.Items.Item("Item_5").Enabled = False
            'objForm.Items.Item("Item_8").Enabled = False
            'objForm.Items.Item("Item_9").Enabled = False
            'objForm.Items.Item("Item_16").Enabled = False
            'objForm.Items.Item("Item_34").Enabled = False
            'objForm.Items.Item("Item_17").Enabled = False
            'objForm.Items.Item("Item_35").Enabled = False
            'objForm.Items.Item("Item_31").Enabled = False

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
        'End If
    End Function


#Region " Right Click Event"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)

        Dim objForm As SAPbouiCOM.Form
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_APPHDR")
        oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_APPREQ")
        oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_APPDOC")
        oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@SBO_APPAUT")
        objMatrix1 = objForm.Items.Item("M1").Specific
        objMatrix2 = objForm.Items.Item("M3").Specific

        Try
            If eventInfo.FormUID = objForm.UniqueID Then
                If (eventInfo.BeforeAction = True) Then
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then 'SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        '  If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Or objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then

                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then 'SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                            '  objMatrix = objForm.Items.Item("Matx").Specific
                            objForm.EnableMenu("1292", True)
                            objForm.EnableMenu("774", True)

                            If eventInfo.ItemUID = "M1" And eventInfo.ColUID = "V_-1" Then
                                Try

                                Catch ex As Exception
                                    objMain.objApplication.StatusBar.SetText(ex.Message)
                                End Try
                            ElseIf eventInfo.ItemUID = "M3" And eventInfo.ColUID = "V_-1" Then
                                Try

                                Catch ex As Exception
                                    objForm.Freeze(False)

                                    objMain.objApplication.StatusBar.SetText(ex.Message)
                                End Try

                            End If

                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

End Class


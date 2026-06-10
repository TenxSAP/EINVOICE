'Public Class ClsAPPROVALSTAGES


'#Region "       Declaration             "
'    Public objForm, objSTRForm As SAPbouiCOM.Form
'    Dim oDBs_Head, oDBs_Details1 As SAPbouiCOM.DBDataSource
'    Dim objMatrix1, MATRIXS As SAPbouiCOM.Matrix
'    Dim objComboBox As SAPbouiCOM.ComboBox
'    Dim str, str1, str2, str3, Query, str4 As String
'    Public rs, RsNum As SAPbobsCOM.Recordset
'    Dim LostFocusFlag As Boolean = False
'    Dim objutilities As Utilities
'    Dim oDt As SAPbouiCOM.DataTable
'    Dim oGrid As SAPbouiCOM.Grid
'    Dim r1, r2, r3 As RadioButton
'    Dim column, column1 As SAPbouiCOM.EditTextColumn

'#End Region

'    Sub CreateForm()
'        Try
'            objMain.objUtilities.LoadForm("Approvalstages.xml", "SBO_AST", ResourceType.Embeded)
'            objForm = objMain.objApplication.Forms.GetForm("SBO_AST", objMain.objApplication.Forms.ActiveForm.TypeCount)
'            objForm.Freeze(True)

'            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_AST")
'            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_AST_C0")
'            objMatrix1 = objForm.Items.Item("Mtx").Specific
'            objForm.EnableMenu("774", True)
'            objForm.EnableMenu("1292", True)
'            ' objForm.EnableMenu("1293", True)

'            objMatrix1.AutoResizeColumns()
'            objMatrix1.Clear()
'            objForm.DataBrowser.BrowseBy = "SName"


'            objMatrix1.FlushToDataSource()

'            Me.SetNewLine1(objForm.UniqueID)
'            Me.fillDeptCombo(objForm.UniqueID)
'            objForm.Freeze(False)
'        Catch ex As Exception
'            objForm.Freeze(False)
'        End Try

'    End Sub
'    Sub SetNewLine1(ByVal FormUID As String)
'        Try
'            objForm = objMain.objApplication.Forms.Item(FormUID)
'            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_AST")
'            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_AST_C0")
'            objMatrix1 = objForm.Items.Item("Mtx").Specific

'            objMatrix1.AddRow()
'            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
'            oDBs_Details1.SetValue("U_AUTH", oDBs_Details1.Offset, "")
'            oDBs_Details1.SetValue("U_EmpId", oDBs_Details1.Offset, "")
'            oDBs_Details1.SetValue("U_Dept", oDBs_Details1.Offset, "")
'            oDBs_Details1.SetValue("U_UKey", oDBs_Details1.Offset, "")

'            objMatrix1.SetLineData(objMatrix1.VisualRowCount)
'            objMatrix1.AutoResizeColumns()
'        Catch ex As Exception
'            objForm.Freeze(False)

'            objMain.objApplication.StatusBar.SetText(ex.Message)
'        End Try
'    End Sub
'    Public Sub fillDeptCombo(ByVal FormUID As String)
'        Try

'            objForm = objMain.objApplication.Forms.Item(FormUID)
'            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_AST")
'            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_AST_C0")
'            objMatrix1 = objForm.Items.Item("Mtx").Specific
'            Dim DOCK As SAPbouiCOM.Column = objMatrix1.Columns.Item("Dept")
'            Dim QURY As String = "SELECT DISTINCT ""Code"", ""Name"" FROM OUDP"
'            objMain.objUtilities.MatrixComboBoxValues(DOCK, QURY)
'            objMatrix1.Columns.Item("Dept").DisplayDesc = True
'            objMatrix1.Columns.Item("Dept").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly



'        Catch ex As Exception
'            objForm.Freeze(False)
'            objMain.objApplication.StatusBar.SetText(ex.Message)
'        End Try
'    End Sub

'    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
'        Try
'            If pVal.MenuUID = "SBO_AST" And pVal.BeforeAction = False Then
'                Me.CreateForm()
'                'Me.SetDefault(objForm.UniqueID)
'            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
'                'Me.SetDefault(objForm.UniqueID)
'            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then

'                Me.SetNewLine1(objForm.UniqueID)
'            ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then

'                objMatrix1 = objForm.Items.Item("Mtx").Specific

'                Dim row As Integer = objMatrix1.VisualRowCount

'                If objMatrix1.IsRowSelected(1) <> True And objMatrix1.VisualRowCount < 1 Then

'                    objMatrix1.AddRow()

'                    oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)

'                    objMatrix1.SetLineData(objMatrix1.VisualRowCount)

'                End If

'                If objMatrix1.IsRowSelected(row) = True Then

'                    objMatrix1.DeleteRow(row)

'                Else

'                    For i As Integer = 1 To objMatrix1.VisualRowCount - 1

'                        If objMatrix1.IsRowSelected(i) = True Then

'                            objMatrix1.DeleteRow(i)

'                        End If

'                    Next

'                End If

'                For i As Integer = 1 To objMatrix1.VisualRowCount

'                    objMatrix1.Columns.Item("#").Cells.Item(i).Specific.Value = i

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



'            'Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
'            '    objForm = objMain.objApplication.Forms.Item(FormUID)
'            '        Dim oCFL As SAPbouiCOM.ChooseFromList
'            '        Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
'            '        Dim CFL_Id As String
'            '        CFL_Id = CFLEvent.ChooseFromListUID
'            '        oCFL = objForm.ChooseFromLists.Item(CFL_Id)
'            '        Dim oDT As SAPbouiCOM.DataTable
'            '        oDT = CFLEvent.SelectedObjects
'            '        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
'            '    ' If oCFL.UniqueID = "CFL_1" And pVal.BeforeAction = True Then
'            '    '    ' Me.CFLFilter("CFL_Sale2")
'            '    'End If
'            '    If (Not oDT Is Nothing) And pVal.BeforeAction = False Then
'            '        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
'            '        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
'            '        If oCFL.UniqueID = "CFL_1" Then
'            '            'objForm.Items.Item("ORG").Specific.value = oDT.GetValue("U_NAME", 0)
'            '            'objForm.DataSources.Item("AUTH").Value = oDT.GetValue("U_NAME", 0)
'            '            objMatrix1.Columns.Item("AUTH").Cells.Item(pVal.Row).Specific.Value() = oDT.GetValue("U_NAME", 0)
'            '            objMatrix1.Columns.Item("UKey").Cells.Item(pVal.Row).Specific.Value() = oDT.GetValue("USER_CODE", 0)
'            '            ' objMatrix1.Columns.Item("Dept").Cells.Item(pVal.Row).Specific.Value() = oDT.GetValue("Department", 0)
'            '            'objMatrix1.Columns.Item("AUTH").Cells.Item(pVal.Row).Specific.Value() = oDT.GetValue("U_NAME", 0)
'            '            'objMatrix1.Columns.Item("AUTH").Cells.Item(pVal.Row).Specific.Value() = oDT.GetValue("U_NAME", 0)

'            '        End If
'            '        If oCFL.UniqueID = "CFL_EMP" Then
'            '            ' objMatrix1.Columns.Item("EmpId").Value = oDT.GetValue("empID", 0)
'            '            objMatrix1.Columns.Item("EmpId").Cells.Item(pVal.Row).Specific.Value() = oDT.GetValue("empID", 0)



'            '        End If
'                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
'                    Dim oCFL As SAPbouiCOM.ChooseFromList
'                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
'                    Dim CFL_Id As String
'                    CFL_Id = CFLEvent.ChooseFromListUID
'                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)

'                    Dim oDT As SAPbouiCOM.DataTable
'                    oDT = CFLEvent.SelectedObjects
'                    If oCFL.UniqueID = "CFL_1" And pVal.BeforeAction = True Then
'                        'Me.CFLRoomFilter(Me.m_SBO_Form, "CFL_R", objMatrix.Columns.Item("Build").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())
'                    End If
'                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
'                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
'                        If oCFL.UniqueID = "CFL_1" Then
'                            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
'                            oDBs_Details1.SetValue("U_AUTH", oDBs_Details1.Offset, oDT.GetValue("U_NAME", 0))
'                            oDBs_Details1.SetValue("U_Dept", oDBs_Details1.Offset, oDT.GetValue("Department", 0))
'                            oDBs_Details1.SetValue("U_UKey", oDBs_Details1.Offset, oDT.GetValue("USER_CODE", 0))
'                            objMatrix1.SetLineData(pVal.Row)
'                        End If
'                        If oCFL.UniqueID = "CFL_EMP" Then
'                            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
'                            oDBs_Details1.SetValue("U_EmpId", oDBs_Details1.Offset, oDT.GetValue("empID", 0))
'                            oDBs_Details1.SetValue("U_Dept", oDBs_Details1.Offset, oDT.GetValue("dept", 0))
'                            oDBs_Details1.SetValue("U_UKey", oDBs_Details1.Offset, "")
'                            objMatrix1.SetLineData(pVal.Row)
'                        End If


'                        '    If oCFL.UniqueID = "CFL_2" Then
'                        '        objForm.DataSources.UserDataSources.Item("ORG_2").Value = oDT.GetValue("USER_CODE", 0)

'                        '    End If
'                        '    If oCFL.UniqueID = "CFL_3" Then
'                        '        objForm.DataSources.UserDataSources.Item("AUT_2").Value = oDT.GetValue("USER_CODE", 0)

'                        '    End If
'                        '    If oCFL.UniqueID = "CFL_4" Then
'                        '        objForm.DataSources.UserDataSources.Item("AUT_1").Value = oDT.GetValue("USER_CODE", 0)

'                        '    End If
'                        '    If oCFL.UniqueID = "CFL_5" Then
'                        '        objForm.DataSources.UserDataSources.Item("TEM_1").Value = oDT.GetValue("Code", 0)
'                        '    End If
'                        '    If oCFL.UniqueID = "CFL_6" Then
'                        '        objForm.DataSources.UserDataSources.Item("TEM_2").Value = oDT.GetValue("Code", 0)
'                        '    End If
'                        '    If oCFL.UniqueID = "CFL_7" Then
'                        '        objForm.DataSources.UserDataSources.Item("EMP_1").Value = oDT.GetValue("empID", 0)
'                        '    End If
'                        '    If oCFL.UniqueID = "CFL_8" Then
'                        '        objForm.DataSources.UserDataSources.Item("EMP_2").Value = oDT.GetValue("empID", 0)
'                        '    End If
'                    End If
'                    'Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
'                    '    objForm = objMain.objApplication.Forms.Item(FormUID)
'                    '    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_AST")
'                    '    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_AST_C0")
'                    '    objMatrix1 = objForm.Items.Item("Mtx").Specific

'                    '    If pVal.ItemUID = "Mtx" And pVal.ColUID = "AUTH" Then
'                    '        If objMatrix1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() <> "" Then
'                    '            If pVal.Row = objMatrix1.VisualRowCount Then
'                    '                Me.SetNewLine1(objForm.UniqueID)
'                    '            End If
'                    '        End If
'                    '    End If


'            End Select
'        Catch ex As Exception
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

'#Region " Right Click Event"
'    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)

'        Dim objForm As SAPbouiCOM.Form
'        Dim oMenuItem As SAPbouiCOM.MenuItem
'        Dim oMenus As SAPbouiCOM.Menus
'        objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)
'        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
'        oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
'        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
'        oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_AST")
'        oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_AST_C0")
'        objMatrix1 = objForm.Items.Item("Mtx").Specific

'        Try
'            If eventInfo.FormUID = objForm.UniqueID Then
'                If (eventInfo.BeforeAction = True) Then
'                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then 'SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
'                        '  If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Or objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then

'                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then 'SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
'                            '  objMatrix = objForm.Items.Item("Matx").Specific
'                            objForm.EnableMenu("1292", True)
'                            objForm.EnableMenu("774", True)

'                            If eventInfo.ItemUID = "Mtx" And eventInfo.ColUID = "#" Then
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
'            objForm.Freeze(False)

'            objMain.objApplication.StatusBar.SetText(ex.Message)
'        End Try
'    End Sub
'#End Region


'End Class
'madhu commented
Public Class ClsAPPROVALSTAGES


#Region "       Declaration             "
    Public objForm, objSTRForm, objMIForm As SAPbouiCOM.Form

    Dim oDBs_Head, oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim objMatrix1, MATRIXS As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim str, str1, str2, str3, Query, str4 As String
    Public rs, RsNum As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False
    Dim objutilities As Utilities
    Dim oDt As SAPbouiCOM.DataTable
    Dim oGrid As SAPbouiCOM.Grid
    Dim r1, r2, r3 As RadioButton
    Dim column, column1 As SAPbouiCOM.EditTextColumn

#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("Approvalstages.xml", "SBO_AST", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("SBO_AST", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_AST")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_AST_C0")
            objMatrix1 = objForm.Items.Item("Mtx").Specific
            objForm.EnableMenu("774", True)
            objForm.EnableMenu("1292", True)
            ' objForm.EnableMenu("1293", True)

            objMatrix1.AutoResizeColumns()
            objMatrix1.Clear()
            objForm.DataBrowser.BrowseBy = "SName"


            objMatrix1.FlushToDataSource()

            Me.SetNewLine1(objForm.UniqueID)
            Me.fillDeptCombo(objForm.UniqueID)
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
        End Try

    End Sub
    Sub SetNewLine1(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_AST")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_AST_C0")
            objMatrix1 = objForm.Items.Item("Mtx").Specific

            objMatrix1.AddRow()
            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
            oDBs_Details1.SetValue("U_AUTH", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_EmpId", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_Dept", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_UKey", oDBs_Details1.Offset, "")

            objMatrix1.SetLineData(objMatrix1.VisualRowCount)
            objMatrix1.AutoResizeColumns()
        Catch ex As Exception
            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Sub fillDeptCombo(ByVal FormUID As String)
        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_AST")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_AST_C0")
            objMatrix1 = objForm.Items.Item("Mtx").Specific
            Dim DOCK As SAPbouiCOM.Column = objMatrix1.Columns.Item("Dept")
            Dim QURY As String = "SELECT DISTINCT ""Code"", ""Name"" FROM OUDP"
            objMain.objUtilities.MatrixComboBoxValues(DOCK, QURY)
            objMatrix1.Columns.Item("Dept").DisplayDesc = True
            objMatrix1.Columns.Item("Dept").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly



        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "SBO_AST" And pVal.BeforeAction = False Then
                Me.CreateForm()
                'Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                'Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then

                Me.SetNewLine1(objForm.UniqueID)
            ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then

                objMatrix1 = objForm.Items.Item("Mtx").Specific

                Dim row As Integer = objMatrix1.VisualRowCount

                If objMatrix1.IsRowSelected(1) <> True And objMatrix1.VisualRowCount < 1 Then

                    objMatrix1.AddRow()

                    oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)

                    objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                End If

                If objMatrix1.IsRowSelected(row) = True Then

                    objMatrix1.DeleteRow(row)

                Else

                    For i As Integer = 1 To objMatrix1.VisualRowCount - 1

                        If objMatrix1.IsRowSelected(i) = True Then

                            objMatrix1.DeleteRow(i)

                        End If

                    Next

                End If

                For i As Integer = 1 To objMatrix1.VisualRowCount

                    objMatrix1.Columns.Item("#").Cells.Item(i).Specific.Value = i

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
                'Akhila
                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix1 = objForm.Items.Item("Mtx").Specific

                    If pVal.ItemUID = "Mtx" And pVal.ColUID = "EmpId" And pVal.Before_Action = False Then
                        objMatrix1 = objForm.Items.Item("Mtx").Specific

                        Dim DocumentNumber As String = objMatrix1.Columns.Item("EmpId").Cells.Item(pVal.Row).Specific.value
                        objMain.objApplication.ActivateMenuItem("3590")
                        objMIForm = objMain.objApplication.Forms.GetForm("60100", objMain.objApplication.Forms.ActiveForm.TypeCount)
                        objMIForm.Freeze(True)
                        objMIForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objMIForm.Items.Item("480002078").Specific.Value = DocumentNumber
                        objMIForm.Items.Item("2").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objMIForm.Freeze(False)

                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        If Me.Validation(objForm.UniqueID) = False Then BubbleEvent = False
                    End If

            'Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
            '    objForm = objMain.objApplication.Forms.Item(FormUID)
            '        Dim oCFL As SAPbouiCOM.ChooseFromList
            '        Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
            '        Dim CFL_Id As String
            '        CFL_Id = CFLEvent.ChooseFromListUID
            '        oCFL = objForm.ChooseFromLists.Item(CFL_Id)
            '        Dim oDT As SAPbouiCOM.DataTable
            '        oDT = CFLEvent.SelectedObjects
            '        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
            '    ' If oCFL.UniqueID = "CFL_1" And pVal.BeforeAction = True Then
            '    '    ' Me.CFLFilter("CFL_Sale2")
            '    'End If
            '    If (Not oDT Is Nothing) And pVal.BeforeAction = False Then
            '        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            '        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
            '        If oCFL.UniqueID = "CFL_1" Then
            '            'objForm.Items.Item("ORG").Specific.value = oDT.GetValue("U_NAME", 0)
            '            'objForm.DataSources.Item("AUTH").Value = oDT.GetValue("U_NAME", 0)
            '            objMatrix1.Columns.Item("AUTH").Cells.Item(pVal.Row).Specific.Value() = oDT.GetValue("U_NAME", 0)
            '            objMatrix1.Columns.Item("UKey").Cells.Item(pVal.Row).Specific.Value() = oDT.GetValue("USER_CODE", 0)
            '            ' objMatrix1.Columns.Item("Dept").Cells.Item(pVal.Row).Specific.Value() = oDT.GetValue("Department", 0)
            '            'objMatrix1.Columns.Item("AUTH").Cells.Item(pVal.Row).Specific.Value() = oDT.GetValue("U_NAME", 0)
            '            'objMatrix1.Columns.Item("AUTH").Cells.Item(pVal.Row).Specific.Value() = oDT.GetValue("U_NAME", 0)

            '        End If
            '        If oCFL.UniqueID = "CFL_EMP" Then
            '            ' objMatrix1.Columns.Item("EmpId").Value = oDT.GetValue("empID", 0)
            '            objMatrix1.Columns.Item("EmpId").Cells.Item(pVal.Row).Specific.Value() = oDT.GetValue("empID", 0)



            '        End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)

                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    If oCFL.UniqueID = "CFL_1" And pVal.BeforeAction = True Then
                        'Me.CFLRoomFilter(Me.m_SBO_Form, "CFL_R", objMatrix.Columns.Item("Build").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())
                    End If
                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If oCFL.UniqueID = "CFL_1" Then
                            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
                            oDBs_Details1.SetValue("U_AUTH", oDBs_Details1.Offset, oDT.GetValue("U_NAME", 0))
                            oDBs_Details1.SetValue("U_Dept", oDBs_Details1.Offset, oDT.GetValue("Department", 0))
                            oDBs_Details1.SetValue("U_UKey", oDBs_Details1.Offset, oDT.GetValue("USER_CODE", 0))
                            objMatrix1.SetLineData(pVal.Row)
                        End If
                        If oCFL.UniqueID = "CFL_EMP" Then
                            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
                            oDBs_Details1.SetValue("U_EmpId", oDBs_Details1.Offset, oDT.GetValue("empID", 0))
                            oDBs_Details1.SetValue("U_Dept", oDBs_Details1.Offset, oDT.GetValue("dept", 0))
                            oDBs_Details1.SetValue("U_UKey", oDBs_Details1.Offset, "")
                            objMatrix1.SetLineData(pVal.Row)
                        End If


                        '    If oCFL.UniqueID = "CFL_2" Then
                        '        objForm.DataSources.UserDataSources.Item("ORG_2").Value = oDT.GetValue("USER_CODE", 0)

                        '    End If
                        '    If oCFL.UniqueID = "CFL_3" Then
                        '        objForm.DataSources.UserDataSources.Item("AUT_2").Value = oDT.GetValue("USER_CODE", 0)

                        '    End If
                        '    If oCFL.UniqueID = "CFL_4" Then
                        '        objForm.DataSources.UserDataSources.Item("AUT_1").Value = oDT.GetValue("USER_CODE", 0)

                        '    End If
                        '    If oCFL.UniqueID = "CFL_5" Then
                        '        objForm.DataSources.UserDataSources.Item("TEM_1").Value = oDT.GetValue("Code", 0)
                        '    End If
                        '    If oCFL.UniqueID = "CFL_6" Then
                        '        objForm.DataSources.UserDataSources.Item("TEM_2").Value = oDT.GetValue("Code", 0)
                        '    End If
                        '    If oCFL.UniqueID = "CFL_7" Then
                        '        objForm.DataSources.UserDataSources.Item("EMP_1").Value = oDT.GetValue("empID", 0)
                        '    End If
                        '    If oCFL.UniqueID = "CFL_8" Then
                        '        objForm.DataSources.UserDataSources.Item("EMP_2").Value = oDT.GetValue("empID", 0)
                        '    End If
                    End If
                    'Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    '    objForm = objMain.objApplication.Forms.Item(FormUID)
                    '    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_AST")
                    '    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_AST_C0")
                    '    objMatrix1 = objForm.Items.Item("Mtx").Specific

                    '    If pVal.ItemUID = "Mtx" And pVal.ColUID = "AUTH" Then
                    '        If objMatrix1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() <> "" Then
                    '            If pVal.Row = objMatrix1.VisualRowCount Then
                    '                Me.SetNewLine1(objForm.UniqueID)
                    '            End If
                    '        End If
                    '    End If


            End Select
        Catch ex As Exception
            objForm.Freeze(False)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String)

        objForm = objMain.objApplication.Forms.Item(FormUID)

        objForm = objMain.objApplication.Forms.Item(FormUID)
        oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_AST_C0")
        objMatrix1 = objForm.Items.Item("Mtx").Specific

        If oDBs_Details1.GetValue("U_AUTH", oDBs_Details1.Offset) = "" Then
            objMain.objApplication.StatusBar.SetText(
        "In the ""No. of Approvals Required"" field, enter a number less than or equal to the value in ""Authorizers"" column",
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        '  oDBs_Head = objForm.DataSources.DBDataSources.Item("@FISH_13")
        ' oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@FISH_41")
        ' objMatrix1 = objForm.Items.Item("Mtx").Specific

        'If oDBs_Head.GetValue("U_NAME", oDBs_Head.Offset) = "" Then
        '    objMain.objApplication.StatusBar.SetText("Name is missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '    Return False
        'End If
        Return True


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
        oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_AST")
        oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_AST_C0")
        objMatrix1 = objForm.Items.Item("Mtx").Specific

        Try
            If eventInfo.FormUID = objForm.UniqueID Then
                If (eventInfo.BeforeAction = True) Then
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then 'SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        '  If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Or objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then

                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then 'SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                            '  objMatrix = objForm.Items.Item("Matx").Specific
                            objForm.EnableMenu("1292", True)
                            objForm.EnableMenu("774", True)

                            If eventInfo.ItemUID = "Mtx" And eventInfo.ColUID = "#" Then
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
            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region


End Class


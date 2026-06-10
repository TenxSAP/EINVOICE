Public Class ClsGRIDES

#Region "       Declaration             "
    Public objForm, objSTRForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim str, str1, str2, str3, Query, str4, Query2 As String
    Public rs, RsNum As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False
    Dim objutilities As Utilities
    Dim oDt As SAPbouiCOM.DataTable
    Dim oGrid, oSubGrid As SAPbouiCOM.Grid
    Dim r1, r2, r3 As RadioButton
    Dim column, column1 As SAPbouiCOM.EditTextColumn

#End Region
    Sub CreateForm(ByVal Query2 As String)

        objMain.objUtilities.LoadForm("grides.xml", "GRIDES", ResourceType.Embeded)
        objForm = objMain.objApplication.Forms.GetForm("GRIDES", objMain.objApplication.Forms.ActiveForm.TypeCount)
        'objMain.objApplication.StatusBar.SetText("Please Wait Data is Loading.....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        ' Me.GRIDES(objForm.UniqueID, Query1)
        Me.GRIDES(objForm.UniqueID, Query2)
        'objMain.objApplication.StatusBar.SetText("Data is Loaded Sucessfully.....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        'Try
        ' If pVal.MenuUID = "GRIDES" And pVal.BeforeAction = False Then
        '        Me.CreateForm(Query1)
        '    objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        'End Try
    End Sub
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)

        Select Case pVal.EventType
            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                objForm = objMain.objApplication.Forms.Item(FormUID)
                If pVal.ItemUID = "11" And pVal.BeforeAction = False Then
                    objForm.Close()
                    '        ElseIf pVal.ItemUID = "22" And pVal.BeforeAction = False Then
                    '            oGrid.Rows.ExpandAll()
                End If
                If pVal.ItemUID = "GRIDES" And pVal.ColUID = "RowsHeader" And pVal.BeforeAction = False Then
                    'Me.subgrid(objForm.UniqueID)
                    Me.GridDetails(objForm.UniqueID, pVal)
                End If

        End Select
    End Sub


    Sub GridDetails(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent)


        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("GRIDES").Specific

            Dim AppID As String
            Dim DOCTYPE As String
            'For i As Integer = 1 To oGrid.Rows.Count
            '    If oGrid.Rows.IsSelected(i) = True Then
            Dim Test As Integer = pVal.Row
            AppID = oGrid.DataTable.GetValue("Application Num", Test)
            DOCTYPE = oGrid.DataTable.GetValue("Doc Type", Test)
            If AppID <> "" Then
                ' Me.subgrid(FormUID, AppID)

            End If

            '

            '    End If
            'Next

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    'Sub subgrid(ByVal FormUID As String, ByVal AppID As String)

    '        Try

    '            objForm = objMain.objApplication.Forms.Item(FormUID)
    '            oSubGrid = objForm.Items.Item("SUBGRID").Specific
    '            oSubGrid.DataTable = objForm.DataSources.DataTables.Item("DT_GRD")
    '            oSubGrid.DataTable.Clear()
    '            Query2 = "CALL ""TNX_ApprovalStatusDetailReport""('" & AppID & "')"
    '            oSubGrid.DataTable.ExecuteQuery(Query2)
    '            If oSubGrid.Rows.Count = 0 Then
    '                objMain.objApplication.MessageBox("No records found")
    '                Exit Sub
    '            End If
    '            oSubGrid.AutoResizeColumns()
    '            For i As Integer = 0 To oSubGrid.DataTable.Columns.Count - 1
    '                oSubGrid.Columns.Item(i).Editable = False
    '            Next
    '            'oSubGrid.CollapseLevel = 1

    '            oSubGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

    '        Catch ex As Exception
    '            objMain.objApplication.StatusBar.SetText(ex.Message)
    '        End Try

    '    End Sub

    Sub GRIDES(ByVal FormUID As String, ByVal Query2 As String)
        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("GRIDES").Specific


            oGrid.DataTable = objForm.DataSources.DataTables.Item("DT_GRID")
            oGrid.DataTable.Clear()


            oGrid.DataTable.ExecuteQuery(Query2)


            If oGrid.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("No records found")
                Exit Sub
            End If
            oGrid.AutoResizeColumns()
            For i As Integer = 0 To oGrid.DataTable.Columns.Count - 1
                oGrid.Columns.Item(i).Editable = False
                oGrid.Columns.Item("Approval Doc Status").Editable = True
                'oGrid.Columns.Item("Approval Doc Status").
            Next

            'oGrid.CollapseLevel = 1

            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oGrid = objForm.Items.Item("GRIDES").Specific
            column = oGrid.Columns.Item("Application Num")
            column = oGrid.Columns.Item("Application Num")
            column.LinkedObjectType = "SBO_PRATTENDANCE"


            'Dim AppID As String
            ''For i As Integer = 1 To oGrid.Rows.Count
            '' If oGrid.Rows.IsSelected(i) = True Then
            'Dim Test As Integer = pVal.Row
            'AppID = oGrid.DataTable.GetValue("Application Num", Test)
            'If AppID <> "" Then
            '    Me.subgrid(FormUID, AppID)

            'End If
            'oGrid = objForm.Items.Item("GRIDES").Specific
            'column = oGrid.Columns.Item("Application Num")
            'column.LinkedObjectType = "SBO_PRATTENDANCE"


            'column1 = oGrid.Columns.Item("Client Code")
            'column1.LinkedObjectType = "2"

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String)

        objForm = objMain.objApplication.Forms.Item(FormUID)
        '  oDBs_Head = objForm.DataSources.DBDataSources.Item("@FISH_13")
        ' oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@FISH_41")
        ' objMatrix1 = objForm.Items.Item("Mtx").Specific

        If oDBs_Head.GetValue("U_NAME", oDBs_Head.Offset) = "" Then
            objMain.objApplication.StatusBar.SetText("Name is missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If

    End Function
End Class






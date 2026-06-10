Public Class clsCorporateTexConfig

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
            CheckCorporateTaxConfigSetup()
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCNF")
            objMatrix = objForm.Items.Item("MXT_1").Specific
            'oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNX_GLAUDO"))

            Me.MatrixLoad()

            objForm.Freeze(False)
            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("1293", True)
            '  Me.SetNewLine(objForm.UniqueID)
            ' Me.GlAccountCode()
            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub CheckCorporateTaxConfigSetup()

        Try

            Dim rs As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'Check table exists and configuration available
            rs.DoQuery("SELECT COUNT(*) AS ""Cnt"" FROM ""@TNX_CTAXCNF""")

            If CInt(rs.Fields.Item("Cnt").Value) = 0 Then

                objMain.objApplication.StatusBar.SetText(
            "Corporate Tax Configuration not found. Please configure tax slabs.",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            End If

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
        "Corporate Tax Configuration setup missing. Please check table, fields and UDO creation. Details : " & ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
    'Public Function Validate(ByVal FormUID As String) As Boolean
    '    Dim CustomerCode As SAPbouiCOM.Matrix
    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)

    '        Dim APPID As String = objForm.Items.Item("docnum").Specific.Value
    '        Dim TEMID As String = objForm.Items.Item("tempid").Specific.Value

    '        rs4 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        Dim QRYS As String = "Select ""U_FOUT"" from ""@TNX_PCRT"" where ""U_AppID"" = '" & APPID & "'"
    '        rs4.DoQuery(QRYS)

    '        If TEMID = "FLIER" Then
    '            If rs4.Fields.Item("U_FOUT").Value.ToString() = "No" Then

    '                objMain.objApplication.SetStatusBarMessage("Please change the Flyer-Out status to Yes before approving the document.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
    '                Return False

    '            End If
    '        End If

    '        Return True
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message & "Errors In Validation Function", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    End Try
    'End Function
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)

        Try

            If pVal.MenuUID = "CTAXC" AndAlso pVal.BeforeAction = False Then

                Me.CreateForm()

            ElseIf pVal.MenuUID = "1282" AndAlso pVal.BeforeAction = False Then

                'Add mode logic if required

            ElseIf pVal.MenuUID = "1292" AndAlso pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.TypeEx <> "CTAXC" Then Exit Sub

                objMatrix = CType(objForm.Items.Item("MXT_1").Specific, SAPbouiCOM.Matrix)
                Me.SetNewLine(objForm.UniqueID)


            ElseIf pVal.MenuUID = "1293" And pVal.BeforeAction = False Then
                objMatrix = objForm.Items.Item("MXT_1").Specific
                Dim row As Integer = objMatrix.VisualRowCount
                If objMatrix.IsRowSelected(1) <> True And objMatrix.VisualRowCount < 1 Then
                    objMatrix.AddRow()
                    oDBs_Details.SetValue("DocEntry", oDBs_Details.Offset, objMatrix.VisualRowCount)
                    oDBs_Details.SetValue("Code", oDBs_Details.Offset, objMatrix.VisualRowCount)
                    objMatrix.SetLineData(objMatrix.VisualRowCount)
                End If
                If objMatrix.IsRowSelected(row) = True Then
                    objMatrix.DeleteRow(row)
                Else
                    For i As Integer = 1 To objMatrix.VisualRowCount - 1

                        If objMatrix.IsRowSelected(i) = True Then
                            objMatrix.DeleteRow(i)
                        End If
                    Next
                End If
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.Columns.Item("DocEntry").Cells.Item(i).Specific.Value = i
                    objMatrix.Columns.Item("Code").Cells.Item(i).Specific.Value = i
                Next

            End If

        Catch ex As Exception

            Try
                objForm.Freeze(False)
            Catch
            End Try

            objMain.objApplication.StatusBar.SetText(
            "MenuEvent Error : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
    'Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
    '    Try
    '        If pVal.MenuUID = "CTAXC" And pVal.BeforeAction = False Then
    '            Me.CreateForm()
    '        ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
    '        ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
    '            objMatrix = objForm.Items.Item("MXT_1").Specific
    '            Me.SetNewLine(objForm.UniqueID)
    '        End If

    '        ElseIf pVal.MenuUID = "1293" AndAlso pVal.BeforeAction = True Then

    '        Try
    '            objForm = objMain.objApplication.Forms.ActiveForm

    '            If objForm.TypeEx <> "CTAXC" Then Exit Sub

    '            BubbleEvent = False
    '            objForm.Freeze(True)

    '            objMatrix = CType(objForm.Items.Item("MXT_1").Specific, SAPbouiCOM.Matrix)
    '            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCNF")

    '            Dim selectedRow As Integer =
    '                objMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder)

    '            If selectedRow <= 0 Then
    '                objMain.objApplication.StatusBar.SetText(
    '                    "Please select row to delete",
    '                    SAPbouiCOM.BoMessageTime.bmt_Short,
    '                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '                Exit Try
    '            End If

    '            objMatrix.FlushToDataSource()

    '            oDBs_Head.RemoveRecord(selectedRow - 1)

    '            If oDBs_Head.Size = 0 Then
    '                oDBs_Head.InsertRecord(0)
    '                oDBs_Head.SetValue("DocEntry", 0, "1")
    '                oDBs_Head.SetValue("Code", 0, "1")
    '                oDBs_Head.SetValue("U_MnProfit", 0, "")
    '                oDBs_Head.SetValue("U_MxProfit", 0, "")
    '                oDBs_Head.SetValue("U_TaxPrc", 0, "")
    '                oDBs_Head.SetValue("U_LAccount", 0, "")
    '                oDBs_Head.SetValue("U_EAccount", 0, "")
    '            End If

    '            For i As Integer = 0 To oDBs_Head.Size - 1
    '                oDBs_Head.SetValue("DocEntry", i, (i + 1).ToString())
    '                oDBs_Head.SetValue("Code", i, (i + 1).ToString())
    '            Next

    '            objMatrix.LoadFromDataSource()
    '            objMatrix.AutoResizeColumns()

    '            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
    '                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
    '            End If

    '        Catch ex As Exception
    '            objMain.objApplication.StatusBar.SetText(
    '                "Delete Row Error : " & ex.Message,
    '                SAPbouiCOM.BoMessageTime.bmt_Short,
    '                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '        Finally
    '            Try
    '                objForm.Freeze(False)
    '            Catch
    '            End Try
    '        End Try



    '    Catch ex As Exception
    '        objForm.Freeze(False)
    '        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    End Try



    'End Sub

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

                    If oCFL.UniqueID = "CFL_1" And pVal.BeforeAction = True Then
                        CFLFilterGLAccounts(FormUID, "CFL_1")
                    ElseIf oCFL.UniqueID = "CFL_0" And pVal.BeforeAction = True Then
                        CFLFilterGLAccounts(FormUID, "CFL_0")

                    End If

                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If oCFL.UniqueID = "CFL_1" Then
                            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMatrix.Columns.Item("Code").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Head.SetValue("U_MnProfit", oDBs_Head.Offset, objMatrix.Columns.Item("MnProfit").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Head.SetValue("U_MxProfit", oDBs_Head.Offset, objMatrix.Columns.Item("MxProfit").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Head.SetValue("U_TaxPrc", oDBs_Head.Offset, objMatrix.Columns.Item("TaxPrc").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Head.SetValue("U_EAccount", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))
                            oDBs_Head.SetValue("U_LAccount", oDBs_Head.Offset, objMatrix.Columns.Item("LAccount").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Head.SetValue("U_FINA", oDBs_Head.Offset, objMatrix.Columns.Item("FINA").Cells.Item(pVal.Row).Specific.Value)
                            'oDBs_Head.SetValue("U_TNXSTUS", oDBs_Head.Offset, objMatrix.Columns.Item("TNXSTUS").Cells.Item(pVal.Row).Specific.Value)

                            objMatrix.SetLineData(pVal.Row)
                        End If
                        If oCFL.UniqueID = "CFL_0" Then
                            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMatrix.Columns.Item("Code").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Head.SetValue("U_MnProfit", oDBs_Head.Offset, objMatrix.Columns.Item("MnProfit").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Head.SetValue("U_MxProfit", oDBs_Head.Offset, objMatrix.Columns.Item("MxProfit").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Head.SetValue("U_TaxPrc", oDBs_Head.Offset, objMatrix.Columns.Item("TaxPrc").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Head.SetValue("U_EAccount", oDBs_Head.Offset, objMatrix.Columns.Item("EAccount").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Head.SetValue("U_LAccount", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))
                            oDBs_Head.SetValue("U_FINA", oDBs_Head.Offset, objMatrix.Columns.Item("FINA").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix.SetLineData(pVal.Row)
                            ' CFLFilterGLAccount(FormUID, "CFL_0")
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
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.BeforeAction = True Then

                        If pVal.ItemUID = "1" Then   'OK / Update button

                            objForm = objMain.objApplication.Forms.Item(FormUID)
                            objMatrix = CType(objForm.Items.Item("MXT_1").Specific, SAPbouiCOM.Matrix)

                            For i As Integer = 1 To objMatrix.VisualRowCount
                                If ValidateProfitRange(FormUID, i) = False Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            Next

                            If ValidateGLAccounts(FormUID) = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If

                        End If

                    End If
                    'Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                    '    If pVal.BeforeAction = False Then

                    '        If pVal.ItemUID = "MXT_1" AndAlso
                    '           (pVal.ColUID = "MnProfit" OrElse pVal.ColUID = "MxProfit") Then

                    '            If ValidateProfitRange(FormUID, pVal.Row) = False Then
                    '                BubbleEvent = False
                    '            End If

                    '        End If

                    '    End If
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                    If pVal.BeforeAction = False Then

                        If pVal.ItemUID = "MXT_1" AndAlso
                           (pVal.ColUID = "MnProfit" OrElse
                            pVal.ColUID = "MxProfit") Then

                            objForm = objMain.objApplication.Forms.Item(FormUID)
                            objMatrix = CType(objForm.Items.Item("MXT_1").Specific, SAPbouiCOM.Matrix)

                            Dim MinValue As String =
                                CType(objMatrix.Columns.Item("MnProfit").Cells.Item(pVal.Row).Specific,
                                SAPbouiCOM.EditText).Value.Trim()

                            Dim MaxValue As String =
                                CType(objMatrix.Columns.Item("MxProfit").Cells.Item(pVal.Row).Specific,
                                SAPbouiCOM.EditText).Value.Trim()

                            'Run validation only when both values entered
                            If MinValue <> "" AndAlso MaxValue <> "" Then

                                If ValidateProfitRange(FormUID, pVal.Row) = False Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                            End If

                        End If

                    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Function ToDecimalValue(ByVal value As String) As Decimal
        If value Is Nothing Then Return 0D

        value = value.Trim().Replace(",", "")

        Dim result As Decimal = 0D
        Decimal.TryParse(value, result)

        Return result
    End Function

    Private Function ValidateProfitRange(ByVal FormUID As String, ByVal CurrentRow As Integer) As Boolean

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = CType(objForm.Items.Item("MXT_1").Specific, SAPbouiCOM.Matrix)

            If CurrentRow <= 0 Then Return True

            Dim CurMin As Decimal =
            ToDecimalValue(CType(objMatrix.Columns.Item("MnProfit").Cells.Item(CurrentRow).Specific, SAPbouiCOM.EditText).Value)

            Dim CurMax As Decimal =
            ToDecimalValue(CType(objMatrix.Columns.Item("MxProfit").Cells.Item(CurrentRow).Specific, SAPbouiCOM.EditText).Value)

            If CurMin = 0 OrElse CurMax = 0 Then Return True

            If CurMin >= CurMax Then
                objMain.objApplication.StatusBar.SetText(
                "Minimum Profit should be less than Maximum Profit.",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Return False
            End If

            For i As Integer = 1 To objMatrix.VisualRowCount

                If i = CurrentRow Then Continue For

                Dim MinVal As Decimal =
                ToDecimalValue(CType(objMatrix.Columns.Item("MnProfit").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value)

                Dim MaxVal As Decimal =
                ToDecimalValue(CType(objMatrix.Columns.Item("MxProfit").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value)

                If MinVal = 0 OrElse MaxVal = 0 Then Continue For

                If CurMin <= MaxVal AndAlso CurMax >= MinVal Then

                    objMain.objApplication.StatusBar.SetText(
                    "Profit range overlap found in row " & i & ". Please enter different Minimum and Maximum Profit.",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    Return False
                End If

            Next

            Return True

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "Validate Profit Range Error : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            Return False
        End Try

    End Function

    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
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
            oDBs_Head.SetValue("U_FINA", oDBs_Head.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub CFLFilterGLAccounts(ByVal FormUID As String, ByVal CFL_ID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition

            oChooseFromList = objForm.ChooseFromLists.Item(CFL_ID)

            'Clear existing conditions
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList.SetConditions(emptyCon)

            'Get conditions object
            oConditions = oChooseFromList.GetConditions()

            'Filter only Postable Accounts
            oCondition = oConditions.Add()
            oCondition.Alias = "Postable"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "Y"

            oChooseFromList.SetConditions(oConditions)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message,
                                                 SAPbouiCOM.BoMessageTime.bmt_Short,
                                                 SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
    'Private Function ValidateGLAccounts(ByVal FormUID As String) As Boolean

    '    Try

    '        objForm = objMain.objApplication.Forms.Item(FormUID)

    '        objMatrix =
    '    CType(objForm.Items.Item("MXT_1").Specific,
    '    SAPbouiCOM.Matrix)

    '        For i As Integer = 1 To objMatrix.VisualRowCount

    '            Dim MinProfit As String =
    '        CType(objMatrix.Columns.Item("MnProfit").Cells.Item(i).Specific,
    '        SAPbouiCOM.EditText).Value.Trim()

    '            Dim MaxProfit As String =
    '        CType(objMatrix.Columns.Item("MxProfit").Cells.Item(i).Specific,
    '        SAPbouiCOM.EditText).Value.Trim()

    '            Dim TaxPrc As String =
    '        CType(objMatrix.Columns.Item("TaxPrc").Cells.Item(i).Specific,
    '        SAPbouiCOM.EditText).Value.Trim()

    '            Dim FYear As String =
    '        CType(objMatrix.Columns.Item("FINA").Cells.Item(i).Specific,
    '        SAPbouiCOM.EditText).Value.Trim()

    '            Dim LAccount As String =
    '        CType(objMatrix.Columns.Item("LAccount").Cells.Item(i).Specific,
    '        SAPbouiCOM.EditText).Value.Trim()

    '            Dim EAccount As String =
    '        CType(objMatrix.Columns.Item("EAccount").Cells.Item(i).Specific,
    '        SAPbouiCOM.EditText).Value.Trim()

    '            If MinProfit <> "" OrElse
    '           MaxProfit <> "" OrElse
    '           TaxPrc <> "" Then

    '                '-----------------------------------
    '                ' FINANCIAL YEAR VALIDATION
    '                '-----------------------------------
    '                If FYear = "" Then

    '                    objMain.objApplication.StatusBar.SetText(
    '                "Please enter Financial Year in row " & i & ".",
    '                SAPbouiCOM.BoMessageTime.bmt_Short,
    '                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '                    objMatrix.Columns.Item("FINA").Cells.Item(i).Click()

    '                    Return False

    '                End If

    '                '-----------------------------------
    '                ' LIABILITY ACCOUNT VALIDATION
    '                '-----------------------------------
    '                If LAccount = "" Then

    '                    objMain.objApplication.StatusBar.SetText(
    '                "Please select Liability Account in row " & i & ".",
    '                SAPbouiCOM.BoMessageTime.bmt_Short,
    '                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '                    objMatrix.Columns.Item("LAccount").Cells.Item(i).Click()

    '                    Return False

    '                End If

    '                '-----------------------------------
    '                ' EXPENDITURE ACCOUNT VALIDATION
    '                '-----------------------------------
    '                If EAccount = "" Then

    '                    objMain.objApplication.StatusBar.SetText(
    '                "Please select Expenditure Account in row " & i & ".",
    '                SAPbouiCOM.BoMessageTime.bmt_Short,
    '                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '                    objMatrix.Columns.Item("EAccount").Cells.Item(i).Click()

    '                    Return False

    '                End If

    '            End If

    '        Next

    '        Return True

    '    Catch ex As Exception

    '        objMain.objApplication.StatusBar.SetText(
    '    "GL Account Validation Error : " & ex.Message,
    '    SAPbouiCOM.BoMessageTime.bmt_Short,
    '    SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '        Return False

    '    End Try

    'End Function

    Private Function ValidateGLAccounts(ByVal FormUID As String) As Boolean
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = CType(objForm.Items.Item("MXT_1").Specific, SAPbouiCOM.Matrix)

            For i As Integer = 1 To objMatrix.VisualRowCount

                Dim MinProfit As String =
                CType(objMatrix.Columns.Item("MnProfit").Cells.Item(i).Specific,
                SAPbouiCOM.EditText).Value.Trim()

                Dim MaxProfit As String =
                CType(objMatrix.Columns.Item("MxProfit").Cells.Item(i).Specific,
                SAPbouiCOM.EditText).Value.Trim()

                Dim TaxPrc As String =
                CType(objMatrix.Columns.Item("TaxPrc").Cells.Item(i).Specific,
                SAPbouiCOM.EditText).Value.Trim()

                Dim LAccount As String =
                CType(objMatrix.Columns.Item("LAccount").Cells.Item(i).Specific,
                SAPbouiCOM.EditText).Value.Trim()

                Dim EAccount As String =
                CType(objMatrix.Columns.Item("EAccount").Cells.Item(i).Specific,
                SAPbouiCOM.EditText).Value.Trim()

                Dim FINA As String = ""
Dim oFinCombo As SAPbouiCOM.ComboBox =
CType(objMatrix.Columns.Item("FINA").Cells.Item(i).Specific,
SAPbouiCOM.ComboBox)

                If oFinCombo.Selected IsNot Nothing Then
                    FINA = oFinCombo.Selected.Value.Trim()
                End If
                If MinProfit <> "" OrElse MaxProfit <> "" OrElse TaxPrc <> "" Then

                    If LAccount = "" Then
                        objMain.objApplication.StatusBar.SetText(
                        "Please select Liability Account in row " & i & ".",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        objMatrix.Columns.Item("LAccount").Cells.Item(i).Click()
                        Return False
                    End If

                    If EAccount = "" Then
                        objMain.objApplication.StatusBar.SetText(
                        "Please select Expenditure Account in row " & i & ".",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        objMatrix.Columns.Item("EAccount").Cells.Item(i).Click()
                        Return False
                    End If
                    If FINA = "" Then

                        objMain.objApplication.StatusBar.SetText(
                        "Please select Financial Year in row " & i & ".",
                     SAPbouiCOM.BoMessageTime.bmt_Short,
                         SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        objMatrix.Columns.Item("FINA").Cells.Item(i).Click()

                        Return False

                    End If

                End If

            Next

            Return True

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "GL Account Validation Error : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            Return False
        End Try
    End Function

    Public Sub MatrixLoad()
        Try

            oDBs_Head = objForm.DataSources.DBDataSources.Add("@TNX_CTAXCNF")
            Dim objsectMat As SAPbouiCOM.Matrix = objForm.Items.Item("MXT_1").Specific

            Dim rs1 As String = "SELECT ""DocEntry"",""Code"", ""U_MnProfit"", ""U_MxProfit"", ""U_TaxPrc"", ""U_LAccount"", ""U_EAccount"",""U_FINA"" " &
"FROM ""@TNX_CTAXCNF"" ORDER BY ""Code"" "
            Dim ors1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ors1.DoQuery(rs1)
            If ors1.RecordCount > 0 Then
                objsectMat.Clear()
                For j As Integer = 1 To ors1.RecordCount
                    objsectMat.AddRow()
                    oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.VisualRowCount)
                    oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMatrix.VisualRowCount)
                    oDBs_Head.SetValue("U_MnProfit", oDBs_Head.Offset, ors1.Fields.Item("U_MnProfit").Value)
                    oDBs_Head.SetValue("U_MxProfit", oDBs_Head.Offset, ors1.Fields.Item("U_MxProfit").Value)
                    oDBs_Head.SetValue("U_TaxPrc", oDBs_Head.Offset, ors1.Fields.Item("U_TaxPrc").Value)
                    oDBs_Head.SetValue("U_LAccount", oDBs_Head.Offset, ors1.Fields.Item("U_LAccount").Value)
                    oDBs_Head.SetValue("U_EAccount", oDBs_Head.Offset, ors1.Fields.Item("U_EAccount").Value)
                    oDBs_Head.SetValue("U_FINA", oDBs_Head.Offset, ors1.Fields.Item("U_FINA").Value)
                    objsectMat.SetLineData(objsectMat.VisualRowCount)
                    ors1.MoveNext()
                Next
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            Else
                Me.SetNewLine(objForm.UniqueID)
            End If
            objMatrix.FlushToDataSource()
            ' objsectMat.AutoResizeColumns()
            ' objsectMat.FlushToDataSource()
            'objMatrix.LoadFromDataSource()

        Catch ex As Exception
            oDBs_Head.Freeze(False)
        End Try
    End Sub

End Class

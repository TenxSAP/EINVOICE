Public Class clsApprovalDecision

#Region "       Declaration             "
    Public objForm, objSTRForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim objMatrix1 As SAPbouiCOM.Matrix
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

        objMain.objUtilities.LoadForm("ApprovalD.xml", "ME_ADR", ResourceType.Embeded)
        objForm = objMain.objApplication.Forms.GetForm("ME_ADR", objMain.objApplication.Forms.ActiveForm.TypeCount)

    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_ADR" And pVal.BeforeAction = False Then
                Me.CreateForm()
                'Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                'Me.SetDefault(objForm.UniqueID)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)

        Select Case pVal.EventType
            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                objForm = objMain.objApplication.Forms.Item(FormUID)

                If pVal.ItemUID = "11" And pVal.BeforeAction = False And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                    Dim NoDecisionYet As String = objForm.DataSources.UserDataSources.Item("NDY").Value
                    Dim Approved As String = objForm.DataSources.UserDataSources.Item("APP").Value
                    Dim Rejected As String = objForm.DataSources.UserDataSources.Item("REJ").Value
                    Dim ORG_1 As String = objForm.DataSources.UserDataSources.Item("ORG_1").Value
                    Dim ORG_2 As String = objForm.DataSources.UserDataSources.Item("ORG_2").Value
                    Dim AUT_1 As String = objForm.DataSources.UserDataSources.Item("AUT_1").Value
                    Dim AUT_2 As String = objForm.DataSources.UserDataSources.Item("AUT_2").Value
                    Dim TEM_1 As String = objForm.DataSources.UserDataSources.Item("TEM_1").Value
                    Dim TEM_2 As String = objForm.DataSources.UserDataSources.Item("TEM_2").Value
                    Dim RDATE_1 As String = objForm.DataSources.UserDataSources.Item("RDATE_1").Value
                    Dim RDATE_2 As String = objForm.DataSources.UserDataSources.Item("RDATE_2").Value


                    Dim Query2 As String = "CALL ""TNX_ApprovalDescionReport""('" & NoDecisionYet & "','" & Approved & "'," &
                        "'" & Rejected & "','" & ORG_1 & "'," &
                                  "'" & ORG_2 & "','" & AUT_1 & "','" & AUT_2 & "','" & TEM_1 & "','" & TEM_2 & "','" & RDATE_1 & "','" & RDATE_2 & "')"

                    objMain.ObjGRIDES.CreateForm(Query2)

                End If

            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                objForm = objMain.objApplication.Forms.Item(FormUID)
                Dim oCFL As SAPbouiCOM.ChooseFromList
                Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                Dim CFL_Id As String
                CFL_Id = CFLEvent.ChooseFromListUID
                oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                Dim oDT As SAPbouiCOM.DataTable
                oDT = CFLEvent.SelectedObjects
                objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                ' If oCFL.UniqueID = "CFL_1" And pVal.BeforeAction = True Then
                '    ' Me.CFLFilter("CFL_Sale2")
                'End If
                If (Not oDT Is Nothing) And pVal.BeforeAction = False Then
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                    If oCFL.UniqueID = "CFL_1" Then
                        'objForm.Items.Item("ORG").Specific.value = oDT.GetValue("U_NAME", 0)
                        objForm.DataSources.UserDataSources.Item("ORG_1").Value = oDT.GetValue("USER_CODE", 0)
                    End If
                    If oCFL.UniqueID = "CFL_2" Then
                        objForm.DataSources.UserDataSources.Item("ORG_2").Value = oDT.GetValue("USER_CODE", 0)

                    End If
                    If oCFL.UniqueID = "CFL_3" Then
                        objForm.DataSources.UserDataSources.Item("AUT_1").Value = oDT.GetValue("USER_CODE", 0)

                    End If
                    If oCFL.UniqueID = "CFL_4" Then
                        objForm.DataSources.UserDataSources.Item("AUT_2").Value = oDT.GetValue("USER_CODE", 0)

                    End If
                    If oCFL.UniqueID = "CFL_5" Then
                        objForm.DataSources.UserDataSources.Item("TEM_1").Value = oDT.GetValue("Code", 0)
                    End If
                    If oCFL.UniqueID = "CFL_6" Then
                        objForm.DataSources.UserDataSources.Item("TEM_2").Value = oDT.GetValue("Code", 0)
                    End If
                End If

        End Select
    End Sub

    Function Validation(ByVal FormUID As String)

        objForm = objMain.objApplication.Forms.Item(FormUID)
        '  oDBs_Head = objForm.DataSources.DBDataSources.Item("@FISH_13")
        ' oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@FISH_41")
        ' objMatrix1 = objForm.Items.Item("Mtx").Specific

        'If oDBs_Head.GetValue("U_NAME", oDBs_Head.Offset) = "" Then
        '    objMain.objApplication.StatusBar.SetText("Name is missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '    Return False
        'End If

    End Function
    'Sub GRID(ByVal FormUID As String)
    '    Try
    '    If 
    '            objForm = objMain.objApplication.Forms.Item(FormUID)
    '            oGrid = objForm.Items.Item("GRID_1").Specific

    '            oGrid.DataTable = objForm.DataSources.DataTables.Item("GRID_1")
    '            oGrid.DataTable.Clear()

    '            Dim Query1 As String = "SELECT"" T0.""empID"", T0.""lastName"", T0.""firstName"", T0.""middleName"", T0.""dept"", T0.""branch"" FROM""OHEM T0"""
    '            oGrid.DataTable.ExecuteQuery(Query1)
    '            If oGrid.Rows.Count = 0 Then
    '                objMain.objApplication.MessageBox("No records found")
    '                Exit Sub
    '            End If
    '            oGrid.AutoResizeColumns()
    '            column = oGrid.Columns.Item("GRID_1")
    '            'column.LinkedObjectType = "S"

    '            'column1 = oGrid.Columns.Item("Client Code")
    '            'column1.LinkedObjectType = "2"
    '        End If

    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub

End Class


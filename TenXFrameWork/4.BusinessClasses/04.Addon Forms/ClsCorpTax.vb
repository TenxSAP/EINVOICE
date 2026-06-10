Imports System.Runtime.InteropServices

Public Class ClsCorpTax

#Region "Decliration"

    Public objForm As SAPbouiCOM.Form

    Dim objMatrix As SAPbouiCOM.Matrix
    Dim objutilities As Utilities

    Public oDBs_Head As SAPbouiCOM.DBDataSource

#End Region

    Sub CreateForm()

        Try

            objMain.objUtilities.LoadForm(
            "CorpTax.xml",
            "COTAX",
            ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm(
            "COTAX",
            objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Freeze(True)

            objutilities = New Utilities

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CORP")

            objMatrix = objForm.Items.Item("CMtx").Specific

            Dim rs As SAPbobsCOM.Recordset

            rs = objMain.objCompany.GetBusinessObject(
            SAPbobsCOM.BoObjectTypes.BoRecordset)

            rs.DoQuery(
            "SELECT TOP 1 ""Code"" FROM ""@TNX_CORP""")

            If rs.RecordCount = 0 Then

                objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE

                Me.SetNewLine(objForm.UniqueID)

            Else

                Me.MatrixLoad()

                objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

            End If

            objForm.Freeze(False)

            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("1293", True)


            objMain.objApplication.StatusBar.SetText(
            "Successfully initialized, Please proceed...",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception

            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(
            ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent,
              ByRef BubbleEvent As Boolean)

        Try

            If pVal.MenuUID = "COTX" _
        And pVal.BeforeAction = False Then

                Me.CreateForm()

            ElseIf pVal.MenuUID = "1292" _
        And pVal.BeforeAction = False Then

                objMatrix =
            objForm.Items.Item("CMtx").Specific

                Me.SetNewLine(objForm.UniqueID)

            ElseIf pVal.MenuUID = "1293" And pVal.BeforeAction = False Then
                objMatrix = objForm.Items.Item("CMtx").Specific
                Dim row As Integer = objMatrix.VisualRowCount
                If objMatrix.IsRowSelected(1) <> True And objMatrix.VisualRowCount < 1 Then
                    objMatrix.AddRow()
                    oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.VisualRowCount)
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
                Next

            End If

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
        ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Sub ItemEvent(ByVal FormUID As String,
              ByRef pVal As SAPbouiCOM.ItemEvent,
              ByRef BubbleEvent As Boolean)

        Try

            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.ItemUID = "1" _
                And pVal.BeforeAction = False _
                And pVal.ActionSuccess = True Then

                        Me.MatrixLoad()

                        objForm.Mode =
                    SAPbouiCOM.BoFormMode.fm_OK_MODE

                    End If

                    'Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                    '    If pVal.BeforeAction = False _
                    'And pVal.ItemUID = "CMtx" _
                    'And pVal.ColUID = "Code" Then

                    '        objForm =
                    '    objMain.objApplication.Forms.Item(FormUID)

                    '        objMatrix =
                    '    objForm.Items.Item("CMtx").Specific

                    '        If pVal.Row = objMatrix.VisualRowCount Then

                    '            Dim strCode As String

                    '            strCode =
                    '        CType(
                    '        objMatrix.Columns.Item("Code") _
                    '        .Cells.Item(pVal.Row).Specific,
                    '        SAPbouiCOM.EditText).Value.Trim()

                    '            If strCode <> "" _
                    '        And objForm.Mode <>
                    '        SAPbouiCOM.BoFormMode.fm_FIND_MODE Then

                    '                Me.SetNewLine(objForm.UniqueID)

                    '            End If

                    '        End If

                    '    End If

            End Select

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
        ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CORP")
            objMatrix = objForm.Items.Item("CMtx").Specific
            objMatrix.AddRow()
            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.VisualRowCount)
            oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMatrix.VisualRowCount)
            oDBs_Head.SetValue("U_CMR", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_SO", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_GRPO", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_DCNF", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_CMN", oDBs_Head.Offset, "Open")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    'Sub SetNewLine1(ByVal FormUID As String)

    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        objForm.Freeze(True)

    '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CORP")
    '        objMatrix = objForm.Items.Item("CMtx").Specific

    '        ' objMatrix.FlushToDataSource()

    '        Dim rowIndex As Integer = oDBs_Head.Size

    '        oDBs_Head.InsertRecord(rowIndex)
    '        oDBs_Head.Offset = rowIndex

    '        oDBs_Head.SetValue("DocEntry", rowIndex, (rowIndex + 1).ToString())
    '        oDBs_Head.SetValue("Code", rowIndex, (rowIndex + 1).ToString())
    '        oDBs_Head.SetValue("U_CMR", rowIndex, "")
    '        oDBs_Head.SetValue("U_SO", rowIndex, "")
    '        oDBs_Head.SetValue("U_GRPO", rowIndex, "")
    '        oDBs_Head.SetValue("U_DCNF", rowIndex, "")
    '        oDBs_Head.SetValue("U_CMN", rowIndex, "Open")

    '        objMatrix.LoadFromDataSource()

    '        For i As Integer = 1 To objMatrix.VisualRowCount
    '            objMatrix.Columns.Item("DocEntry").Cells.Item(i).Specific.Value = i.ToString()
    '            objMatrix.Columns.Item("Code").Cells.Item(i).Specific.Value = i.ToString()
    '        Next

    '        objMatrix.AutoResizeColumns()

    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(
    '    "SetNewLine Error : " & ex.Message,
    '    SAPbouiCOM.BoMessageTime.bmt_Short,
    '    SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '    Finally
    '        Try
    '            objForm.Freeze(False)
    '        Catch
    '        End Try
    '    End Try

    'End Sub


    'Sub SetNewLine(ByVal FormUID As String)

    '    Try

    '        objForm =
    '    objMain.objApplication.Forms.Item(FormUID)
    '        objForm.Freeze(True)

    '        oDBs_Head =
    '    objForm.DataSources.DBDataSources.Item("@TNX_CORP")

    '        objMatrix =
    '    objForm.Items.Item("CMtx").Specific

    '        Dim rowIndex As Integer

    '        rowIndex = oDBs_Head.Size

    '        objMatrix.FlushToDataSource()

    '        oDBs_Head.InsertRecord(rowIndex)

    '        oDBs_Head.Offset = rowIndex

    '        oDBs_Head.SetValue("Code", rowIndex, (rowIndex + 1).ToString())
    '        oDBs_Head.SetValue("U_CMR", rowIndex, "")
    '        oDBs_Head.SetValue("U_SO", rowIndex, "")
    '        oDBs_Head.SetValue("U_GRPO", rowIndex, "")
    '        oDBs_Head.SetValue("U_DCNF", rowIndex, "")
    '        oDBs_Head.SetValue("U_CMN", rowIndex, "Open")

    '        objMatrix.LoadFromDataSource()

    '        For i As Integer = 1 To objMatrix.VisualRowCount

    '            objMatrix.Columns.Item("DocEntry") _
    '        .Cells.Item(i).Specific.Value = i

    '        Next

    '        objMatrix.AutoResizeColumns()
    '        objForm.Freeze(False)

    '    Catch ex As Exception

    '        objMain.objApplication.StatusBar.SetText(
    '    ex.Message)

    '    End Try

    'End Sub
    'Public Sub MatrixLoad1()
    '    Try
    '        objForm.Freeze(True)

    '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CORP")
    '        objMatrix = objForm.Items.Item("CMtx").Specific

    '        Dim rs1 As String =
    '    "SELECT ""DocEntry"", ""Code"", ""U_CMR"", ""U_SO"", ""U_GRPO"", ""U_DCNF"", ""U_CMN"" " &
    '    "FROM ""@TNX_CORP"" ORDER BY ""Code"""

    '        Dim ors1 As SAPbobsCOM.Recordset =
    '    objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        ors1.DoQuery(rs1)

    '        If ors1.RecordCount > 0 Then
    '            objMatrix.Clear()

    '            For j As Integer = 1 To ors1.RecordCount
    '                objMatrix.AddRow()

    '                oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.VisualRowCount)
    '                oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMatrix.VisualRowCount)
    '                oDBs_Head.SetValue("U_CMR", oDBs_Head.Offset, ors1.Fields.Item("U_CMR").Value.ToString())
    '                oDBs_Head.SetValue("U_SO", oDBs_Head.Offset, ors1.Fields.Item("U_SO").Value.ToString())
    '                oDBs_Head.SetValue("U_GRPO", oDBs_Head.Offset, ors1.Fields.Item("U_GRPO").Value.ToString())
    '                oDBs_Head.SetValue("U_DCNF", oDBs_Head.Offset, ors1.Fields.Item("U_DCNF").Value.ToString())
    '                oDBs_Head.SetValue("U_CMN", oDBs_Head.Offset, ors1.Fields.Item("U_CMN").Value.ToString())

    '                objMatrix.SetLineData(objMatrix.VisualRowCount)
    '                ors1.MoveNext()
    '            Next

    '            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
    '        Else
    '            Me.SetNewLine(objForm.UniqueID)
    '        End If

    '        'objMatrix.FlushToDataSource()
    '        objForm.Freeze(False)

    '    Catch ex As Exception
    '        objForm.Freeze(False)
    '        objMain.objApplication.StatusBar.SetText(
    '    ex.Message,
    '    SAPbouiCOM.BoMessageTime.bmt_Short,
    '    SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    End Try
    'End Sub
    Public Sub MatrixLoad()

        Try
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CORP")

            objMatrix =
        objForm.Items.Item("CMtx").Specific

            Dim rs1 As String

            rs1 =
        "SELECT ""Code"", ""U_CMR"", ""U_SO"", " &
        """U_GRPO"", ""U_DCNF"", ""U_CMN"" " &
        "FROM ""@TNX_CORP"" ORDER BY ""Code"""

            Dim ors1 As SAPbobsCOM.Recordset

            ors1 =
        objMain.objCompany.GetBusinessObject(
        SAPbobsCOM.BoObjectTypes.BoRecordset)

            ors1.DoQuery(rs1)

            oDBs_Head.Clear()

            If ors1.RecordCount > 0 Then

                For j As Integer = 0 To ors1.RecordCount - 1

                    oDBs_Head.InsertRecord(j)

                    oDBs_Head.Offset = j

                    oDBs_Head.SetValue(
                "Code",
                j,
                ors1.Fields.Item("Code").Value.ToString())

                    oDBs_Head.SetValue(
                "U_CMR",
                j,
                ors1.Fields.Item("U_CMR").Value.ToString())

                    oDBs_Head.SetValue(
                "U_SO",
                j,
                ors1.Fields.Item("U_SO").Value.ToString())

                    oDBs_Head.SetValue(
                "U_GRPO",
                j,
                ors1.Fields.Item("U_GRPO").Value.ToString())

                    oDBs_Head.SetValue(
                "U_DCNF",
                j,
                CType(
                ors1.Fields.Item("U_DCNF").Value,
                DateTime).ToString("yyyyMMdd"))

                    oDBs_Head.SetValue(
                "U_CMN",
                j,
                ors1.Fields.Item("U_CMN").Value.ToString())

                    ors1.MoveNext()

                Next

                objMatrix.LoadFromDataSource()

                For i As Integer = 1 To objMatrix.VisualRowCount

                    objMatrix.Columns.Item("DocEntry") _
                .Cells.Item(i).Specific.Value = i

                Next

                objForm.Mode =
            SAPbouiCOM.BoFormMode.fm_OK_MODE

            Else

                Me.SetNewLine(objForm.UniqueID)

            End If

            objMatrix.AutoResizeColumns()
            objForm.Freeze(False)

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
        ex.Message)

        End Try

    End Sub
End Class

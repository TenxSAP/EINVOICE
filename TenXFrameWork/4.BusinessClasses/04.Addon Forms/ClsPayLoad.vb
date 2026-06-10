Imports Newtonsoft.Json.Linq
Imports SAPbobsCOM
Imports SAPbouiCOM

Public Class ClsPayLoad

#Region "Declaration"

    Public objForm As SAPbouiCOM.Form

#End Region

#Region "Create Form"

    Public Sub CreateForm()

        Try
            'Load XML Form
            objMain.objUtilities.LoadForm("PayResponse.xml", "PAYR", ResourceType.Embeded)

            'Get Opened Form
            objForm = objMain.objApplication.Forms.GetForm("PAYR", objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Visible = True
            objForm.Select()
            '  LoadPayloadTextToGrid(objForm.UniqueID)

            objMain.objApplication.StatusBar.SetText(
                "Form opened successfully...",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
                ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

#End Region

#Region "Menu Event"

    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent,
                         ByRef BubbleEvent As Boolean)

        Try
            If pVal.MenuUID = "PAYR" AndAlso pVal.BeforeAction = False Then

                CreateForm()
            End If

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
                ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

#End Region

#Region "Item Event"

    Public Sub ItemEvent(ByVal FormUID As String,
                         ByRef pVal As SAPbouiCOM.ItemEvent,
                         ByRef BubbleEvent As Boolean)

        Try
            If pVal.FormTypeEx <> "PAYR" Then Exit Sub

            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.BeforeAction = False Then

                        objForm = objMain.objApplication.Forms.Item(FormUID)

                        If pVal.ItemUID = "2" Then
                            objForm.Close()
                        End If

                    End If

            End Select

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
                ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Private Function ItemExists(ByVal oForm As SAPbouiCOM.Form,
                             ByVal sItemUID As String) As Boolean
        Try
            Dim oItem As SAPbouiCOM.Item
            oItem = oForm.Items.Item(sItemUID)
            Return True
        Catch
            Return False
        End Try
    End Function



    'Public Sub LoadPayloadTextToGrid(ByVal FormUID As String)

    '    Dim oForm As SAPbouiCOM.Form = Nothing

    '    Try
    '        oForm = objMain.objApplication.Forms.Item(FormUID)

    '        If ItemExists(oForm, "Grid_1") = False Then
    '            objMain.objApplication.StatusBar.SetText(
    '            "Grid_1 not found",
    '            SAPbouiCOM.BoMessageTime.bmt_Short,
    '            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Exit Sub
    '        End If

    '        oForm.Freeze(True)

    '        Dim oGrid As SAPbouiCOM.Grid =
    '        CType(oForm.Items.Item("Grid_1").Specific, SAPbouiCOM.Grid)

    '        Dim oDT As SAPbouiCOM.DataTable

    '        Try
    '            oDT = oForm.DataSources.DataTables.Item("DT_0")
    '            oDT.Clear()
    '        Catch
    '            oDT = oForm.DataSources.DataTables.Add("DT_0")
    '        End Try

    '        If oDT.Columns.Count = 0 Then
    '            oDT.Columns.Add("Field", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 300)
    '            oDT.Columns.Add("Value", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 500)
    '            oDT.Columns.Add("Validation", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
    '        End If

    '        Dim payloadText As String = ReadPayloadText()
    '        payloadText = payloadText.Trim()

    '        If payloadText = "" Then
    '            Throw New Exception("PayLoad.txt is empty")
    '        End If

    '        Dim lines() As String = payloadText.Replace(vbCrLf, vbLf).Split(vbLf)

    '        Dim i As Integer = 0

    '        For Each line As String In lines

    '            line = line.Trim()

    '            If line = "" Then Continue For

    '            Dim fld As String = ""
    '            Dim val As String = ""

    '            If line.Contains("=") Then
    '                Dim pos As Integer = line.IndexOf("=")
    '                fld = line.Substring(0, pos).Trim()
    '                val = line.Substring(pos + 1).Trim()

    '            ElseIf line.Contains(":") Then
    '                Dim pos As Integer = line.IndexOf(":")
    '                fld = line.Substring(0, pos).Trim()
    '                val = line.Substring(pos + 1).Trim()

    '            Else
    '                fld = line
    '                val = ""
    '            End If

    '            If fld <> "" Then
    '                AddRow(oDT, i, fld, val)
    '            End If

    '        Next

    '        oGrid.DataTable = oDT
    '        oGrid.Columns.Item("Field").TitleObject.Caption = "Field"
    '        oGrid.Columns.Item("Value").TitleObject.Caption = "Value"
    '        oGrid.Columns.Item("Validation").TitleObject.Caption = "Validation"
    '        oGrid.AutoResizeColumns()

    '        oForm.Freeze(False)

    '    Catch ex As Exception

    '        If oForm IsNot Nothing Then
    '            Try
    '                oForm.Freeze(False)
    '            Catch
    '            End Try
    '        End If

    '        objMain.objApplication.StatusBar.SetText(
    '        "Payload load error: " & ex.Message,
    '        SAPbouiCOM.BoMessageTime.bmt_Short,
    '        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '    End Try

    'End Sub
    'Public Function ReadPayloadText() As String

    '    Try
    '        Dim assembly = System.Reflection.Assembly.GetExecutingAssembly()

    '        Dim resourceName As String =
    '        assembly.GetManifestResourceNames().
    '        FirstOrDefault(Function(x) x.EndsWith("PayloadData.txt"))

    '        If String.IsNullOrEmpty(resourceName) Then
    '            Throw New Exception("PayloadData.txt not found in embedded resources")
    '        End If

    '        Using stream = assembly.GetManifestResourceStream(resourceName)
    '            Using reader As New IO.StreamReader(stream)
    '                Return reader.ReadToEnd()
    '            End Using
    '        End Using

    '    Catch ex As Exception
    '        Throw New Exception("Error reading PayloadData.txt: " & ex.Message)
    '    End Try

    'End Function




    'Public Sub LoadFieldValues(ByVal FormUID As String)

    '    Dim oSource As SAPbouiCOM.Form = Nothing
    '    Dim oTarget As SAPbouiCOM.Form = Nothing
    '    Dim oGrid As SAPbouiCOM.Grid = Nothing
    '    Dim oDT As SAPbouiCOM.DataTable = Nothing

    '    Try

    '        oSource = objMain.objApplication.Forms.Item(FormUID)
    '        oTarget = objMain.objApplication.Forms.ActiveForm

    '        If oTarget Is Nothing Then Exit Sub
    '        If ItemExists(oTarget, "Grid_1") = False Then Exit Sub

    '        oTarget.Freeze(True)

    '        Dim postStatus As String = ""
    '        Dim totalAmt As String = ""
    '        Dim taxAmt As String = ""
    '        Dim invDate As String = ""
    '        Dim jsonText As String = ""

    '        postStatus = CType(oSource.Items.Item("TXTPSA").Specific, SAPbouiCOM.EditText).Value.Trim()
    '        totalAmt = CType(oSource.Items.Item("TXTAM").Specific, SAPbouiCOM.EditText).Value.Trim()
    '        taxAmt = CType(oSource.Items.Item("TXTTAX").Specific, SAPbouiCOM.EditText).Value.Trim()
    '        invDate = CType(oSource.Items.Item("TXTDATE").Specific, SAPbouiCOM.EditText).Value.Trim()
    '        jsonText = CType(oSource.Items.Item("TXTPAY").Specific, SAPbouiCOM.EditText).Value.Trim()

    '        oGrid = CType(oTarget.Items.Item("Grid_1").Specific, SAPbouiCOM.Grid)
    '        oDT = oTarget.DataSources.DataTables.Item("DT_0")

    '        If oDT.Columns.Count = 0 Then
    '            oDT.Columns.Add("Field", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 300)
    '            oDT.Columns.Add("Value", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 500)
    '            oDT.Columns.Add("Validation", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
    '        End If

    '        oDT.Rows.Clear()

    '        Dim i As Integer = 0

    '        ------------------------------------
    '         Header Rows
    '        ------------------------------------
    '        AddRow(oDT, i, "Posting Status", postStatus)
    '        AddRow(oDT, i, "Total Amount", totalAmt)
    '        AddRow(oDT, i, "Tax Amount", taxAmt)
    '        AddRow(oDT, i, "Date", invDate)

    '        ------------------------------------
    '         JSON Rows
    '        ------------------------------------
    '        If jsonText <> "" Then

    '            Dim jObj As JObject = JObject.Parse(jsonText)
    '            AddJsonRows(jObj, "", oDT, i)

    '        End If

    '        ------------------------------------
    '         Bind Grid
    '        ------------------------------------
    '        oGrid.DataTable = oDT

    '        oGrid.Columns.Item("Field").TitleObject.Caption = "Field"
    '        oGrid.Columns.Item("Value").TitleObject.Caption = "Value"
    '        oGrid.Columns.Item("Validation").TitleObject.Caption = "Validation"

    '        oGrid.AutoResizeColumns()

    '        ------------------------------------
    '         Posting Status Color only
    '        ------------------------------------
    '        If UCase(postStatus).Contains("SUCCESS") Then
    '            oGrid.CommonSetting.SetCellBackColor(1, 2, RGB(0, 255, 0))
    '        Else
    '            oGrid.CommonSetting.SetCellBackColor(1, 2, RGB(255, 0, 0))
    '        End If

    '        ------------------------------------
    '         Blank Values Red Color
    '        ------------------------------------
    '        Dim r As Integer

    '        For r = 1 To oDT.Rows.Count

    '            Dim val As String = oDT.GetValue("Value", r - 1).ToString.Trim()



    '        Next

    '        oTarget.Freeze(False)

    '    Catch ex As Exception

    '        If oTarget IsNot Nothing Then
    '            Try
    '                oTarget.Freeze(False)
    '            Catch
    '            End Try
    '        End If

    '        objMain.objApplication.StatusBar.SetText(
    '    ex.Message,
    '    SAPbouiCOM.BoMessageTime.bmt_Short,
    '    SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '    End Try

    'End Sub





    'Private Sub AddRow(ByRef oDT As SAPbouiCOM.DataTable,
    '               ByRef i As Integer,
    '               ByVal fld As String,
    '               ByVal val As String)

    '    oDT.Rows.Add(1)

    '    oDT.SetValue("Field", i, fld)
    '    oDT.SetValue("Value", i, val)

    '    All Fields mandatory
    '    oDT.SetValue("Validation", i, "Mandatory")

    '    i += 1

    'End Sub


    'Private Sub AddJsonRows(ByVal token As JToken,
    '                    ByVal prefix As String,
    '                    ByRef oDT As SAPbouiCOM.DataTable,
    '                    ByRef i As Integer)

    '    Dim child As JToken

    '    If token.Type = JTokenType.Object Then

    '        For Each child In token.Children(Of JProperty)()

    '            Dim newKey As String = child.Path.Replace(".", "_")
    '            AddJsonRows(child.First, newKey, oDT, i)

    '        Next

    '    ElseIf token.Type = JTokenType.Array Then

    '        Dim arr As JToken

    '        For Each arr In token.Children()
    '            AddJsonRows(arr, prefix, oDT, i)
    '        Next

    '    Else

    '        AddRow(oDT, i, prefix, token.ToString())

    '    End If

    'End Sub
    Public Sub LoadFieldValues(ByVal FormUID As String)

        Dim oSource As SAPbouiCOM.Form = Nothing
        Dim oTarget As SAPbouiCOM.Form = Nothing
        Dim oGrid As SAPbouiCOM.Grid = Nothing
        Dim oDT As SAPbouiCOM.DataTable = Nothing

        Try

            oSource = objMain.objApplication.Forms.Item(FormUID)
            oTarget = objMain.objApplication.Forms.ActiveForm

            If oTarget Is Nothing Then Exit Sub

            If ItemExists(oTarget, "Grid_1") = False Then
                objMain.objApplication.StatusBar.SetText(
                "Grid_1 not found",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            oTarget.Freeze(True)

            '-----------------------------------
            ' Read Main Fields
            '-----------------------------------
            Dim postStatus As String = ""
            Dim totalAmt As String = ""
            Dim taxAmt As String = ""
            Dim invDate As String = ""
            Dim jsonText As String = ""

            postStatus = CType(oSource.Items.Item("TXTPSA").Specific, SAPbouiCOM.EditText).Value.Trim()
            totalAmt = CType(oSource.Items.Item("TXTAM").Specific, SAPbouiCOM.EditText).Value.Trim()
            taxAmt = CType(oSource.Items.Item("TXTTAX").Specific, SAPbouiCOM.EditText).Value.Trim()
            invDate = CType(oSource.Items.Item("TXTDATE").Specific, SAPbouiCOM.EditText).Value.Trim()
            jsonText = CType(oSource.Items.Item("TXTPAY").Specific, SAPbouiCOM.EditText).Value.Trim()

            '-----------------------------------
            ' Grid
            '-----------------------------------
            oGrid = CType(oTarget.Items.Item("Grid_1").Specific, SAPbouiCOM.Grid)
            oDT = oTarget.DataSources.DataTables.Item("DT_0")

            If oDT.Columns.Count = 0 Then
                oDT.Columns.Add("Field", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 250)
                oDT.Columns.Add("Value", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 500)
            End If

            oDT.Rows.Clear()

            Dim i As Integer = 0

            '-----------------------------------
            ' Add Required Header Fields First
            '-----------------------------------
            oDT.Rows.Add(1)
            oDT.SetValue("Field", i, "Posting Status")
            oDT.SetValue("Value", i, postStatus)
            i += 1

            oDT.Rows.Add(1)
            oDT.SetValue("Field", i, "Total Amount")
            oDT.SetValue("Value", i, totalAmt)
            i += 1

            oDT.Rows.Add(1)
            oDT.SetValue("Field", i, "Tax Amount")
            oDT.SetValue("Value", i, taxAmt)
            i += 1

            oDT.Rows.Add(1)
            oDT.SetValue("Field", i, "Date")
            oDT.SetValue("Value", i, invDate)
            i += 1

            '-----------------------------------
            ' JSON Response Data
            '-----------------------------------
            jsonText = Replace(jsonText, """Invoice"":", "")
            jsonText = Replace(jsonText, "{", "")
            jsonText = Replace(jsonText, "}", "")
            jsonText = Replace(jsonText, "[", "")
            jsonText = Replace(jsonText, "]", "")

            Dim arr() As String
            arr = Split(jsonText, ",")

            Dim line As String
            Dim fld As String
            Dim val As String

            For Each line In arr

                If InStr(line, ":") > 0 Then

                    fld = Trim(Replace(Split(line, ":")(0), """", ""))
                    val = Trim(Mid(line, InStr(line, ":") + 1))
                    val = Replace(val, """", "")

                    If fld <> "" Then
                        oDT.Rows.Add(1)
                        oDT.SetValue("Field", i, fld)
                        oDT.SetValue("Value", i, val)
                        i += 1
                    End If

                End If

            Next

            '-----------------------------------
            ' Bind Grid
            '-----------------------------------
            oGrid.DataTable = oDT
            oGrid.AutoResizeColumns()

            oGrid.Columns.Item("Field").TitleObject.Caption = "Field"
            oGrid.Columns.Item("Value").TitleObject.Caption = "Value"

            oTarget.Freeze(False)

        Catch ex As Exception

            If oTarget IsNot Nothing Then
                Try
                    oTarget.Freeze(False)
                Catch
                End Try
            End If

            objMain.objApplication.StatusBar.SetText(
            ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

#End Region

End Class
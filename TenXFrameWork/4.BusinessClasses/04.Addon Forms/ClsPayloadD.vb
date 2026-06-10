Imports SAPbouiCOM
Imports System.IO
Imports System.Reflection

Public Class ClsPayloadD

    Public objForm As SAPbouiCOM.Form

    '-------------------------------
    ' Open Form from Menu
    '-------------------------------
    Public Sub CreateForm()

        Try
            objMain.objUtilities.LoadForm("PayLoadD.xml", "PAYLD", ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm(
                "PAYLD",
                objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Visible = True
            objForm.Select()

            LoadPayloadTextToGrid(objForm.UniqueID)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
                ex.Message,
                BoMessageTime.bmt_Short,
                BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
    Public Sub ItemEvent(ByVal FormUID As String,
                         ByRef pVal As SAPbouiCOM.ItemEvent,
                         ByRef BubbleEvent As Boolean)

        Try
            If pVal.FormTypeEx <> "PAYLD" Then Exit Sub

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


    '-------------------------------
    ' Menu Event
    '-------------------------------
    Public Sub MenuEvent(ByRef pVal As MenuEvent,
                         ByRef BubbleEvent As Boolean)

        If pVal.BeforeAction = False Then
            If pVal.MenuUID = "PAYLD" Then

                Dim obj As New ClsPayloadD
                obj.CreateForm()

            End If
        End If

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
    Private Sub AddRow(ByRef oDT As SAPbouiCOM.DataTable,
                   ByRef i As Integer,
                   ByVal fld As String,
                   ByVal val As String)

        oDT.Rows.Add(1)
        oDT.SetValue("Field", i, fld)
        oDT.SetValue("Value", i, val)
        oDT.SetValue("Validation", i, "Mandatory")
        i += 1

    End Sub



    Public Sub LoadPayloadTextToGrid(ByVal FormUID As String)

        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            oForm = objMain.objApplication.Forms.Item(FormUID)

            If ItemExists(oForm, "Grid_1") = False Then
                objMain.objApplication.StatusBar.SetText(
                "Grid_1 not found",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            oForm.Freeze(True)

            Dim oGrid As SAPbouiCOM.Grid =
            CType(oForm.Items.Item("Grid_1").Specific, SAPbouiCOM.Grid)

            Dim oDT As SAPbouiCOM.DataTable

            Try
                oDT = oForm.DataSources.DataTables.Item("DT_0")
                oDT.Clear()
            Catch
                oDT = oForm.DataSources.DataTables.Add("DT_0")
            End Try

            If oDT.Columns.Count = 0 Then
                oDT.Columns.Add("Field", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 300)
                oDT.Columns.Add("Value", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 500)
                oDT.Columns.Add("Validation", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            End If

            Dim payloadText As String = ReadPayloadText()
            payloadText = payloadText.Trim()

            If payloadText = "" Then
                Throw New Exception("PayLoad.txt is empty")
            End If

            Dim lines() As String = payloadText.Replace(vbCrLf, vbLf).Split(vbLf)

            Dim i As Integer = 0

            For Each line As String In lines

                line = line.Trim()

                If line = "" Then Continue For

                Dim fld As String = ""
                Dim val As String = ""

                If line.Contains("=") Then
                    Dim pos As Integer = line.IndexOf("=")
                    fld = line.Substring(0, pos).Trim()
                    val = line.Substring(pos + 1).Trim()

                ElseIf line.Contains(":") Then
                    Dim pos As Integer = line.IndexOf(":")
                    fld = line.Substring(0, pos).Trim()
                    val = line.Substring(pos + 1).Trim()

                Else
                    fld = line
                    val = ""
                End If

                If fld <> "" Then
                    AddRow(oDT, i, fld, val)
                End If

            Next

            oGrid.DataTable = oDT
            oGrid.Columns.Item("Field").TitleObject.Caption = "Field"
            oGrid.Columns.Item("Value").TitleObject.Caption = "Value"
            oGrid.Columns.Item("Validation").TitleObject.Caption = "Validation"
            oGrid.AutoResizeColumns()

            oForm.Freeze(False)

        Catch ex As Exception

            If oForm IsNot Nothing Then
                Try
                    oForm.Freeze(False)
                Catch
                End Try
            End If

            objMain.objApplication.StatusBar.SetText(
            "Payload load error: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
    Public Function ReadPayloadText() As String

        Try
            Dim assembly = System.Reflection.Assembly.GetExecutingAssembly()

            Dim resourceName As String =
            assembly.GetManifestResourceNames().
            FirstOrDefault(Function(x) x.EndsWith("PayloadData.txt"))

            If String.IsNullOrEmpty(resourceName) Then
                Throw New Exception("PayloadData.txt not found in embedded resources")
            End If

            Using stream = assembly.GetManifestResourceStream(resourceName)
                Using reader As New IO.StreamReader(stream)
                    Return reader.ReadToEnd()
                End Using
            End Using

        Catch ex As Exception
            Throw New Exception("Error reading PayloadData.txt: " & ex.Message)
        End Try

    End Function



End Class
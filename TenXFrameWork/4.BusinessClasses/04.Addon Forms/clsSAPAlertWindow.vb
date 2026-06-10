Public Class clsSAPAlertWindow
    Public objform As SAPbouiCOM.Form

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.BeforeAction = True Then
                        'Me.AddItems(objform.UniqueID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Row > 0 And pVal.BeforeAction = False Then
                        Dim omatrix As SAPbouiCOM.Matrix = objform.Items.Item("6").Specific
                        Dim Title As String = omatrix.Columns.Item("V_0").Title
                        If Title = "Application Number" Then
                            Dim objMatrix As SAPbouiCOM.Matrix
                            objMatrix = objform.Items.Item("6").Specific
                            If pVal.ItemUID = "6" And (pVal.ColUID = "0" Or pVal.ColUID = "V_0") Then
                                Dim Code As String = objMatrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value.ToString()
                                Dim Stage As String = objMatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value.ToString()

                                objMain.ObjVIEW.CreateForm(Code, Stage)
                            End If
                        End If
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

End Class


Imports System
Imports System.IO
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
Imports System.Security.Cryptography.Xml
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Xml.XmlReader
Imports System.Xml.XmlWriter
Imports Newtonsoft.Json
Imports Org.BouncyCastle.Crypto
Imports Org.BouncyCastle.Crypto.Parameters
Imports Org.BouncyCastle.Crypto.Signers
Imports Org.BouncyCastle.Math.EC
Imports Org.BouncyCastle.OpenSsl
Imports Org.BouncyCastle.Security
Imports SAPbobsCOM
Imports SAPbouiCOM
Imports TenXFrameWork.FlickINVModel
Imports TenXFrameWork.Zatca.Models
Imports TenXFrameWork.Zatca_CSID.Models
Imports uuids
Imports Zatca.EInvoice.SDK
Imports Zatca.EInvoice.SDK.Contracts.Models
Imports ZXing
Imports ECCurve = System.Security.Cryptography.ECCurve
Imports FormatException = System.FormatException
Public Class ClsCorporateTaxCalcu

    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix

    Dim objutilities As Utilities
    Public rs, RsNum As SAPbobsCOM.Recordset
    Private CancelDocEntry As String = ""
    Private CancelDocNum As String = ""

    'Form Creation
    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("CorporateTaxCalculation.xml", "CTAXCAL", ResourceType.Embeded) 'Loading Form.xml file
            objForm = objMain.objApplication.Forms.GetForm("CTAXCAL", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Dim oStatic As SAPbouiCOM.StaticText

            objForm.Items.Item("COTURL").TextStyle = 4

            oStatic = CType(objForm.Items.Item("COTURL").Specific, SAPbouiCOM.StaticText)

            Dim corporatevaturl As String = oStatic.Caption

            oStatic.Caption = corporatevaturl

            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL_C2")
            '   oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL_C0")
            objMatrix = objForm.Items.Item("MXT_1").Specific
            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNX_CTCAUDO"))
            oDBs_Head.SetValue("U_DAT", 0, DateTime.Now.ToString("yyyyMMdd"))
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DAT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DAT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Item_11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Item_11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("FDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("FDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("TDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("TDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("JEPOSTD").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("JEPOSTD").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Branch").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Branch").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("YEAR").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("YEAR").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("2").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            ' SetButtonStatus(objForm.UniqueID)
            Me.SetDefault(objForm.UniqueID)
            SetAutoManagedFields(objForm.UniqueID)
            SetButtonStatus(objForm.UniqueID)
            SetBranchFieldStatus(objForm.UniqueID)
            LoadBranches(objForm.UniqueID)
            '  objForm.Items.Item("Item_14").Enabled = False
            ' objForm.Items.Item("Item_4").Enabled = False   'Submit
            ' objForm.Items.Item("Item_7").Enabled = False   'Post JE
            'objForm.Items.Item("Item_14").Enabled = False  'Profit & Loss
            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("1293", True)
            objForm.EnableMenu("519", True)
            objForm.EnableMenu("520", True)

            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL_C2")
            ' oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL_C0")
            'objMatrix = objForm.Items.Item("Matrix").Specific
            'Dim GetDocNum As String = "Select Max(""DocNum"")+1 From ""@TNX_CTAXCAL"" "
            'Dim rsDocNum As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'rsDocNum.DoQuery(GetDocNum)
            'oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNX_CTCALUDO"))
            oDBs_Head.SetValue("U_DAT", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("U_DST", 0, "Open")
            Me.SetNewLine(objForm.UniqueID)
            AutoDocentryNumber(objForm.UniqueID)
            'objMatrix.AutoResizeColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            Dim objForm As SAPbouiCOM.Form
            Dim oMenuItem As SAPbouiCOM.MenuItem
            Dim oMenus As SAPbouiCOM.Menus
            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
            oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL_C2")
            objMatrix = objForm.Items.Item("MTX_1").Specific

            If eventInfo.FormUID = objForm.UniqueID Then
                If (eventInfo.BeforeAction = True) Then
                    If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE And CStr(oDBs_Head.GetValue("U_JENo", 0)).Trim = "" Then

                        If eventInfo.ItemUID = "MTX_1" And eventInfo.ColUID = "#" Then
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Delete Row") = False Then
                                    oCreationPackage.UniqueID = "Delete Row"
                                    oCreationPackage.String = "Delete Row"
                                    oCreationPackage.Enabled = True
                                    oMenus.AddEx(oCreationPackage)
                                End If

                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        Else
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Delete Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Delete Row")
                                End If
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try

                        End If

                    Else

                        Try
                            oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                            oMenus = oMenuItem.SubMenus
                            If oMenus.Exists("Delete Row") = True Then
                                objMain.objApplication.Menus.RemoveEx("Delete Row")
                            End If
                            If oMenus.Exists("Close") = True Then
                                objMain.objApplication.Menus.RemoveEx("Close")
                            End If
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL_C2")
            objMatrix = CType(objForm.Items.Item("MXT_1").Specific, SAPbouiCOM.Matrix)

            objMatrix.AddRow()

            Dim rowIndex As Integer = oDBs_Details.Size - 1

            oDBs_Details.SetValue("LineId", rowIndex, objMatrix.VisualRowCount.ToString())
            oDBs_Details.SetValue("U_TPH", rowIndex, "")
            oDBs_Details.SetValue("U_FNM", rowIndex, "")
            oDBs_Details.SetValue("U_ATCD", rowIndex, "")
            oDBs_Details.SetValue("U_FTR", rowIndex, "")
            ' oDBs_Details.SetValue("U_BPN", rowIndex, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()
            'objMatrix.LoadFromDataSource()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "CTAXCAL" And pVal.BeforeAction = False Then
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                objMatrix = objForm.Items.Item("MXT_1").Specific
                Me.SetDefault(objForm.UniqueID)
                LoadBranches(objForm.UniqueID)
                SetBranchFieldStatus(objForm.UniqueID)
                SetButtonStatus(objForm.UniqueID)

                'oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNXLMNUDO"))
                ' Me.CreateForm()
                ' Me.SetNewLine(objForm.UniqueID)
            ElseIf pVal.MenuUID = "519" AndAlso pVal.BeforeAction = True Then

                Try

                    objForm = objMain.objApplication.Forms.ActiveForm

                    If objForm.TypeEx <> "CTAXCAL" Then Exit Sub

                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL")

                    Dim frs As SAPbobsCOM.Recordset =
                    objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'CORPTAXREPORT' AND ""Type"" = 'C'")

                    If frs.RecordCount = 0 Then

                        objMain.objApplication.MessageBox(
                        "CORPTAXREPORT Layout not found. Please import Crystal Layout with same name.",
                        0,
                        "OK")

                    Else

                        objMain.objApplication.ActivateMenuItem(
                        frs.Fields.Item(0).Value.ToString())

                        Dim CrForm As SAPbouiCOM.Form
                        Dim oedt As SAPbouiCOM.EditText

                        CrForm = objMain.objApplication.Forms.ActiveForm

                        oedt = CrForm.Items.Item("1000003").Specific

                        oedt.Value = oDBs_Head.GetValue("DocEntry", 0)

                        CrForm.Items.Item("1").Click(
                        SAPbouiCOM.BoCellClickType.ct_Regular)

                    End If

                Catch ex As Exception

                    objMain.objApplication.StatusBar.SetText(
                    "Preview Error : " & ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                End Try


                '====================================================
                ' PRINT LAYOUT
                '====================================================
            ElseIf pVal.MenuUID = "520" AndAlso pVal.BeforeAction = True Then

                Try

                    Dim LayoutSelection As Integer =
                    objMain.objApplication.MessageBox(
                    "Please select layout for printing",
                    1,
                    "Corporate Tax Report",
                    "")

                    If LayoutSelection = 1 Then

                        objForm = objMain.objApplication.Forms.ActiveForm

                        If objForm.TypeEx <> "CTAXCAL" Then Exit Sub

                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL")

                        Dim frs As SAPbobsCOM.Recordset =
                        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'CORPTAXREPORT' AND ""Type"" = 'C'")

                        If frs.RecordCount = 0 Then

                            objMain.objApplication.MessageBox(
                            "CORPTAXREPORT Layout not found. Please import Crystal Layout with same name.",
                            0,
                            "OK")

                        Else

                            objMain.objApplication.ActivateMenuItem(
                            frs.Fields.Item(0).Value.ToString())

                            Dim CrForm As SAPbouiCOM.Form
                            Dim oedt As SAPbouiCOM.EditText

                            CrForm = objMain.objApplication.Forms.ActiveForm

                            oedt = CrForm.Items.Item("1000003").Specific

                            oedt.Value = oDBs_Head.GetValue("DocEntry", 0)

                            CrForm.Items.Item("1").Click(
                            SAPbouiCOM.BoCellClickType.ct_Regular)

                        End If

                    End If

                Catch ex As Exception

                    objMain.objApplication.StatusBar.SetText(
                    "Print Error : " & ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                End Try
            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                ' Me.SetDefault(objForm.UniqueID)
                objMatrix = objForm.Items.Item("MXT_1").Specific
                ' oDBs_Head.SetValue("U_DAT", 0, DateTime.Now.ToString("yyyyMMdd"))
                'objMatrix1 = objForm.Items.Item("MXT_2").Specific
                'objMatrix2 = objForm.Items.Item("MXT_3").Specific
                Me.SetNewLine(objForm.UniqueID)
                SetButtonStatus(objForm.UniqueID)


            ElseIf pVal.MenuUID = "1284" AndAlso pVal.BeforeAction = True Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.TypeEx <> "CTAXCAL" Then Exit Sub

                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL")

                CancelDocEntry = oDBs_Head.GetValue("DocEntry", 0).Trim()
                CancelDocNum = oDBs_Head.GetValue("DocNum", 0).Trim()

                If ReverseCorporateTaxJE(objForm.UniqueID) = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

            ElseIf pVal.MenuUID = "1284" AndAlso pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If CancelDocEntry <> "" Then

                    Dim rsUpdate As SAPbobsCOM.Recordset =
                        CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
                        SAPbobsCOM.Recordset)

                    Dim Qry As String =
                        "UPDATE ""@TNX_CTAXCAL"" " &
                        "SET ""U_DST"" = 'N' " &
                        "WHERE ""DocEntry"" = '" & CancelDocEntry & "'"

                    rsUpdate.DoQuery(Qry)

                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    objForm.Items.Item("DocNum").Specific.Value = CancelDocNum
                    objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                    CancelDocEntry = ""
                    CancelDocNum = ""

                    SetButtonStatus(objForm.UniqueID)

                End If

            ElseIf pVal.MenuUID = "1293" And pVal.BeforeAction = False Then
                objMatrix = objForm.Items.Item("MXT_1").Specific
                Dim row As Integer = objMatrix.VisualRowCount
                If objMatrix.IsRowSelected(1) <> True And objMatrix.VisualRowCount < 1 Then
                    objMatrix.AddRow()
                    oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
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
                    objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.Value = i
                Next


            ElseIf pVal.MenuUID = "1281" AndAlso pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.TypeEx = "CTAXCAL" Then
                    SetButtonStatus(objForm.UniqueID)
                End If

                'ElseIf pVal.MenuUID = "1293" AndAlso pVal.BeforeAction = True Then

                '    Try
                '        objForm = objMain.objApplication.Forms.ActiveForm

                '        If objForm.TypeEx <> "CTAXCAL" Then Exit Sub

                '        BubbleEvent = False
                '        objForm.Freeze(True)

                '        objMatrix = CType(objForm.Items.Item("MXT_1").Specific, SAPbouiCOM.Matrix)
                '        oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL_C2")

                '        Dim selectedRow As Integer =
                'objMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder)

                '        If selectedRow <= 0 Then
                '            objMain.objApplication.StatusBar.SetText(
                '    "Please select row to delete",
                '    SAPbouiCOM.BoMessageTime.bmt_Short,
                '    SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                '            Exit Try
                '        End If

                '        'First sync current matrix data
                '        objMatrix.FlushToDataSource()

                '        'Remove selected row directly from DBDataSource
                '        oDBs_Details.RemoveRecord(selectedRow - 1)

                '        'If no rows, add one blank row
                '        If oDBs_Details.Size = 0 Then
                '            oDBs_Details.InsertRecord(0)
                '            oDBs_Details.SetValue("LineId", 0, "1")
                '            oDBs_Details.SetValue("U_TPH", 0, "")
                '            oDBs_Details.SetValue("U_FNM", 0, "")
                '            oDBs_Details.SetValue("U_ATCD", 0, "")
                '            oDBs_Details.SetValue("U_FTR", 0, "")
                '        End If

                '        'Re-sequence line numbers: 1,2,3...
                '        For i As Integer = 0 To oDBs_Details.Size - 1
                '            oDBs_Details.SetValue("LineId", i, (i + 1).ToString())
                '        Next

                '        'Reload matrix
                '        objMatrix.LoadFromDataSource()
                '        objMatrix.AutoResizeColumns()

                '        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                '            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                '        End If

                '        SetButtonStatus(objForm.UniqueID)

                '    Catch ex As Exception
                '        objMain.objApplication.StatusBar.SetText(
                '"Delete Row Error : " & ex.Message,
                'SAPbouiCOM.BoMessageTime.bmt_Short,
                'SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                '    Finally
                '        Try
                '            objForm.Freeze(False)
                '        Catch
                '        End Try
                '    End Try

                'End If

                'ElseIf pVal.MenuUID = "1293" AndAlso pVal.BeforeAction = False Then

                '    Try
                '        objForm = objMain.objApplication.Forms.ActiveForm

                '        If objForm.TypeEx <> "CTAXCAL" Then Exit Sub

                '        objForm.Freeze(True)

                '        objMatrix = CType(objForm.Items.Item("MXT_1").Specific, SAPbouiCOM.Matrix)
                '        oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL_C2")

                '        Dim row As Integer = objMatrix.VisualRowCount

                '        If row < 1 Then
                '            objMatrix.AddRow()
                '            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, "1")
                '            objMatrix.SetLineData(1)
                '            Exit Try
                '        End If

                '        If objMatrix.IsRowSelected(row) = True Then

                '            objMatrix.DeleteRow(row)

                '        Else

                '            For i As Integer = objMatrix.VisualRowCount To 1 Step -1

                '                If objMatrix.IsRowSelected(i) = True Then
                '                    objMatrix.DeleteRow(i)
                '                    Exit For
                '                End If

                '            Next

                '        End If

                '        If objMatrix.VisualRowCount = 0 Then

                '            objMatrix.AddRow()
                '            oDBs_Details.SetValue("LineId", 0, "1")
                '            oDBs_Details.SetValue("U_TPH", 0, "")
                '            oDBs_Details.SetValue("U_FNM", 0, "")
                '            oDBs_Details.SetValue("U_ATCD", 0, "")
                '            oDBs_Details.SetValue("U_FTR", 0, "")
                '            objMatrix.SetLineData(1)

                '        End If

                '        For i As Integer = 1 To objMatrix.VisualRowCount

                '            CType(objMatrix.Columns.Item("LineId").Cells.Item(i).Specific,
                '      SAPbouiCOM.EditText).Value = i.ToString()

                '        Next

                '        objMatrix.FlushToDataSource()

                '        For i As Integer = 0 To oDBs_Details.Size - 1
                '            oDBs_Details.SetValue("LineId", i, (i + 1).ToString())
                '        Next

                '        objMatrix.LoadFromDataSource()
                '        objMatrix.AutoResizeColumns()

                '        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                '            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                '        End If

                '    Catch ex As Exception

                '        objMain.objApplication.StatusBar.SetText(
                '"Delete Row Error : " & ex.Message,
                'SAPbouiCOM.BoMessageTime.bmt_Short,
                'SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                '    Finally
                '        Try
                '            objForm.Freeze(False)
                '        Catch
                '        End Try
                '    End Try

            End If






        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Function ReverseCorporateTaxJE(ByVal FormUID As String) As Boolean

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL")

            Dim DocNum As String = oDBs_Head.GetValue("DocNum", 0).Trim()
            Dim jeAbsEntry As String = oDBs_Head.GetValue("U_JENo", 0).Trim()

            If jeAbsEntry = "" Then
                Return True
            End If

            Dim Ans As Integer = objMain.objApplication.MessageBox(
            "JE already posted for this document. Do you want to reverse JE and cancel document?",
            1,
            "Yes",
            "No")

            If Ans <> 1 Then
                Return False
            End If

            Dim oJournalEntry As SAPbobsCOM.JournalEntries =
            CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries),
            SAPbobsCOM.JournalEntries)

            If oJournalEntry.GetByKey(CInt(jeAbsEntry)) Then

                Dim cancelResult As Integer = oJournalEntry.Cancel()

                Dim errCode As Integer = 0
                Dim errMsg As String = ""
                objMain.objCompany.GetLastError(errCode, errMsg)

                If cancelResult = 0 _
            OrElse errMsg.ToLower().Contains("already reversed") _
            OrElse errMsg.ToLower().Contains("already cancelled") Then

                    objMain.objApplication.StatusBar.SetText(
                    "Reverse Entry posted successfully for JE : " & jeAbsEntry,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                    Return True

                Else

                    objMain.objApplication.StatusBar.SetText(
                    "Error cancelling Journal Entry [" & jeAbsEntry & "] : " & errMsg,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    Return False

                End If

            Else

                objMain.objApplication.StatusBar.SetText(
                "Journal Entry not found for AbsEntry : " & jeAbsEntry,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Return False

            End If

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
            "Exception in ReverseCorporateTaxJE : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            Return False

        End Try

    End Function
    'Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
    '    Try
    '        If pVal.MenuUID = "CTAXCAL" And pVal.BeforeAction = False Then
    '            Me.CreateForm()

    '        End If

    '    Catch ex As Exception
    '        objForm.Freeze(False)
    '        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    End Try
    'End Sub
    Public Sub AutoDocentryNumber(ByVal FormUID As String)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL")

            If oDBs_Head.Size = 0 Then
                oDBs_Head.InsertRecord(0)
            End If

            Dim oRsDocNum As SAPbobsCOM.Recordset =
            CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
            SAPbobsCOM.Recordset)

            Dim Query1 As String =
            "SELECT IFNULL(MAX(""DocNum""),0)+1 AS ""DocNum"" FROM ""@TNX_CTAXCAL"""

            oRsDocNum.DoQuery(Query1)

            oDBs_Head.SetValue("DocNum", 0, oRsDocNum.Fields.Item("DocNum").Value.ToString())


            Dim rsAppId As SAPbobsCOM.Recordset =
            CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
            SAPbobsCOM.Recordset)

            Dim str As String =
            "SELECT 'C' || LPAD(" &
            "TO_NVARCHAR(IFNULL(MAX(TO_INTEGER(REPLACE(""U_AIPD"", 'C', ''))),0)+1), 6, '0') AS ""AppId"" " &
            "FROM ""@TNX_CTAXCAL"""

            rsAppId.DoQuery(str)

            oDBs_Head.SetValue("U_AIPD", 0, rsAppId.Fields.Item("AppId").Value.ToString())

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "AutoDocentryNumber Error : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
    Sub ApprovalTrigger(ByVal DocType As String, ByVal AppID As String, ByVal Table As String, ByVal AppIDField As String, ByVal AppStatField As String, ByVal Document As String)
        Dim ff As Integer
        Dim draftquery As String
        Dim approverid As String
        Dim docdate As String = objForm.Items.Item("DAT").Specific.Value
        Dim rsapp As SAPbobsCOM.Recordset
        Dim posrs As SAPbobsCOM.Recordset

        ff = objMain.objApplication.MessageBox("Do You want Submit the Application?", 1, "Ok", "Cancel")
        If ff = 1 Then
            Dim result As Integer
            result = objMain.objApplication.MessageBox("You can not change the Application after you Sending for Approval,continue", 1, "Yes", "No")
            If result = 1 Then
                rsapp = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim str As String = "SELECT T0.""Code"" As ""TemplateID"",T1.""U_Name"" As ""Originator"",S1.""Code"" As ""Stage"",S2.""U_UKey"" As ""Authorizers"" " &
            " ,S1.""U_NAP"" As ""NoOfAuth"",S1.""U_NRJ"" As ""NoOfRej"",T0.""CreateDate"",(Case When T3.""LineId""='1' Then 'S' Else 'O' End) As ""Status"" FROM ""@SBO_APPHDR"" T0 " &
            " INNER JOIN ""@SBO_APPREQ"" T1 ON T1.""Code""=T0.""Code"" " &
            " INNER JOIN ""@SBO_APPDOC"" T2 ON T2.""Code"" = T0.""Code"" " &
            " INNER JOIN ""@SBO_APPAUT"" T3 ON T3.""Code"" = T0.""Code"" " &
            " INNER JOIN ""@SBO_AST"" S1 ON T3.""U_M3_1""=S1.""Code"" " &
            " INNER JOIN ""@SBO_AST_C0"" S2 ON S1.""Code""=S2.""Code"" " &
            " WHERE ""U_Active"" = 'Y' AND T1.""U_Name"" = '" & objMain.objCompany.UserName & "' AND T2.""U_" & DocType & """  = 'Y' AND IFNULL(S2.""U_UKey"",'')<>''  order by T3.""LineId""  ;"
                rsapp.DoQuery(str)

                If rsapp.RecordCount > 0 Then
                    posrs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    draftquery = " SELECT * FROM ""@SBO_OWDD"" WHERE ""U_statusN"" IN ('S','A') AND ""U_Isdraft"" IN ('Y','N') AND ""U_docnum"" = '" & AppID & "';"
                    posrs.DoQuery(draftquery)
                    If posrs.RecordCount > 0 Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(posrs)
                        objMain.objApplication.StatusBar.SetText("Document Already sent for approval", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    Else
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(posrs)
                        ' Dim ofor As SAPbouiCOM.Form
                        'Me.SendMessageAlert(DocType, objMain.objCompany.UserName)

                        'ofor = New DraftProcedureStages(DocType, AppID, objMain.objCompany.UserName, rsapp, Table, AppIDField, AppStatField, Document, docdate)
                        ' ofor = Nothing
                        'objMain.ObjDraftProcedure.CreateForm()
                        objMain.ObjDraftProcedure.CreateForm(objForm, DocType, AppID, objMain.objCompany.UserName, rsapp, Table, AppIDField, AppStatField, Document)
                        ''.CreateForm(objForm.UniqueID, objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value)
                        'objMain.objApplication.ActivateMenuItem("1304")
                    End If
                Else
                    objMain.objApplication.StatusBar.SetText("Login User Don't have Approvals Please Define Approvals and Add the Document", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        End If
    End Sub
    Private Sub SetLoadButtonStatus(ByVal FormUID As String)

        Try

            objForm =
        objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head =
        objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL")

            '-----------------------------------
            ' FIND MODE = ENABLE
            '-----------------------------------
            If objForm.Mode =
        SAPbouiCOM.BoFormMode.fm_FIND_MODE Then

                objForm.Items.Item("Item_11").Enabled = True
                Exit Sub

            End If

            '-----------------------------------
            ' CHECK PROFIT PERIOD
            '-----------------------------------
            Dim Profit As String =
        oDBs_Head.GetValue("U_PPeriod", 0).Trim()

            '-----------------------------------
            ' OK MODE = DISABLE
            '-----------------------------------
            If objForm.Mode =
        SAPbouiCOM.BoFormMode.fm_OK_MODE _
        AndAlso Profit <> "" _
        AndAlso Profit <> "0" _
        AndAlso Profit <> "0.00" Then

                objForm.Items.Item("Item_11").Enabled = False

            Else

                objForm.Items.Item("Item_11").Enabled = True

            End If

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
        "SetLoadButtonStatus Error : " & ex.Message)

        End Try

    End Sub
    Private Sub SetButtonStatus(ByVal FormUID As String)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL")

            Dim JENo As String = oDBs_Head.GetValue("U_JENo", 0).Trim()
            JENo = JENo.Replace(",", "").Trim()

            If JENo.Contains(".") Then
                JENo = JENo.Split("."c)(0)
            End If

            If JENo = "0" OrElse JENo = "0.0" OrElse JENo = "0.00" Then
                JENo = ""
            End If

            Dim DocStatus As String = oDBs_Head.GetValue("U_DST", 0).Trim().ToUpper()

            'JE already posted means always disable both buttons
            If JENo <> "" Then
                objForm.Items.Item("Item_4").Enabled = False   'Submit
                objForm.Items.Item("Item_7").Enabled = False   'Post JE
                objForm.Items.Item("Item_11").Enabled = False  'Load
                objForm.Items.Item("Item_14").Enabled = True   'Profit & Loss
                Exit Sub
            End If

            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                objForm.Items.Item("Item_11").Enabled = True
                objForm.Items.Item("Item_14").Enabled = True
                objForm.Items.Item("Item_4").Enabled = False
                objForm.Items.Item("Item_7").Enabled = False
                Exit Sub
            End If

            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                objForm.Items.Item("Item_11").Enabled = False
                objForm.Items.Item("Item_14").Enabled = False
                objForm.Items.Item("Item_4").Enabled = False
                objForm.Items.Item("Item_7").Enabled = False
                Exit Sub
            End If

            objForm.Items.Item("Item_11").Enabled = False
            objForm.Items.Item("Item_14").Enabled = True

            If DocStatus = "" OrElse DocStatus = "OPEN" OrElse DocStatus = "O" Then
                objForm.Items.Item("Item_4").Enabled = True
                objForm.Items.Item("Item_7").Enabled = False
            Else
                objForm.Items.Item("Item_4").Enabled = False
                objForm.Items.Item("Item_7").Enabled = True
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("SetButtonStatus Error : " & ex.Message)
        End Try

    End Sub
    'Private Sub SetButtonStatus(ByVal FormUID As String)

    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL")

    '        Dim JENo As String = oDBs_Head.GetValue("U_JENo", 0).Trim()
    '        Dim DocStatus As String = oDBs_Head.GetValue("U_DST", 0).Trim().ToUpper()

    '        'FIND MODE
    '        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then

    '            objForm.Items.Item("Item_11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
    '            objForm.Items.Item("Item_4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
    '            objForm.Items.Item("Item_7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
    '            objForm.Items.Item("Item_14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

    '            objForm.Items.Item("DST").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
    '            objForm.Items.Item("AIPD").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

    '            Exit Sub
    '        End If

    '        'ADD MODE
    '        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

    '            objForm.Items.Item("Item_11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)   'Load
    '            objForm.Items.Item("Item_14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)   'Profit & Loss

    '            objForm.Items.Item("Item_4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)   'Submit
    '            objForm.Items.Item("Item_7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)   'Post JE

    '            Exit Sub
    '        End If
    '        'OK / UPDATE MODE
    '        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE _
    '    OrElse objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then

    '            objForm.Items.Item("Item_11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

    '            objForm.Items.Item("DST").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
    '            objForm.Items.Item("AIPD").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

    '            'Submit Button
    '            If JENo = "" AndAlso
    '           (DocStatus = "" OrElse DocStatus = "OPEN" OrElse DocStatus = "O") Then

    '                objForm.Items.Item("Item_4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

    '            Else

    '                objForm.Items.Item("Item_4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

    '            End If

    '            'Post JE Button
    '            If JENo = "" AndAlso
    '           Not (DocStatus = "" OrElse DocStatus = "OPEN" OrElse DocStatus = "O") Then

    '                objForm.Items.Item("Item_7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

    '            Else

    '                objForm.Items.Item("Item_7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

    '            End If

    '        End If

    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText("SetButtonStatus Error : " & ex.Message)
    '    End Try

    'End Sub
    'Private Sub SetButtonStatus(ByVal FormUID As String)

    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL")

    '        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

    '            objForm.Items.Item("Item_11").Enabled = True    'Load
    '            objForm.Items.Item("Item_4").Enabled = False    'Submit
    '            objForm.Items.Item("Item_7").Enabled = False    'Post JE

    '            Exit Sub

    '        End If

    '        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

    '            objForm.Items.Item("Item_11").Enabled = False   'Load Disable after Add / OK Mode
    '            objForm.Items.Item("Item_4").Enabled = True     'Submit Enable

    '            Dim JENo As String = oDBs_Head.GetValue("U_JENo", 0).Trim()

    '            If JENo = "" Then
    '                objForm.Items.Item("Item_7").Enabled = True     'Post JE Enable
    '            Else
    '                objForm.Items.Item("Item_7").Enabled = False    'Post JE Disable
    '            End If

    '        End If

    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(
    '        "SetButtonStatus Error : " & ex.Message,
    '        SAPbouiCOM.BoMessageTime.bmt_Short,
    '        SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    End Try

    'End Sub


    Private Function ValidateCorporateTaxAdding(ByVal oForm As SAPbouiCOM.Form) As Boolean

        Dim rs As SAPbobsCOM.Recordset = Nothing

        Try
            Dim PostingDate As String = GetEditValue(oForm, "JEPOSTD")
            If PostingDate = "" Then
                objMain.objApplication.StatusBar.SetText(
                "Please enter posting date",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Dim FromDate As String = GetEditValue(oForm, "FDate")
            Dim ToDate As String = GetEditValue(oForm, "TDate")

            If FromDate = "" OrElse ToDate = "" Then
                objMain.objApplication.StatusBar.SetText(
                "Please enter From Date and To Date",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If



            Dim dtFrom As DateTime
            Dim dtTo As DateTime

            If Not DateTime.TryParseExact(FromDate, "yyyyMMdd",
            System.Globalization.CultureInfo.InvariantCulture,
            Globalization.DateTimeStyles.None, dtFrom) Then

                objMain.objApplication.StatusBar.SetText(
                "Invalid From Date",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If Not DateTime.TryParseExact(ToDate, "yyyyMMdd",
            System.Globalization.CultureInfo.InvariantCulture,
            Globalization.DateTimeStyles.None, dtTo) Then

                objMain.objApplication.StatusBar.SetText(
                "Invalid To Date",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If dtFrom > dtTo Then
                objMain.objApplication.StatusBar.SetText(
                "From Date should not be greater than To Date",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If dtFrom = dtTo Then
                objMain.objApplication.StatusBar.SetText(
                "From Date and To Date should not be same",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Dim currentDocNum As String = ""
            Try
                currentDocNum = CType(oForm.Items.Item("DocNum").Specific, SAPbouiCOM.EditText).Value.Trim()
            Catch
                currentDocNum = ""
            End Try

            rs = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
                   SAPbobsCOM.Recordset)

            Dim sql As String =
            "SELECT TOP 1 ""DocNum"", ""U_FDate"", ""U_TDate"" " &
            "FROM ""@TNX_CTAXCAL"" " &
            "WHERE IFNULL(""Canceled"",'N') = 'N' " &
            "AND ""U_FDate"" IS NOT NULL " &
            "AND ""U_TDate"" IS NOT NULL " &
            "AND TO_DATE('" & dtFrom.ToString("yyyy-MM-dd") & "','YYYY-MM-DD') <= ""U_TDate"" " &
            "AND TO_DATE('" & dtTo.ToString("yyyy-MM-dd") & "','YYYY-MM-DD') >= ""U_FDate"" "

            If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                If currentDocNum <> "" Then
                    sql &= " AND ""DocNum"" <> '" & currentDocNum & "' "
                End If
            End If

            rs.DoQuery(sql)

            If rs.RecordCount > 0 Then

                Dim existDoc As String = rs.Fields.Item("DocNum").Value.ToString()

                objMain.objApplication.StatusBar.SetText(
                "Selected period overlaps with existing Corporate Tax Document No : " & existDoc,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Return False
            End If

            Return True

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
            "Date Validation Error : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            Return False

        Finally
            If rs IsNot Nothing Then
                Try
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
                Catch
                End Try
                rs = Nothing
            End If
        End Try

    End Function

    Private Function ValidateCorporateTaxDates(ByVal oForm As SAPbouiCOM.Form) As Boolean

        Dim rs As SAPbobsCOM.Recordset = Nothing

        Try
            Dim FromDate As String = GetEditValue(oForm, "FDate")
            Dim ToDate As String = GetEditValue(oForm, "TDate")

            If FromDate = "" OrElse ToDate = "" Then
                objMain.objApplication.StatusBar.SetText(
                "Please enter From Date and To Date",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If



            Dim dtFrom As DateTime
            Dim dtTo As DateTime

            If Not DateTime.TryParseExact(FromDate, "yyyyMMdd",
            System.Globalization.CultureInfo.InvariantCulture,
            Globalization.DateTimeStyles.None, dtFrom) Then

                objMain.objApplication.StatusBar.SetText(
                "Invalid From Date",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If Not DateTime.TryParseExact(ToDate, "yyyyMMdd",
            System.Globalization.CultureInfo.InvariantCulture,
            Globalization.DateTimeStyles.None, dtTo) Then

                objMain.objApplication.StatusBar.SetText(
                "Invalid To Date",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If dtFrom > dtTo Then
                objMain.objApplication.StatusBar.SetText(
                "From Date should not be greater than To Date",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If dtFrom = dtTo Then
                objMain.objApplication.StatusBar.SetText(
                "From Date and To Date should not be same",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Dim currentDocNum As String = ""
            Try
                currentDocNum = CType(oForm.Items.Item("DocNum").Specific, SAPbouiCOM.EditText).Value.Trim()
            Catch
                currentDocNum = ""
            End Try

            rs = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
                   SAPbobsCOM.Recordset)

            Dim sql As String =
            "SELECT TOP 1 ""DocNum"", ""U_FDate"", ""U_TDate"" " &
            "FROM ""@TNX_CTAXCAL"" " &
            "WHERE IFNULL(""Canceled"",'N') = 'N' " &
            "AND ""U_FDate"" IS NOT NULL " &
            "AND ""U_TDate"" IS NOT NULL " &
            "AND TO_DATE('" & dtFrom.ToString("yyyy-MM-dd") & "','YYYY-MM-DD') <= ""U_TDate"" " &
            "AND TO_DATE('" & dtTo.ToString("yyyy-MM-dd") & "','YYYY-MM-DD') >= ""U_FDate"" "

            If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                If currentDocNum <> "" Then
                    sql &= " AND ""DocNum"" <> '" & currentDocNum & "' "
                End If
            End If

            rs.DoQuery(sql)

            If rs.RecordCount > 0 Then

                Dim existDoc As String = rs.Fields.Item("DocNum").Value.ToString()

                objMain.objApplication.StatusBar.SetText(
                "Selected period overlaps with existing Corporate Tax Document No : " & existDoc,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Return False
            End If

            Return True

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
            "Date Validation Error : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            Return False

        Finally
            If rs IsNot Nothing Then
                Try
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
                Catch
                End Try
                rs = Nothing
            End If
        End Try

    End Function
    Private Function IsBranchAvailable() As Boolean
        Try
            Dim rs As SAPbobsCOM.Recordset
            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            rs.DoQuery("SELECT COUNT(*) AS ""Cnt"" FROM OBPL WHERE IFNULL(""Disabled"",'N') = 'N'")

            If CInt(rs.Fields.Item("Cnt").Value) > 0 Then
                Return True
            End If

        Catch ex As Exception
            Return False
        End Try

        Return False
    End Function
    Private Sub SetBranchFieldStatus(ByVal FormUID As String)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim IsBranchDB As Boolean = IsBranchAvailable()

            If IsBranchDB = True Then

                'Branch Label
                objForm.Items.Item("Item_2").Visible = True

                'Branch ComboBox
                objForm.Items.Item("Branch").Visible = True
                objForm.Items.Item("Branch").Enabled = True

            Else

                'Branch Label
                objForm.Items.Item("Item_2").Visible = False

                'Branch ComboBox
                objForm.Items.Item("Branch").Visible = False
                objForm.Items.Item("Branch").Enabled = False

            End If

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
        "Branch field status error : " & ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Private Function GetEditValue(ByVal oForm As SAPbouiCOM.Form,
                              ByVal ItemUID As String) As String
        Try
            Return CType(oForm.Items.Item(ItemUID).Specific,
                     SAPbouiCOM.EditText).Value.Trim()
        Catch
            Return ""
        End Try
    End Function
    Private Sub SetPeriodDates(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oDBs_Head =
        objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL")

            '---------------------------------------
            ' GET YEAR VALUE
            '---------------------------------------
            Dim oYear As SAPbouiCOM.ComboBox =
        CType(objForm.Items.Item("YEAR").Specific,
        SAPbouiCOM.ComboBox)

            If oYear.Selected Is Nothing Then
                objForm.Freeze(False)
                Exit Sub
            End If

            Dim YearVal As String =
        oYear.Selected.Value.Trim()

            If YearVal = "" Then
                objForm.Freeze(False)
                Exit Sub
            End If

            '---------------------------------------
            ' GET START / END PERIOD
            '---------------------------------------
            Dim rs As SAPbobsCOM.Recordset =
        CType(objMain.objCompany.GetBusinessObject(
        SAPbobsCOM.BoObjectTypes.BoRecordset),
        SAPbobsCOM.Recordset)

            Dim Qry As String =
        "SELECT TOP 1 ""U_CMR"", ""U_SO"" " &
        "FROM ""@TNX_CORP"" " &
        "WHERE IFNULL(""U_CMR"",'') <> '' " &
        "AND IFNULL(""U_SO"",'') <> '' " &
        "ORDER BY ""Code"""

            rs.DoQuery(Qry)

            If rs.RecordCount = 0 Then

                objForm.Freeze(False)

                objMain.objApplication.StatusBar.SetText(
            "Corporate Tax setup data not found",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Exit Sub

            End If

            Dim StartMonthName As String =
        rs.Fields.Item("U_CMR").Value.ToString().Trim()

            Dim EndMonthName As String =
        rs.Fields.Item("U_SO").Value.ToString().Trim()

            Dim StartMonth As Integer =
        GetMonthNo(StartMonthName)

            Dim EndMonth As Integer =
        GetMonthNo(EndMonthName)
            objMain.objApplication.StatusBar.SetText(
"Start Month = " & StartMonthName &
" / End Month = " & EndMonthName,
SAPbouiCOM.BoMessageTime.bmt_Short,
SAPbouiCOM.BoStatusBarMessageType.smt_Warning)


            If StartMonth = 0 OrElse EndMonth = 0 Then

                objForm.Freeze(False)

                objMain.objApplication.StatusBar.SetText(
            "Invalid month setup in Corporate Tax Master",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Exit Sub

            End If

            '---------------------------------------
            ' BUILD FROM DATE
            '---------------------------------------
            Dim FromDate As New DateTime(
        CInt(YearVal),
        StartMonth,
        1)

            '---------------------------------------
            ' BUILD TO DATE
            '---------------------------------------
            Dim ToYear As Integer = CInt(YearVal)

            If EndMonth < StartMonth Then
                ToYear += 1
            End If

            Dim ToDate As New DateTime(
        ToYear,
        EndMonth,
        DateTime.DaysInMonth(ToYear, EndMonth))

            '---------------------------------------
            ' SET VALUES TO DATASOURCE
            '---------------------------------------
            oDBs_Head.SetValue(
        "U_FDate",
        0,
        FromDate.ToString("yyyyMMdd"))

            oDBs_Head.SetValue(
        "U_TDate",
        0,
        ToDate.ToString("yyyyMMdd"))

            '---------------------------------------
            ' SET VALUES TO SCREEN
            '---------------------------------------
            CType(objForm.Items.Item("FDate").Specific,
        SAPbouiCOM.EditText).Value =
        FromDate.ToString("yyyyMMdd")

            CType(objForm.Items.Item("TDate").Specific,
        SAPbouiCOM.EditText).Value =
        ToDate.ToString("yyyyMMdd")

            '---------------------------------------
            ' UPDATE MODE
            '---------------------------------------
            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If

            objForm.Freeze(False)

        Catch ex As Exception

            Try
                objForm.Freeze(False)
            Catch
            End Try

            objMain.objApplication.StatusBar.SetText(
        "SetPeriodDates Error : " & ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
    Private Function GetMonthNo(ByVal MonthName As String) As Integer

        Dim m As String =
    MonthName.Trim().ToUpper()

        Select Case m

            Case "1", "01", "JAN", "JANUARY"
                Return 1

            Case "2", "02", "FEB", "FEBRUARY"
                Return 2

            Case "3", "03", "MAR", "MARCH"
                Return 3

            Case "4", "04", "APR", "APRIL"
                Return 4

            Case "5", "05", "MAY"
                Return 5

            Case "6", "06", "JUN", "JUNE"
                Return 6

            Case "7", "07", "JUL", "JULY"
                Return 7

            Case "8", "08", "AUG", "AUGUST"
                Return 8

            Case "9", "09", "SEP", "SEPT", "SEPTEMBER"
                Return 9

            Case "10", "OCT", "OCTOBER"
                Return 10

            Case "11", "NOV", "NOVEMBER"
                Return 11

            Case "12", "DEC", "DECEMBER"
                Return 12

        End Select

        Return 0

    End Function

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)

        Try

            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                    If pVal.BeforeAction = False AndAlso pVal.ItemUID = "YEAR" Then

                        SetPeriodDates(FormUID)

                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL_C2")

                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String = CFLEvent.ChooseFromListUID

                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)

                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects

                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                    If (Not oDT Is Nothing) AndAlso
                   pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE AndAlso
                   pVal.BeforeAction = False Then

                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If

                        If oCFL.UniqueID = "CFL_1" Then
                            oDBs_Head.SetValue("U_Branch", oDBs_Head.Offset, oDT.GetValue("BPLId", 0))
                        End If

                    End If


                Case SAPbouiCOM.BoEventTypes.et_CLICK

                    If pVal.ItemUID = "Item_15" AndAlso pVal.BeforeAction = False Then

                        Try

                            objForm = objMain.objApplication.Forms.Item(FormUID)
                            objForm.Freeze(True)

                            objMatrix = CType(objForm.Items.Item("MXT_1").Specific, SAPbouiCOM.Matrix)
                            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL_C2")

                            Dim selectedRow As Integer = 0

                            For i As Integer = 1 To objMatrix.VisualRowCount
                                If objMatrix.IsRowSelected(i) = True Then
                                    selectedRow = i
                                    Exit For
                                End If
                            Next

                            If selectedRow = 0 Then
                                objMain.objApplication.StatusBar.SetText(
                                "Please select attachment row.",
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Exit Try
                            End If

                            objMatrix.FlushToDataSource()

                            oDBs_Details.RemoveRecord(selectedRow - 1)

                            If oDBs_Details.Size = 0 Then

                                oDBs_Details.InsertRecord(0)
                                oDBs_Details.SetValue("LineId", 0, "1")
                                oDBs_Details.SetValue("U_TPH", 0, "")
                                oDBs_Details.SetValue("U_FNM", 0, "")
                                oDBs_Details.SetValue("U_ATCD", 0, "")
                                oDBs_Details.SetValue("U_FTR", 0, "")

                            Else

                                For i As Integer = 0 To oDBs_Details.Size - 1
                                    oDBs_Details.SetValue("LineId", i, (i + 1).ToString())
                                Next

                            End If

                            objMatrix.LoadFromDataSource()
                            objMatrix.AutoResizeColumns()

                            objForm.Items.Item("Item_15").Enabled = False
                            SetButtonStatus(FormUID)

                        Catch ex As Exception

                            objMain.objApplication.StatusBar.SetText(
                            "Delete Row Error : " & ex.Message,
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        Finally

                            If objForm IsNot Nothing Then
                                objForm.Freeze(False)
                            End If

                        End Try

                    End If
                    If pVal.ItemUID = "COTURL" Then

                        Try

                            Dim rs As SAPbobsCOM.Recordset =
                            objMain.objCompany.GetBusinessObject(
                            SAPbobsCOM.BoObjectTypes.BoRecordset)

                            rs.DoQuery("SELECT TOP 1 ""U_CTAX"" FROM ""@TNX_LKMTR""")

                            If rs.RecordCount > 0 Then

                                Dim COTURL As String =
                                rs.Fields.Item("U_CTAX").Value.ToString().Trim()

                                If COTURL <> "" Then

                                    'Open browser
                                    Process.Start(COTURL)

                                Else

                                    objMain.objApplication.StatusBar.SetText(
                                    "Corporate Tax URL is empty in Link Master",
                                    SAPbouiCOM.BoMessageTime.bmt_Short,
                                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                                End If

                            Else

                                objMain.objApplication.StatusBar.SetText(
                                "Link Master data not found",
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                            End If

                        Catch ex As Exception

                            objMain.objApplication.StatusBar.SetText(
                            "URL Open Error : " & ex.Message,
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        End Try

                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    '====================================================
                    ' BEFORE ADD BUTTON LOGIC
                    '====================================================
                    If pVal.ItemUID = "1" AndAlso pVal.BeforeAction = False AndAlso pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Me.SetDefault(objForm.UniqueID)
                        SetButtonStatus(objForm.UniqueID)

                    End If
                    If pVal.ItemUID = "1" AndAlso pVal.BeforeAction = False Then
                        SetButtonStatus(FormUID)
                    End If
                    If pVal.ItemUID = "1" AndAlso
                   pVal.BeforeAction = True AndAlso
                   pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                        Try
                            If Not ValidateCorporateTaxAdding(objForm) Then

                                BubbleEvent = False
                                Exit Sub

                            End If


                            Try
                                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL")
                            Catch ex As Exception
                                oDBs_Head = Nothing
                            End Try

                            If oDBs_Head Is Nothing Then Exit Try

                            Dim currentAppId As String = ""

                            Try
                                currentAppId = oDBs_Head.GetValue("U_AIPD", oDBs_Head.Offset).Trim()
                            Catch
                                currentAppId = ""
                            End Try

                            If String.IsNullOrEmpty(currentAppId) Then

                                AutoDocentryNumber(objForm.UniqueID)

                                Try
                                    currentAppId = oDBs_Head.GetValue("U_AIPD", oDBs_Head.Offset).Trim()
                                Catch
                                    currentAppId = ""
                                End Try

                            End If

                            Dim DocType As String = "CPA"
                            Dim TableName As String = "@TNX_CTAXCAL"
                            Dim AppIDField As String = "U_AIPD"
                            Dim AppStatField As String = "U_DST"

                            Dim rsapp As SAPbobsCOM.Recordset = Nothing
                            Dim approvalExists As Boolean = False

                            Try

                                rsapp = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                Dim q As String =
                                "SELECT T0.""Code"" As ""TemplateID"",T1.""U_Name"" As ""Originator"", " &
                                "S1.""Code"" As ""Stage"",S2.""U_UKey"" As ""Authorizers"", " &
                                "S1.""U_NAP"" As ""NoOfAuth"",S1.""U_NRJ"" As ""NoOfRej"", " &
                                "T0.""CreateDate"", " &
                                "(Case When T3.""LineId""='1' Then 'S' Else 'O' End) As ""Status"" " &
                                "FROM ""@SBO_APPHDR"" T0 " &
                                "INNER JOIN ""@SBO_APPREQ"" T1 ON T1.""Code""=T0.""Code"" " &
                                "INNER JOIN ""@SBO_APPDOC"" T2 ON T2.""Code"" = T0.""Code"" " &
                                "INNER JOIN ""@SBO_APPAUT"" T3 ON T3.""Code"" = T0.""Code"" " &
                                "INNER JOIN ""@SBO_AST"" S1 ON T3.""U_M3_1""=S1.""Code"" " &
                                "INNER JOIN ""@SBO_AST_C0"" S2 ON S1.""Code""=S2.""Code"" " &
                                "WHERE ""U_Active"" = 'Y' " &
                                "AND T1.""U_Name"" = '" & objMain.objCompany.UserName & "' " &
                                "AND T2.""U_" & DocType & """ = 'Y' " &
                                "AND IFNULL(S2.""U_UKey"",'')<>'' " &
                                "ORDER BY T3.""LineId"""

                                rsapp.DoQuery(q)

                                If rsapp.RecordCount > 0 Then
                                    approvalExists = True
                                End If

                            Catch ex As Exception

                                objMain.objApplication.StatusBar.SetText(
                                "Approval check error: " & ex.Message,
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                            Finally

                                If rsapp IsNot Nothing Then
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsapp)
                                    rsapp = Nothing
                                End If

                            End Try

                            If approvalExists = True Then

                                'Me.ApprovalTrigger(
                                'DocType,
                                'currentAppId,
                                'TableName,
                                'AppIDField,
                                'AppStatField,
                                '"Corporate Report")

                            Else

                                Try
                                    oDBs_Head.SetValue("U_DST", oDBs_Head.Offset, "A")
                                Catch ex As Exception
                                End Try

                                objMain.objApplication.StatusBar.SetText(
                                "Document will be auto-approved.",
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                            End If

                        Catch ex As Exception

                            objMain.objApplication.StatusBar.SetText(
                            "ItemEvent Before Add Error : " & ex.Message,
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        End Try

                        Exit Sub

                    End If

                    If pVal.BeforeAction = True Then Exit Sub

                    '====================================================
                    ' PROFIT & LOSS BUTTON
                    '====================================================
                    If pVal.ItemUID = "Item_14" Then

                        Dim FromDate As String = objForm.Items.Item("FDate").Specific.Value
                        Dim ToDate As String = objForm.Items.Item("TDate").Specific.Value

                        If FromDate = "" OrElse ToDate = "" Then

                            objMain.objApplication.StatusBar.SetText(
                            "Please enter From Date and To Date",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                            Exit Sub

                        End If

                        Try

                            objMain.objApplication.ActivateMenuItem("9731")

                            Dim oPLForm As SAPbouiCOM.Form
                            oPLForm = objMain.objApplication.Forms.ActiveForm

                            CType(oPLForm.Items.Item("3").Specific, SAPbouiCOM.EditText).Value = FromDate
                            CType(oPLForm.Items.Item("6").Specific, SAPbouiCOM.EditText).Value = ToDate
                            oPLForm.Items.Item("1").Click(BoCellClickType.ct_Regular)
                        Catch ex As Exception

                            objMain.objApplication.StatusBar.SetText(
                            "Profit and Loss Open Error : " & ex.Message,
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        End Try
                    ElseIf pVal.ItemUID = "Item_11" Then   'Load Button

                        objForm = objMain.objApplication.Forms.Item(FormUID)

                        Dim FromDate As String = GetEditValue(objForm, "FDate")
                        Dim ToDate As String = GetEditValue(objForm, "TDate")

                        If FromDate = "" OrElse ToDate = "" Then
                            objMain.objApplication.StatusBar.SetText(
            "Please enter From Date and To Date",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Sub
                        End If

                        If Not ValidateCorporateTaxDates(objForm) Then
                            Exit Sub
                        End If

                        LoadTaxCalculation(FormUID)


                    ElseIf pVal.ItemUID = "Item_7" Then

                        PostJournalEntry(FormUID)

                    ElseIf pVal.ItemUID = "Item_4" Then

                        Try

                            Dim AppID As String = objForm.Items.Item("AIPD").Specific.Value

                            Me.ApprovalTrigger(
                            "CPA",
                            AppID,
                            "@TNX_CTAXCAL",
                            "U_AIPD",
                            "U_DST",
                            "Corporate Tax")
                            SetButtonStatus(FormUID)

                        Catch ex As Exception

                            objMain.objApplication.StatusBar.SetText(
                            "Approval Error : " & ex.Message,
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        End Try

                    End If


                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                    If pVal.ItemUID = "MXT_1" AndAlso
                   pVal.ColUID = "TPH" AndAlso
                   pVal.BeforeAction = False Then

                        objForm = objMain.objApplication.Forms.Item(FormUID)

                        Dim objMatrix As SAPbouiCOM.Matrix
                        objMatrix = objForm.Items.Item("MXT_1").Specific

                        If pVal.Row > 0 AndAlso pVal.Row <= objMatrix.VisualRowCount Then

                            Dim fullPath As String
                            fullPath = objMatrix.Columns.Item("TPH").Cells.Item(pVal.Row).Specific.Value

                            If Not String.IsNullOrEmpty(fullPath) AndAlso fullPath.Contains("\") Then

                                Dim indexLoc As Integer = fullPath.LastIndexOf("\")
                                Dim filename As String = fullPath.Substring(indexLoc + 1)

                                objMatrix.Columns.Item("FNM").Cells.Item(pVal.Row).Specific.Value = filename
                                objMatrix.Columns.Item("ATCD").Cells.Item(pVal.Row).Specific.Value = DateTime.Now.ToString("yyyyMMdd")

                                objForm.Items.Item("Item_15").Enabled = True
                                SetButtonStatus(FormUID)

                            End If

                        End If

                    End If

            End Select

        Catch ex As Exception

            Try
                If objForm IsNot Nothing Then objForm.Freeze(False)
            Catch
            End Try

            objMain.objApplication.StatusBar.SetText(
            ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Sub LoadTaxCalculation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL")
            Dim FDate As String = objForm.Items.Item("FDate").Specific.Value
            Dim TDate As String = objForm.Items.Item("TDate").Specific.Value
            Dim oCombo As SAPbouiCOM.ComboBox
            oCombo = objForm.Items.Item("Branch").Specific

            Dim Branch As String = ""

            If oCombo.Selected IsNot Nothing Then
                Branch = oCombo.Selected.Value
            End If
            If FDate = "" Then
                objMain.objApplication.StatusBar.SetText("Enter From Date")
                Exit Sub
            End If
            If TDate = "" Then
                objMain.objApplication.StatusBar.SetText("Enter To Date")
                Exit Sub
            End If
            'If Branch = "" Then
            '    objMain.objApplication.StatusBar.SetText("Select Branch")
            '    Exit Sub
            'End If
            Dim Profit As Double = GetNetProfit(FDate, TDate, Branch)
            Dim YearVal As String =
        CType(objForm.Items.Item("YEAR").Specific,
            SAPbouiCOM.ComboBox).Value.Trim()

            Dim TaxPer As Double =
            GetTaxPercent(Profit, YearVal)
            'Dim TaxPer As Double = GetTaxPercent(Profit)
            Dim TaxVal As Double = (Profit * TaxPer) / 100
            oDBs_Head.SetValue("U_PPeriod", 0, Profit.ToString("0.00"))
            oDBs_Head.SetValue("U_CTax", 0, TaxPer.ToString("0.00"))
            Dim rsCurr As SAPbobsCOM.Recordset

            rsCurr =
objMain.objCompany.GetBusinessObject(
SAPbobsCOM.BoObjectTypes.BoRecordset)

            rsCurr.DoQuery(
"SELECT T1.""CurrName"" " &
"FROM OADM T0 " &
"INNER JOIN OCRN T1 " &
"ON T0.""SysCurrncy"" = T1.""CurrCode""")

            If rsCurr.RecordCount > 0 Then

                Dim SysCurrency As String =
    rsCurr.Fields.Item("CurrName").Value.ToString().Trim()

                oDBs_Head.SetValue(
    "U_CURR",
    0,
    SysCurrency)

                CType(objForm.Items.Item("CURR").Specific,
    SAPbouiCOM.EditText).Value =
    SysCurrency

            End If
            oDBs_Head.SetValue("U_CTaxVal", 0, TaxVal.ToString("0.00"))
            objForm.Update()
            objForm.Items.Item("Item_14").Enabled = True
            objForm.Items.Item("Item_11").Enabled = False
            SetButtonStatus(FormUID)
            ' objMain.objApplication.StatusBar.SetText("Data Loaded Successfully")
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
    "Data Loaded Successfully",
    SAPbouiCOM.BoMessageTime.bmt_Short,
    SAPbouiCOM.BoStatusBarMessageType.smt_Success
)
        End Try
    End Sub

    Function GetNetProfit(ByVal FDate As String, ByVal TDate As String, ByVal Branch As String) As Double

        Dim Profit As Double = 0

        Try
            Dim rs As SAPbobsCOM.Recordset
            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim Qry As String = ""

            If IsBranchAvailable() = True Then
                Qry = "CALL ""TNX_PL_By_Branch"" ('" & FDate & "','" & TDate & "')"
            Else
                Qry = "CALL ""TNX_PL_Without_Branch"" ('" & FDate & "','" & TDate & "')"
            End If

            rs.DoQuery(Qry)

            While rs.EoF = False
                Profit += CDbl(rs.Fields.Item("Total").Value)
                rs.MoveNext()
            End While

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
        "Get Net Profit Error : " & ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

        Return Profit

    End Function
    Function GetTaxPercent(ByVal Profit As Double,
                       ByVal YearVal As String) As Double

        Dim TaxPer As Double = 0

        Try

            Dim rs As SAPbobsCOM.Recordset

            rs = objMain.objCompany.GetBusinessObject(
             SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim Qry As String =
        "SELECT TOP 1 ""U_TaxPrc"" " &
        "FROM ""@TNX_CTAXCNF"" " &
        "WHERE " & Profit & " BETWEEN " &
        "IFNULL(""U_MnProfit"",0) " &
        "AND IFNULL(""U_MxProfit"",0) " &
        "AND ""U_FINA"" = '" & YearVal & "'"

            rs.DoQuery(Qry)

            If rs.RecordCount > 0 Then

                TaxPer =
            CDbl(rs.Fields.Item("U_TaxPrc").Value)

            Else

                objMain.objApplication.StatusBar.SetText(
            "Corporate Tax Slab not found",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            End If

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
        ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

        Return TaxPer

    End Function

    'Function GetTaxPercent(ByVal Profit As Double) As Double

    '    Dim TaxPer As Double = 0

    '    Try

    '        Dim rs As SAPbobsCOM.Recordset
    '        rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        Dim Qry As String

    '        Qry = "SELECT TOP 1 ""U_TaxPrc"" FROM ""@TNX_CTAXCNF"" " &
    '              "WHERE " & Profit & " BETWEEN ""U_MnProfit"" AND ""U_MxProfit"""

    '        rs.DoQuery(Qry)

    '        If rs.RecordCount > 0 Then
    '            TaxPer = CDbl(rs.Fields.Item("U_TaxPrc").Value)
    '        End If

    '    Catch ex As Exception

    '    End Try

    '    Return TaxPer

    'End Function
    Private Sub LoadBranches(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim oCombo As SAPbouiCOM.ComboBox
            oCombo = objForm.Items.Item("Branch").Specific

            While oCombo.ValidValues.Count > 0
                oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
            End While

            Dim rs As SAPbobsCOM.Recordset
            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            rs.DoQuery(
        "SELECT ""BPLId"", ""BPLName"" " &
        "FROM OBPL " &
        "WHERE IFNULL(""Disabled"",'N') = 'N' " &
        "ORDER BY ""BPLId""")

            While rs.EoF = False

                oCombo.ValidValues.Add(
            rs.Fields.Item("BPLId").Value.ToString(),
            rs.Fields.Item("BPLName").Value.ToString())

                rs.MoveNext()

            End While

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
        "Branch Load Error : " & ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
    Private Function IsUserAssignedToBranch(ByVal BranchID As Integer) As Boolean

        Try

            Dim rs As SAPbobsCOM.Recordset
            rs = objMain.objCompany.GetBusinessObject(
        SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim Qry As String =
        "SELECT COUNT(*) AS ""Cnt"" " &
        "FROM USR6 " &
        "WHERE ""UserCode"" = '" & objMain.objCompany.UserName & "' " &
        "AND ""BPLId"" = " & BranchID

            rs.DoQuery(Qry)

            If CInt(rs.Fields.Item("Cnt").Value) > 0 Then
                Return True
            End If

        Catch ex As Exception
        End Try

        Return False

    End Function
    Sub PostJournalEntry(ByVal FormUID As String)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL")

            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE _
        OrElse objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then

                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    objMain.objApplication.StatusBar.SetText(
                "Please complete mandatory fields.",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If

            End If

            Dim BranchID As Integer = 0
            Dim IsBranchDB As Boolean = IsBranchAvailable()

            If IsBranchDB = True Then

                Dim oCombo As SAPbouiCOM.ComboBox
                oCombo = objForm.Items.Item("Branch").Specific

                Dim branchStr As String = ""

                If oCombo.Selected IsNot Nothing Then
                    branchStr = oCombo.Selected.Value
                End If

                If branchStr = "" OrElse
               Not Integer.TryParse(branchStr, BranchID) OrElse
               BranchID <= 0 Then

                    objMain.objApplication.StatusBar.SetText(
                "Please select valid Branch before posting JE.",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If

                If IsUserAssignedToBranch(BranchID) = False Then
                    objMain.objApplication.StatusBar.SetText(
                "You are not assigned to selected branch. Please contact administrator.",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If

            End If

            Dim TaxAmt As Double = 0
            Dim taxStr As String = oDBs_Head.GetValue("U_CTaxVal", 0).Trim()

            If Not Double.TryParse(taxStr, TaxAmt) Then
                TaxAmt = 0
            End If

            If TaxAmt <= 0 Then
                objMain.objApplication.StatusBar.SetText(
            "Tax Amount Missing",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            Dim ExpAcc As String = ""
            Dim LiabAcc As String = ""

            Dim ProfitAmt As Double = 0
            Dim profitStr As String = oDBs_Head.GetValue("U_PPeriod", 0).Trim()

            If Not Double.TryParse(profitStr, ProfitAmt) Then
                ProfitAmt = 0
            End If

            If ProfitAmt <= 0 Then
                objMain.objApplication.StatusBar.SetText(
            "Profit Period Missing. Please click Load before posting JE.",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            GetAccounts(ProfitAmt, ExpAcc, LiabAcc)

            If ExpAcc = "" OrElse LiabAcc = "" Then
                objMain.objApplication.StatusBar.SetText(
            "Tax Configuration Missing for Profit : " & ProfitAmt.ToString("0.00"),
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            Dim oJE As SAPbobsCOM.JournalEntries
            oJE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

            Dim PostDate As Date = Today

            Try
                Dim postDateStr As String = oDBs_Head.GetValue("U_JEPOSTD", 0).Trim()

                If postDateStr <> "" Then
                    PostDate = DateTime.ParseExact(
                postDateStr,
                "yyyyMMdd",
                System.Globalization.CultureInfo.InvariantCulture)
                End If
            Catch
                PostDate = Today
            End Try

            oJE.ReferenceDate = PostDate
            oJE.TaxDate = PostDate
            oJE.DueDate = PostDate
            oJE.Memo = "Corporate Tax Entry"

            oJE.Lines.AccountCode = ExpAcc

            If IsBranchDB = True Then
                oJE.Lines.BPLID = BranchID
            End If

            oJE.Lines.Debit = TaxAmt
            oJE.Lines.Add()

            oJE.Lines.AccountCode = LiabAcc

            If IsBranchDB = True Then
                oJE.Lines.BPLID = BranchID
            End If

            oJE.Lines.Credit = TaxAmt

            If oJE.Add() <> 0 Then
                objMain.objApplication.StatusBar.SetText(
            objMain.objCompany.GetLastErrorDescription,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            Dim JENo As String = objMain.objCompany.GetNewObjectKey().Trim()

            JENo = JENo.Replace(",", "").Trim()

            If JENo.Contains(".") Then
                JENo = JENo.Split("."c)(0)
            End If

            oDBs_Head.SetValue("U_JENo", 0, JENo)
            oDBs_Head.SetValue("U_JEPOSTD", 0, Format(PostDate, "yyyyMMdd"))
            oDBs_Head.SetValue("U_Status", 0, "Posted")

            'Show JE No in UI as plain number
            Try
                CType(objForm.Items.Item("JENO").Specific,
            SAPbouiCOM.EditText).Value = JENo
            Catch
            End Try

            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            SetButtonStatus(FormUID)

            objMain.objApplication.StatusBar.SetText(
        "JE Posted Successfully. JE No : " & JENo,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
        "Post JE Error : " & ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        Finally

            Try
                objForm.Freeze(False)
            Catch
            End Try

        End Try

    End Sub
    'Sub PostJournalEntry(ByVal FormUID As String)

    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        objForm.Freeze(True)

    '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL")

    '        'Save form first if Add / Update mode
    '        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE _
    '    OrElse objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then

    '            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

    '            If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
    '                objMain.objApplication.StatusBar.SetText(
    '            "Please complete mandatory fields.",
    '            SAPbouiCOM.BoMessageTime.bmt_Short,
    '            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                Exit Sub
    '            End If

    '        End If

    '        '---------------------------------------
    '        ' BRANCH LOGIC
    '        '---------------------------------------
    '        Dim BranchID As Integer = 0
    '        Dim IsBranchDB As Boolean = IsBranchAvailable()

    '        If IsBranchDB = True Then

    '            Dim oCombo As SAPbouiCOM.ComboBox
    '            oCombo = objForm.Items.Item("Branch").Specific

    '            Dim branchStr As String = ""

    '            If oCombo.Selected IsNot Nothing Then
    '                branchStr = oCombo.Selected.Value
    '            End If

    '            If branchStr = "" OrElse
    '           Not Integer.TryParse(branchStr, BranchID) OrElse
    '           BranchID <= 0 Then

    '                objMain.objApplication.StatusBar.SetText(
    '            "Please select valid Branch before posting JE.",
    '            SAPbouiCOM.BoMessageTime.bmt_Short,
    '            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '                Exit Sub
    '            End If

    '            If IsUserAssignedToBranch(BranchID) = False Then

    '                objMain.objApplication.StatusBar.SetText(
    '            "You are not assigned to selected branch. Please contact administrator.",
    '            SAPbouiCOM.BoMessageTime.bmt_Short,
    '            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '                Exit Sub
    '            End If

    '        End If

    '        '---------------------------------------
    '        ' TAX AMOUNT
    '        '---------------------------------------
    '        Dim TaxAmt As Double = 0
    '        Dim taxStr As String = oDBs_Head.GetValue("U_CTaxVal", 0).Trim()

    '        If Not Double.TryParse(taxStr, TaxAmt) Then
    '            TaxAmt = 0
    '        End If

    '        If TaxAmt <= 0 Then
    '            objMain.objApplication.StatusBar.SetText(
    '        "Tax Amount Missing",
    '        SAPbouiCOM.BoMessageTime.bmt_Short,
    '        SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Exit Sub
    '        End If

    '        '---------------------------------------
    '        ' GET ACCOUNTS
    '        '---------------------------------------
    '        Dim ExpAcc As String = ""
    '        Dim LiabAcc As String = ""

    '        GetAccounts(TaxAmt, ExpAcc, LiabAcc)

    '        If ExpAcc = "" OrElse LiabAcc = "" Then
    '            objMain.objApplication.StatusBar.SetText(
    '        "Tax Configuration Missing",
    '        SAPbouiCOM.BoMessageTime.bmt_Short,
    '        SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Exit Sub
    '        End If

    '        '---------------------------------------
    '        ' CREATE JE
    '        '---------------------------------------
    '        Dim oJE As SAPbobsCOM.JournalEntries
    '        oJE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

    '        Dim PostDate As Date = Today

    '        Try
    '            Dim postDateStr As String = oDBs_Head.GetValue("U_JEPOSTD", 0).Trim()

    '            If postDateStr <> "" Then
    '                PostDate = DateTime.ParseExact(
    '            postDateStr,
    '            "yyyyMMdd",
    '            System.Globalization.CultureInfo.InvariantCulture)
    '            End If
    '        Catch
    '            PostDate = Today
    '        End Try

    '        oJE.ReferenceDate = PostDate
    '        oJE.TaxDate = PostDate
    '        oJE.DueDate = PostDate
    '        oJE.Memo = "Corporate Tax Entry"

    '        'Debit Line
    '        oJE.Lines.AccountCode = ExpAcc

    '        If IsBranchDB = True Then
    '            oJE.Lines.BPLID = BranchID
    '        End If

    '        oJE.Lines.Debit = TaxAmt
    '        oJE.Lines.Add()

    '        'Credit Line
    '        oJE.Lines.AccountCode = LiabAcc

    '        If IsBranchDB = True Then
    '            oJE.Lines.BPLID = BranchID
    '        End If

    '        oJE.Lines.Credit = TaxAmt

    '        If oJE.Add() <> 0 Then

    '            objMain.objApplication.StatusBar.SetText(
    '        objMain.objCompany.GetLastErrorDescription,
    '        SAPbouiCOM.BoMessageTime.bmt_Short,
    '        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '            Exit Sub
    '        End If

    '        '---------------------------------------
    '        ' UPDATE JE NUMBER
    '        '---------------------------------------
    '        Dim JENo As String = objMain.objCompany.GetNewObjectKey()

    '        oDBs_Head.SetValue("U_JENo", 0, JENo)
    '        oDBs_Head.SetValue("U_JEPOSTD", 0, Format(PostDate, "yyyyMMdd"))
    '        oDBs_Head.SetValue("U_Status", 0, "Posted")

    '        objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
    '        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

    '        objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
    '        SetButtonStatus(FormUID)

    '        objMain.objApplication.StatusBar.SetText(
    '    "JE Posted Successfully",
    '    SAPbouiCOM.BoMessageTime.bmt_Short,
    '    SAPbouiCOM.BoStatusBarMessageType.smt_Success)

    '    Catch ex As Exception

    '        objMain.objApplication.StatusBar.SetText(
    '    "Post JE Error : " & ex.Message,
    '    SAPbouiCOM.BoMessageTime.bmt_Short,
    '    SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '    Finally

    '        Try
    '            objForm.Freeze(False)
    '        Catch
    '        End Try

    '    End Try

    'End Sub
    'Sub PostJournalEntry(ByVal FormUID As String)

    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        objForm.Freeze(True)

    '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CTAXCAL")

    '        'First save form if Add/Update mode
    '        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE _
    '    OrElse objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then

    '            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

    '            If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
    '                objMain.objApplication.StatusBar.SetText("Please complete mandatory fields.")
    '                Exit Sub
    '            End If

    '        End If

    '        'Branch validation
    '        Dim BranchID As Integer = 0
    '        Dim oCombo As SAPbouiCOM.ComboBox
    '        oCombo = objForm.Items.Item("Branch").Specific

    '        Dim branchStr As String = ""

    '        If oCombo.Selected IsNot Nothing Then
    '            branchStr = oCombo.Selected.Value
    '        End If

    '        If branchStr = "" OrElse
    '       Not Integer.TryParse(branchStr, BranchID) OrElse
    '       BranchID <= 0 Then

    '            objMain.objApplication.StatusBar.SetText(
    '        "Please select valid Branch before posting JE.",
    '        SAPbouiCOM.BoMessageTime.bmt_Short,
    '        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '            Exit Sub
    '        End If
    '        If IsUserAssignedToBranch(BranchID) = False Then

    '            objMain.objApplication.StatusBar.SetText(
    '"You are not assigned to selected branch. Please contact administrator.",
    'SAPbouiCOM.BoMessageTime.bmt_Short,
    'SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '            Exit Sub

    '        End If

    '        'Tax amount
    '        Dim TaxAmt As Double = 0
    '        Dim taxStr As String = oDBs_Head.GetValue("U_CTaxVal", 0).Trim()

    '        If Not Double.TryParse(taxStr, TaxAmt) Then
    '            TaxAmt = 0
    '        End If

    '        If TaxAmt <= 0 Then
    '            objMain.objApplication.StatusBar.SetText("Tax Amount Missing")
    '            Exit Sub
    '        End If

    '        'Accounts
    '        Dim ExpAcc As String = ""
    '        Dim LiabAcc As String = ""

    '        GetAccounts(TaxAmt, ExpAcc, LiabAcc)

    '        If ExpAcc = "" OrElse LiabAcc = "" Then
    '            objMain.objApplication.StatusBar.SetText("Tax Configuration Missing")
    '            Exit Sub
    '        End If

    '        'Create JE
    '        Dim oJE As SAPbobsCOM.JournalEntries
    '        oJE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

    '        Dim PostDate As Date = Today

    '        oJE.ReferenceDate = PostDate
    '        oJE.TaxDate = PostDate
    '        oJE.DueDate = PostDate
    '        oJE.Memo = "Corporate Tax Entry"

    '        'Debit Line
    '        oJE.Lines.AccountCode = ExpAcc
    '        oJE.Lines.BPLID = BranchID
    '        oJE.Lines.Debit = TaxAmt
    '        oJE.Lines.Add()

    '        'Credit Line
    '        oJE.Lines.AccountCode = LiabAcc
    '        oJE.Lines.BPLID = BranchID
    '        oJE.Lines.Credit = TaxAmt

    '        If oJE.Add() <> 0 Then
    '            objMain.objApplication.StatusBar.SetText(
    '        objMain.objCompany.GetLastErrorDescription,
    '        SAPbouiCOM.BoMessageTime.bmt_Short,
    '        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '            Exit Sub
    '        End If

    '        'Get JE No
    '        Dim JENo As String = objMain.objCompany.GetNewObjectKey()

    '        'Update current document
    '        oDBs_Head.SetValue("U_JENo", 0, JENo)
    '        oDBs_Head.SetValue("U_JEPOSTD", 0, Format(PostDate, "yyyyMMdd"))
    '        oDBs_Head.SetValue("U_Status", 0, "Posted")

    '        objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
    '        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

    '        objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
    '        SetButtonStatus(FormUID)

    '        objMain.objApplication.StatusBar.SetText(
    '    "JE Posted Successfully",
    '    SAPbouiCOM.BoMessageTime.bmt_Short,
    '    SAPbouiCOM.BoStatusBarMessageType.smt_Success)

    '    Catch ex As Exception

    '        objMain.objApplication.StatusBar.SetText(
    '    "Post JE Error : " & ex.Message,
    '    SAPbouiCOM.BoMessageTime.bmt_Short,
    '    SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '    Finally

    '        Try
    '            objForm.Freeze(False)
    '        Catch
    '        End Try

    '    End Try

    'End Sub



    Private Sub DisableDateFields(ByVal FormUID As String)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Items.Item("FDate").Enabled = False       'From Date
            objForm.Items.Item("TDate").Enabled = False       'To Date
            objForm.Items.Item("JEPOSTD").Enabled = False     'JE Posting Date

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("DisableDateFields Error : " & ex.Message)
        End Try

    End Sub
    Sub GetAccounts(ByVal Profit As Double, ByRef ExpAcc As String, ByRef LiabAcc As String)

        Try

            Dim rs As SAPbobsCOM.Recordset
            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim Qry As String = "SELECT TOP 1 ""U_LAccount"", ""U_EAccount"" " &
      "FROM ""@TNX_CTAXCNF"" " &
      "WHERE " & Replace(Profit.ToString(), ",", "") &
      " BETWEEN IFNULL(""U_MnProfit"",0) " &
      " AND IFNULL(""U_MxProfit"",0)"

            rs.DoQuery(Qry)

            If rs.RecordCount > 0 Then

                LiabAcc = rs.Fields.Item("U_LAccount").Value.ToString().Trim()
                ExpAcc = rs.Fields.Item("U_EAccount").Value.ToString().Trim()

            Else

                ExpAcc = ""
                LiabAcc = ""

                objMain.objApplication.StatusBar.SetText(
                "No Account Configuration Found for Profit : " & Profit,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            End If

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
            "GetAccounts Error : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
    Private Sub SetAutoManagedFields(ByVal FormUID As String)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            'Load Button
            objForm.Items.Item("Item_11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Item_11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            'Profit & Loss Button
            objForm.Items.Item("Item_14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Item_14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'Submit Button
            objForm.Items.Item("Item_4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            'Post JE Button
            objForm.Items.Item("Item_7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            'Application ID
            objForm.Items.Item("AIPD").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("AIPD").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            'Document Status
            objForm.Items.Item("DST").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DST").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("SetAutoManagedFields Error : " & ex.Message)
        End Try

    End Sub
    'Sub GetAccounts(ByVal Profit As Double, ByRef ExpAcc As String, ByRef LiabAcc As String)

    '    Try

    '        Dim rs As SAPbobsCOM.Recordset
    '        rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        Dim Qry As String

    '        Qry = "SELECT TOP 1 ""U_LAccount"", ""U_EAccount"" FROM ""@TNX_CTAXCNF"" " &
    '              "WHERE " & Profit & " BETWEEN ""U_MnProfit"" AND ""U_MxProfit"""

    '        rs.DoQuery(Qry)

    '        If rs.RecordCount > 0 Then
    '            LiabAcc = rs.Fields.Item("U_LAccount").Value.ToString
    '            ExpAcc = rs.Fields.Item("U_EAccount").Value.ToString
    '        End If

    '    Catch ex As Exception

    '    End Try

    'End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.FormTypeEx = "CTAXCAL" Then

                If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD _
                    And BusinessObjectInfo.BeforeAction = False _
                    And BusinessObjectInfo.ActionSuccess = True Then

                    objForm = objMain.objApplication.Forms.Item(BusinessObjectInfo.FormUID)

                    Dim NewDocEntry As String = objMain.objCompany.GetNewObjectKey()

                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    objForm.Items.Item("DocNum").Specific.Value = NewDocEntry
                    objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)


                End If

                If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD _
                    And BusinessObjectInfo.BeforeAction = False _
                    And BusinessObjectInfo.ActionSuccess = True Then

                    objForm = objMain.objApplication.Forms.Item(BusinessObjectInfo.FormUID)

                    'After document loaded, check JE No and disable Post JE button
                    SetButtonStatus(BusinessObjectInfo.FormUID)
                    DisableDateFields(BusinessObjectInfo.FormUID)

                End If
                If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE _
AndAlso BusinessObjectInfo.BeforeAction = False _
AndAlso BusinessObjectInfo.ActionSuccess = True Then

                    objForm = objMain.objApplication.Forms.Item(BusinessObjectInfo.FormUID)

                    SetButtonStatus(BusinessObjectInfo.FormUID)

                End If
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    'Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
    '    Try
    '        If BusinessObjectInfo.FormTypeEx = "CTAXCAL" Then
    '            If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
    '                objForm = objMain.objApplication.Forms.Item(BusinessObjectInfo.FormUID)
    '                Dim NewDocEntry As String = objMain.objCompany.GetNewObjectKey()
    '                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
    '                objForm.Items.Item("DocNum").Specific.Value = NewDocEntry
    '                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
    '            End If
    '        End If
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)

    '    End Try
    'End Sub

End Class


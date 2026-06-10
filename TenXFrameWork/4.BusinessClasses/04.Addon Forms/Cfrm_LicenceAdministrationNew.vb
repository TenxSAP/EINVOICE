
Imports System.Text
Imports SBO_ADDONBASE
Imports SBO_ENUMS
Imports SBO_TABLES
Imports System.Threading
Imports System.IO
Imports System.Security.Cryptography
Imports System.Windows.Forms

Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Public Class Cfrm_LicenceAdministrationNew

#Region "       Declaration             "
    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDS1, oDS2 As SAPbouiCOM.DBDataSource
    Dim objMatrix, objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim str, str1 As String
    Public rs, RsNum As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False
    Dim objutilities As Utilities
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("LicenceNew.xml", "License", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("License", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            objutilities = New Utilities
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_LICENCE")
            oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_LICENCE_C0")
            ' oDS2 = objForm.DataSources.DBDataSources.Item("@SBO_ADDON")
            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("1293", True)
            SetNewLine()
            'SetNewLine1()
            ' Me.SetDefault(objForm.UniqueID)
            'oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "SBO_LICENSE_UDO"))

            'Dim oMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific
            'Dim oItem As SAPbouiCOM.Item = objForm.Items.Item("Mtx")
            ''Set fixed height and width
            'oItem.Height = 200 ' Fixed height in pixels
            'oItem.Width = 300 ' Fixed width in pixels



            'Dim oMatrix1 As SAPbouiCOM.Matrix = objForm.Items.Item("Item_5").Specific
            'Dim oItem1 As SAPbouiCOM.Item = objForm.Items.Item("Item_5")
            ''Set fixed height and width
            'oItem1.Height = 200 ' Fixed height in pixels
            'oItem1.Width = 300 ' Fixed width in pixels







            'objForm.DataBrowser.BrowseBy = "License"
            ' objForm.Items.Item("License").Visible = False
            objMatrix = objForm.Items.Item("Mtx").Specific
            objMatrix.AutoResizeColumns()
            Dim Query As String = "SELECT * FROM ""@TNX_LICENCE"""
            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rs.DoQuery(Query)
            objForm.Items.Item("DB").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DB").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("License").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("License").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("Company").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Company").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("SDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("SDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("EDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("EDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("HKey").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("HKey").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("Total").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Total").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            ' objForm.DataBrowser.BrowseBy = "DocNum"
            objMatrix = objForm.Items.Item("Mtx").Specific
            If rs.RecordCount = 0 Then
                oDBs_Head.SetValue("DocNum", 0, GetNextDocNum(objForm, "TNX_LICENCE_UDO"))

                Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'Dim Str As String = "SELECT T0.""Code"", T0.""Name"", T0.""U_MAC"" FROM ""@SBO_DEVICE""  T0"
                'For i As Integer = 1 To oRs.RecordCount
                '    objMatrix.AddRow()
                '    oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                '    oDS1.SetValue("U_Code", oDS1.Offset, oRs.Fields.Item("Code").Value)
                '    oDS1.SetValue("U_USERC", oDS1.Offset, oRs.Fields.Item("Name").Value)
                '    ' oDS1.SetValue("U_MAC", oDS1.Offset, oRs.Fields.Item("U_MAC").Value)
                '    ' oDS1.SetValue("U_Sts", oDS1.Offset, "")
                '    objMatrix.SetLineData(objMatrix.VisualRowCount)
                '    oRs.MoveNext()
                'Next
                'objMatrix.AutoResizeColumns()
                'Me.SetNewLine()
            Else

                Dim Str1 As String = "SELECT IFNULL(Max(""DocNum""),'0')   FROM ""@TNX_LICENCE"" "
                Dim oRsDocExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsDocExist.DoQuery(Str1)

                If oRsDocExist.RecordCount > 0 Then
                    If oRsDocExist.Fields.Item(0).Value <> "0" Then
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        Dim docnum As String = oRsDocExist.Fields.Item(0).Value
                        objForm.Items.Item("DocNum").Specific.Value = Convert.ToInt32(oRsDocExist.Fields.Item(0).Value.ToString())
                        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                End If
                'objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                'objForm.Items.Item("DocNum").Specific.Value = "1"
                'objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)   '
                objMain.objApplication.ActivateMenuItem("1290")
                '   objForm.EnableMenu("1281", False)
                '    objForm.EnableMenu("1282", False)
            End If
            'objMatrix.Columns.Item("DKey").Visible = False
            objMatrix.AutoResizeColumns()
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String)
        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            objutilities = New Utilities
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_LICENCE")
            oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_LICENCE_C0")
            ' oDS2 = objForm.DataSources.DBDataSources.Item("@SBO_ADDON")
            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("1293", True)
            objMatrix = objForm.Items.Item("Mtx").Specific

            oDS1.Clear()

            objMatrix.FlushToDataSource()


            SetNewLine()
            oDBs_Head.SetValue("DocNum", 0, GetNextDocNum(objForm, "TNX_LICENCE_UDO"))
            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
    Public ReadOnly Property FormMatrix(ByVal pst_MatrixUID As String) As SAPbouiCOM.Matrix
        Get
            Return CType(objForm.Items.Item(pst_MatrixUID).Specific, SAPbouiCOM.Matrix)
        End Get
    End Property

    Public Function GetNextDocNum(ByRef Objform As SAPbouiCOM.Form, ByVal UDOName As String, Optional ByVal SeriesName As String = "Primary") As Integer
        Dim Str As String
        Dim oRs As SAPbobsCOM.Recordset
        Dim DocNum As Integer
        'oRs = Objform.SBO_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Str = "Select ""Series"" From NNM1 Where ""ObjectCode"" = '" & UDOName & "' and ""SeriesName"" = '" & SeriesName & "'"
        Try
            oRs.DoQuery(Str)
            oRs.MoveFirst()
            If oRs.RecordCount > 0 Then
                DocNum = Objform.BusinessObject.GetNextSerialNumber(oRs.Fields.Item(0).Value, UDOName)
            End If
            If DocNum = 0 Then DocNum = 1
            Return DocNum
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Sub SetNewLine()
        Try


            objMatrix = objForm.Items.Item("Mtx").Specific
            objMatrix.AddRow()
            oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_LICENCE_C0")

            oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
            oDS1.SetValue("U_Code", oDS1.Offset, "")
            oDS1.SetValue("U_Name", oDS1.Offset, "")
            oDS1.SetValue("U_Sts", oDS1.Offset, "")

            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Sub
    'Sub SetNewLine1()
    '    objMatrix1 = objForm.Items.Item("Item_5").Specific
    '    objMatrix1.AddRow()
    '    oDS2.SetValue("LineId", oDS2.Offset, objMatrix1.VisualRowCount)
    '    oDS2.SetValue("U_ADCODE", oDS2.Offset, "")
    '    oDS2.SetValue("U_ADNAM", oDS2.Offset, "")
    '    oDS2.SetValue("U_USER", oDS2.Offset, "")
    '    oDS2.SetValue("U_AVBLE", oDS2.Offset, "")

    '    objMatrix1.SetLineData(objMatrix1.VisualRowCount)
    '    objMatrix1.AutoResizeColumns()
    'End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "License" And pVal.BeforeAction = False Then
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
            End If
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    'objMatrix = objForm.Items.Item("Mtx").Specific
                    'If pVal.ItemUID = "Mtx" And pVal.ColUID = "DName" Then
                    '    If pVal.Row = objMatrix.VisualRowCount Then
                    '        Me.SetNewLine()
                    '    End If
                    'End If
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "Mtx" And pVal.ColUID = "Sts" And pVal.BeforeAction = False Then
                        objMatrix = objForm.Items.Item("Mtx").Specific
                        Dim objCheckx As SAPbouiCOM.CheckBox = objMatrix.Columns.Item("Sts").Cells.Item(pVal.Row).Specific
                        Dim TotalLicenses As Integer = CInt(objForm.Items.Item("Total").Specific.Value)
                        If objCheckx.Checked = True Then

                            Dim LicenseCount As Integer = 0
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                Dim objCheck As SAPbouiCOM.CheckBox = objMatrix.Columns.Item("Sts").Cells.Item(i).Specific
                                If objCheck.Checked = True Then
                                    LicenseCount = LicenseCount + 1
                                End If
                            Next
                            If TotalLicenses < LicenseCount Then
                                objMain.objApplication.StatusBar.SetText("You cannot  select more then " & TotalLicenses & " total licenses", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                ''Exit Sub
                                objCheckx.Checked = False
                            End If
                        End If

                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    'If pVal.ItemUID = "Mtx" And pVal.ColUID = "Sts" And pVal.BeforeAction = False Then
                    '    objMatrix = objForm.Items.Item("Mtx").Specific
                    '    Dim objCheck As SAPbouiCOM.CheckBox = objMatrix.Columns.Item("Sts").Cells.Item(pVal.Row).Specific
                    '    ''objMatrix.Columns.Item("DKey").Visible = False
                    '    If objCheck.Checked = True Then
                    '        '---------------------------------------
                    '        Dim TotalLicenses As Integer = CInt(objForm.Items.Item("Total").Specific.Value)
                    '        Dim LicenseCount As Integer = 0

                    '        For i As Integer = 1 To objMatrix.VisualRowCount
                    '            Dim objCheck1 As SAPbouiCOM.CheckBox = objMatrix.Columns.Item("Sts").Cells.Item(i).Specific
                    '            If objCheck1.Checked = True Then
                    '                LicenseCount = LicenseCount + 1
                    '            End If
                    '        Next
                    '        If TotalLicenses < LicenseCount Then
                    '            objMain.objApplication.StatusBar.SetText("You cannot  select more then total licenses", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '            oDS1.SetValue("U_Sts", oDS1.Offset, "N")
                    '            objMatrix.LoadFromDataSource()
                    '            BubbleEvent = False
                    '            Exit Sub
                    '        End If
                    '        '-----------------------------------------
                    '    End If
                    '    objMatrix.AutoResizeColumns()
                    '    ''objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                    'End If
                    'Comment
                    If pVal.ItemUID = "Import" And pVal.BeforeAction = False Then
                        Dim newThread As Thread = New Thread(New ThreadStart(AddressOf ShowFolderBrowser))
                        newThread.SetApartmentState(ApartmentState.STA)
                        newThread.Start()
                    End If
                    If pVal.ItemUID = "Refresh" And pVal.BeforeAction = False Then
                        objMatrix = objForm.Items.Item("Mtx").Specific
                        objMatrix.Clear()

                        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim Str As String = "Select ""USER_CODE"",""U_NAME"" as ""UserName"" from OUSR "
                        oRs.DoQuery(Str)
                        For i As Integer = 1 To oRs.RecordCount
                            objMatrix.AddRow()
                            oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                            oDS1.SetValue("U_Code", oDS1.Offset, oRs.Fields.Item("USER_CODE").Value)
                            oDS1.SetValue("U_Name", oDS1.Offset, oRs.Fields.Item("UserName").Value)
                            oDS1.SetValue("U_Sts", oDS1.Offset, "")
                            objMatrix.SetLineData(objMatrix.VisualRowCount)
                            'End If
                            oRs.MoveNext()
                        Next
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        'If pVal.ItemUID = "Mtx" And pVal.ColUID = "Sts" And pVal.BeforeAction = True Then
                        objMatrix = objForm.Items.Item("Mtx").Specific
                        ' Dim objCheck As SAPbouiCOM.CheckBox = objMatrix.Columns.Item("Sts").Cells.Item(pVal.Row).Specific
                        ''objMatrix.Columns.Item("DKey").Visible = False
                        ' If objCheck.Checked = True Then
                        '---------------------------------------
                        Dim TotalLicenses As Integer = CInt(objForm.Items.Item("Total").Specific.Value)
                        Dim LicenseCount As Integer = 0

                        For i As Integer = 1 To objMatrix.VisualRowCount
                            Dim objCheck1 As SAPbouiCOM.CheckBox = objMatrix.Columns.Item("Sts").Cells.Item(i).Specific
                            If objCheck1.Checked = True Then
                                LicenseCount = LicenseCount + 1
                            End If
                        Next
                        If TotalLicenses < LicenseCount Then
                            objMain.objApplication.StatusBar.SetText("The Number of Selected Licenses : " & LicenseCount & " is greater than the available Licences: " & TotalLicenses & " ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            'oDS1.SetValue("U_Sts", oDS1.Offset, "N")
                            'objMatrix.LoadFromDataSource()
                            BubbleEvent = False
                            Exit Sub
                        End If
                        '-----------------------------------------
                        'End If
                        'End If
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        'objMain.objApplication.ActivateMenuItem("1290")
                        ' ElseIf pVal.ItemUID = "1" And pVal.BeforeAction = True Then
                        objMatrix = objForm.Items.Item("Mtx").Specific
                        Dim TotalLicenses As Integer = CInt(objForm.Items.Item("Total").Specific.Value)
                        ''Dim TotalLicenses As String = objForm.Items.Item("Total").Specific.Value.ToString()

                        Dim LicenseCount As Integer = 0
                        For i As Integer = 1 To objMatrix.VisualRowCount
                            Dim objCheck As SAPbouiCOM.CheckBox = objMatrix.Columns.Item("Sts").Cells.Item(i).Specific
                            If objCheck.Checked = True Then
                                LicenseCount = LicenseCount + 1
                            End If
                        Next
                        If TotalLicenses < LicenseCount Then
                            objMain.objApplication.StatusBar.SetText("You cannot  select more then total licenses", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            ''Exit Sub
                        End If

                        Dim License, Database, Company, HKey, TotalLicense As String
                        Dim StartDate, ExpDate As Date
                        License = objForm.Items.Item("License").Specific.Value.ToString().Trim()
                        Database = objForm.Items.Item("DB").Specific.Value.ToString().Trim()
                        Company = objForm.Items.Item("Company").Specific.Value.ToString().Trim()
                        HKey = objForm.Items.Item("HKey").Specific.Value.ToString().Trim()
                        TotalLicense = objForm.Items.Item("Total").Specific.Value.ToString().Trim()
                        Dim SDate As String = objForm.Items.Item("SDate").Specific.Value.ToString().Trim()
                        Dim EDate As String = objForm.Items.Item("EDate").Specific.Value.ToString().Trim()
                        Dim Addon As String = objForm.Items.Item("Addon").Specific.Value.ToString().Trim()

                        SDate = SDate.Insert(4, "/")
                        SDate = SDate.Insert(7, "/")
                        EDate = EDate.Insert(4, "/")
                        EDate = EDate.Insert(7, "/")
                        StartDate = CDate(SDate)
                        ExpDate = CDate(EDate)
                        Dim uniKey As String = "10X_SAP_B!_"
                        'Dim strKey As String = StartDate.ToString("dd-MMM-yyyy") & "|" & ExpDate.ToString("dd-MMM-yyyy") & "|" & Company & "|" & TotalLicenses & "|" & Database & "|" & HKey
                        'Dim strKey As String = StartDate.ToString("dd-MMM-yyyy") & "|" + ExpDate.ToString("dd-MMM-yyyy") & "|" + Company & "|" & TotalLicenses & "|" + Database & "|" + HKey & "|" + Addon 

                        Dim EInvoice As String = "N"
                        Dim CorporateTax As String = "N"
                        Dim VatReport As String = "N"

                        Dim chkEIVC As SAPbouiCOM.CheckBox = objForm.Items.Item("EIVC").Specific
                        Dim chkCTTX As SAPbouiCOM.CheckBox = objForm.Items.Item("CTTX").Specific
                        Dim chkVRPT As SAPbouiCOM.CheckBox = objForm.Items.Item("VRPT").Specific

                        If chkEIVC.Checked = True Then
                            EInvoice = "Y"
                        End If

                        If chkCTTX.Checked = True Then
                            CorporateTax = "Y"
                        End If

                        If chkVRPT.Checked = True Then
                            VatReport = "Y"
                        End If

                        Dim strKey As String = StartDate.ToString("dd-MMM-yyyy") & "|" +
ExpDate.ToString("dd-MMM-yyyy") & "|" +
Company & "|" &
TotalLicenses & "|" +
Database & "|" +
HKey & "|" +
Addon & "|" +
EInvoice & "|" +
CorporateTax & "|" +
VatReport

                        Dim TestLicense As String = GetSHA1HashData(strKey + GetSHA1HashData(uniKey))
                        Dim GenLicense As String = GetSHA1HashData(strKey + GetSHA1HashData(uniKey))
                        If License <> GetSHA1HashData(strKey + GetSHA1HashData(uniKey)) Then
                            objMain.objApplication.StatusBar.SetText("Invalid license file", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If

            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub ShowFolderBrowser()
        Dim srcfilenamewithfullpath As String
        Dim MyTest1 As New OpenFileDialog
        With MyTest1.DefaultExt = ".ae"
            MyTest1.Filter = "License Files (*.ae)|*.ae|All files (*.*)|*.*"
            MyTest1.FilterIndex = 2
        End With
        If MyTest1.ShowDialog() = DialogResult.OK Then
            Try
                srcfilenamewithfullpath = MyTest1.FileName
            Catch ex As IO.IOException
                objMain.objApplication.MessageBox(ex.Message)
                Exit Sub
            End Try
            System.Windows.Forms.Application.ExitThread()
            Me.CSVUpload(srcfilenamewithfullpath)

        End If
        Exit Sub
    End Sub
    Sub CSVUpload(ByVal path As String)
        Try
            Dim FileSaveWithPath As String = path
            Dim Fulltext As String = ""
            Dim License, Database, Company, HKey, TotalLicenses, Addon As String

            Dim EInvoice As String = "N"
            Dim CorporateTax As String = "N"
            Dim VatReport As String = "N"

            Dim StartDate, ExpDate As Date
            Dim rs, rsCheck As SAPbobsCOM.Recordset
            objMain.objApplication.StatusBar.SetText("License file importing, please wait.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Using sr As New IO.StreamReader(FileSaveWithPath)
                While Not sr.EndOfStream
                    Fulltext = sr.ReadLine.ToString()
                    Dim Str As String() = Fulltext.Split(":")
                    If Str(0) = "[SartDate]" Then
                        StartDate = Str(1)
                    ElseIf Str(0) = "[EndDate]" Then
                        ExpDate = Str(1)
                    ElseIf Str(0) = "[ClientName]" Then
                        Company = Str(1)
                    ElseIf Str(0) = "[NoOfDevices]" Then
                        TotalLicenses = Str(1)
                    ElseIf Str(0) = "[HARDWAREKEY]" Then
                        HKey = Str(1)
                    ElseIf Str(0) = "[DBNAME]" Then
                        Database = Str(1)
                    ElseIf Str(0) = "[CERTIFICATEKEY]" Then
                        License = Str(1)
                    ElseIf Str(0) = "[ADDONNAME]" Then
                        Addon = Str(1)

                    ElseIf Str(0) = "[EINVOICE]" Then
                        EInvoice = Str(1)

                    ElseIf Str(0) = "[CORPORATETAX]" Then
                        CorporateTax = Str(1)

                    ElseIf Str(0) = "[VATREPORT]" Then
                        VatReport = Str(1)
                    End If

                End While
                System.Windows.Forms.Application.ExitThread()
                objForm.EnableMenu("1282", True)
                '---------------------------------------------------
                Dim HKEY1 As String = String.Empty
                Dim oForm As SAPbouiCOM.Form
                objMain.objApplication.ActivateMenuItem("257")
                oForm = objMain.objApplication.Forms.Item(objMain.objApplication.Forms.ActiveForm.UniqueID)
                HKEY1 = oForm.Items.Item("79").Specific.value
                oForm.Close()

                Dim COMANY As String = objMain.objCompany.CompanyDB
                objMain.objApplication.StatusBar.SetText("Verifying DB Name ", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                If COMANY <> Database Then
                    objMain.objApplication.StatusBar.SetText("Invalid license file. Company doesn't match: " & Company & " not equal to " & COMANY, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
                objMain.objApplication.StatusBar.SetText("Verifying Hardware Key ", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                If HKey <> HKEY1 Then
                    objMain.objApplication.StatusBar.SetText("Invalid license file. HardWare Key doesn't match: " & HKey & " not equal to " & HKEY1, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If

                '---------------------------------------------------
                Dim uniKey As String = "10X_SAP_B!_"
                'Dim strKey As String = StartDate.ToString("dd-MMM-yyyy") & "|" + ExpDate.ToString("dd-MMM-yyyy") & "|" + Company & "|" + TotalLicenses & "|" + Database & "|" + HKey
                Dim strKey As String = StartDate.ToString("dd-MMM-yyyy") & "|" + ExpDate.ToString("dd-MMM-yyyy") & "|" + Company & "|" & TotalLicenses & "|" + Database & "|" + HKey & "|" + Addon & "|" +
EInvoice & "|" +
CorporateTax & "|" +
VatReport
                Dim TestLicense As String = GetSHA1HashData(strKey + GetSHA1HashData(uniKey))


                If License <> TestLicense Then
                    objMain.objApplication.StatusBar.SetText("Invalid license file", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If

                '''----------------------------------------------------
                oDBs_Head.SetValue("U_License", 0, License)
                oDBs_Head.SetValue("U_DB", 0, Database)
                oDBs_Head.SetValue("U_SDate", 0, StartDate.ToString("yyyyMMdd"))
                oDBs_Head.SetValue("U_EDate", 0, ExpDate.ToString("yyyyMMdd"))
                oDBs_Head.SetValue("U_Total", 0, TotalLicenses)
                oDBs_Head.SetValue("U_Company", 0, Company)
                oDBs_Head.SetValue("U_HKey", 0, HKey)
                oDBs_Head.SetValue("U_Addon", 0, Addon)

                oDBs_Head.SetValue("U_EIVC", 0, EInvoice)
                oDBs_Head.SetValue("U_CTTX", 0, CorporateTax)
                oDBs_Head.SetValue("U_VRPT", 0, VatReport)

                Dim chkEIVC As SAPbouiCOM.CheckBox
                Dim chkCTTX As SAPbouiCOM.CheckBox
                Dim chkVRPT As SAPbouiCOM.CheckBox

                chkEIVC = objForm.Items.Item("EIVC").Specific
                chkCTTX = objForm.Items.Item("CTTX").Specific
                chkVRPT = objForm.Items.Item("VRPT").Specific

                chkEIVC.Checked = (EInvoice = "Y")
                chkCTTX.Checked = (CorporateTax = "Y")
                chkVRPT.Checked = (VatReport = "Y")

                '------------------------------------------

                Dim Query As String = "SELECT * FROM ""@TNX_LICENCE""T0 Inner Join  ""@TNX_LICENCE_C0"" T1 on T0.""DocEntry""=T1.""DocEntry""where ""U_Addon""='" & Addon & "'"
                rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rs.DoQuery(Query)
                If rs.RecordCount = 0 Then
                    objMatrix = objForm.Items.Item("Mtx").Specific
                    Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim StrQuery As String = "Select ""USER_CODE"",""U_NAME"" as ""UserName"" from OUSR "
                    oRs.DoQuery(StrQuery)
                    objMatrix.Clear()
                    For i As Integer = 1 To oRs.RecordCount
                        objMatrix.AddRow()
                        oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                        oDS1.SetValue("U_Name", oDS1.Offset, oRs.Fields.Item("UserName").Value)
                        oDS1.SetValue("U_Code", oDS1.Offset, oRs.Fields.Item("USER_CODE").Value)
                        oDS1.SetValue("U_Sts", oDS1.Offset, "")
                        objMatrix.SetLineData(objMatrix.VisualRowCount)
                        oRs.MoveNext()
                    Next
                Else
                    For j As Integer = 1 To rs.RecordCount

                        ' objMatrix.AddRow()
                        oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                        oDS1.SetValue("U_Name", oDS1.Offset, rs.Fields.Item("U_Name").Value)
                        oDS1.SetValue("U_Code", oDS1.Offset, rs.Fields.Item("U_Code").Value)
                        oDS1.SetValue("U_Sts", oDS1.Offset, rs.Fields.Item("U_Sts").Value)
                        objMatrix.SetLineData(objMatrix.VisualRowCount)
                        rs.MoveNext()
                    Next
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If

                '------------------------------------------

                objMain.objApplication.StatusBar.SetText("License file imported successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End Using
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Sub
    Private Function GetSHA1HashData(ByVal data As String) As String

        Dim sha1 As SHA1 = SHA1.Create()
        Dim hashData As Byte() = sha1.ComputeHash(Encoding.[Default].GetBytes(data))
        Dim returnValue As System.Text.StringBuilder = New StringBuilder()
        For i As Integer = 0 To hashData.Length - 1
            returnValue.Append(hashData(i).ToString())
        Next
        Return returnValue.ToString()
    End Function

#Region " For Resize"
    'Private Sub SBO_Application_FormResizeAfter(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.SBOItemEventArg)
    '    Try
    '        'objForm = objMain.objApplication.Forms.GetForm("LIECEN", objMain..ActiveForm.TypeCount)
    '        Dim screenWidth As Integer = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width
    '        Dim screenHeight As Integer = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height

    '        If screenWidth < 1024 Or screenHeight < 768 Then
    '            ' Adjust control size or hide non-essential controls for smaller screens

    '            Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)
    '            Dim oItem As SAPbouiCOM.Item = oForm.Items.Item("Mtx")
    '            'Dim oItem1 As SAPbouiCOM.Item = oForm.Items.Item("Mtx")

    '            ' Reapply fixed dimensions
    '            oItem.Height = 200
    '            oItem.Width = 250

    '            'Dim oForm1 As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)
    '            'Dim oItem As SAPbouiCOM.Item = oForm.Items.Item("Mtx")
    '            Dim oItem1 As SAPbouiCOM.Item = oForm.Items.Item("Item_5")

    '            ' Reapply fixed dimensions
    '            oItem1.Height = 200
    '            oItem1.Width = 250
    '        End If


    '    Catch ex As Exception

    '    End Try
    'End Sub
#End Region
End Class




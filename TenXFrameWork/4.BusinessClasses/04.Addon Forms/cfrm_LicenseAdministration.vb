Imports System.Text
Imports SBO_ADDONBASE
Imports SBO_ENUMS
Imports SBO_TABLES
Imports System.Threading
Imports System.IO
Imports System.Security.Cryptography
Imports System.Windows.Forms

Imports System.Runtime.InteropServices

Public Class cfrm_LicenseAdministration

#Region "       Declaration             "
    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDS1 As SAPbouiCOM.DBDataSource
    Dim objMatrix, objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim str, str1 As String
    Public rs, RsNum As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False
    Dim objutilities As Utilities
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("LicenseAdministration.xml", "License", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("License", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            objutilities = New Utilities
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_LICENSE_C")
            oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_LICENSE_C1")
            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("1293", True)
            SetNewLine()
            ' Me.SetDefault(objForm.UniqueID)
            'oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "SBO_LICENSE_UDO"))


            'objForm.DataBrowser.BrowseBy = "License"
            ' objForm.Items.Item("License").Visible = False
            objMatrix = objForm.Items.Item("Mtx").Specific
            objMatrix.AutoResizeColumns()
            Dim Query As String = "SELECT * FROM ""@SBO_LICENSE_C"""
            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rs.DoQuery(Query)
            objForm.Items.Item("DB").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DB").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("License").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("License").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

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

            objMatrix = objForm.Items.Item("Mtx").Specific
            If rs.RecordCount = 0 Then
                oDBs_Head.SetValue("DocNum", 0, GetNextDocNum(objForm, "SBO_LICENSE_UDO"))
                Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim Str As String = "SELECT T0.""Code"", T0.""Name"", T0.""U_MAC"" FROM ""@SBO_DEVICE""  T0"
                For i As Integer = 1 To oRs.RecordCount
                    objMatrix.AddRow()
                    oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                    oDS1.SetValue("U_Code", oDS1.Offset, oRs.Fields.Item("Code").Value)
                    oDS1.SetValue("U_DName", oDS1.Offset, oRs.Fields.Item("Name").Value)
                    oDS1.SetValue("U_MAC", oDS1.Offset, oRs.Fields.Item("U_MAC").Value)
                    oDS1.SetValue("U_Sts", oDS1.Offset, "")
                    objMatrix.SetLineData(objMatrix.VisualRowCount)
                    oRs.MoveNext()
                Next
                objMatrix.AutoResizeColumns()
                'Me.SetNewLine()
            Else


                'Dim Str1 As String = "SELECT IFNULL(Max(""DocNum""),'0')   FROM ""@SBO_LICENSE_C"" "
                'Dim oRsDocExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRsDocExist.DoQuery(Str1)

                'If oRsDocExist.RecordCount > 0 Then
                '    If oRsDocExist.Fields.Item(0).Value <> "0" Then
                '        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                '        Dim docnum As String = oRsDocExist.Fields.Item(0).Value
                '        objForm.Items.Item("DocNum").Specific.Value = Convert.ToInt32(oRsDocExist.Fields.Item(0).Value.ToString())
                '        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                '    End If
                'End If




                'Dim Str1 As String = "SELECT IFNULL(Max(""DocNum""),'0')   FROM ""@SBO_LICENSE_C"" "
                'Dim oRsDocExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRsDocExist.DoQuery(Str1)

                'If oRsDocExist.RecordCount > 0 Then
                '    If oRsDocExist.Fields.Item(0).Value <> "0" Then
                '        objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                '        Dim docnum As String = oRsDocExist.Fields.Item(0).Value
                '        objForm.Items.Item("DocNum").Specific.Value = Convert.ToInt32(oRsDocExist.Fields.Item(0).Value.ToString())
                '        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                '    End If
                'End If





                'objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                'objForm.Items.Item("DocNum").Specific.Value = "1"
                'objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)   '
                objMain.objApplication.ActivateMenuItem("1290")
                objForm.EnableMenu("1281", False)
                objForm.EnableMenu("1282", False)
            End If
            objMatrix.Columns.Item("DKey").Visible = False
            objMatrix.AutoResizeColumns()
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
        objMatrix = objForm.Items.Item("Mtx").Specific
        objMatrix.AddRow()
        oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
        oDS1.SetValue("U_Code", oDS1.Offset, "")
        oDS1.SetValue("U_DName", oDS1.Offset, "")
        oDS1.SetValue("U_Sts", oDS1.Offset, "")
        oDS1.SetValue("U_MAC", oDS1.Offset, "")
        oDS1.SetValue("U_DKey", oDS1.Offset, "")
        objMatrix.SetLineData(objMatrix.VisualRowCount)
        objMatrix.AutoResizeColumns()
    End Sub



    '
    ' Summary:
    '     Computes the System.Security.Cryptography.SHA1 hash for the input data.
    'ComVisible(True)>
    'Public MustInherit Class SHA1
    '    Inherits HashAlgorithm

    'End Class






    'Sub SetDefault(ByVal FormUID As String)
    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_LICENSE_C")
    '        oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_LICENSE_C1")
    '        objForm.Freeze(True)
    '        'Me.SetNewLine()
    '        Dim GetDocNum As String = "Select Max(""DocNum"")+1 From ""@SBO_LICENSE_C"" "
    '        Dim rsDocNum As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        rsDocNum.DoQuery(GetDocNum)

    '        oDBs_Head.SetValue("U_DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "SBO_LICENSE_UDO"))
    '        objMatrix.AutoResizeColumns()
    '        objForm.Freeze(False)
    '    Catch ex As Exception
    '        objForm.Freeze(False)
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "License" And pVal.BeforeAction = False Then
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
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
                    If pVal.ItemUID = "Mtx" And pVal.ColUID = "DName" Then
                        If pVal.Row = objMatrix.VisualRowCount Then
                            Me.SetNewLine()
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "Mtx" And pVal.ColUID = "Sts" And pVal.BeforeAction = False Then
                        objMatrix = objForm.Items.Item("Mtx").Specific

                        Dim objCheck As SAPbouiCOM.CheckBox = objMatrix.Columns.Item("Sts").Cells.Item(pVal.Row).Specific

                        objMatrix.Columns.Item("DKey").Visible = False
                        If objCheck.Checked = True Then
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
                                objMain.objApplication.StatusBar.SetText("You cannot  select more then total licenses", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                oDS1.SetValue("U_Sts", oDS1.Offset, "N")
                                objMatrix.LoadFromDataSource()
                                BubbleEvent = False
                                Exit Sub
                            End If
                            '-----------------------------------------
                            Dim uniKey As String = "MICROEXCEL_SAP_B!_"
                            Dim strKey As String = objMatrix.Columns.Item("MAC").Cells.Item(pVal.Row).Specific.Value
                            Dim HKey As String = GetSHA1HashData(strKey + GetSHA1HashData(uniKey))
                            objMatrix.Columns.Item("DKey").Cells.Item(pVal.Row).Specific.Value = HKey

                        Else
                            objMatrix.Columns.Item("DKey").Cells.Item(pVal.Row).Specific.Value = ""
                        End If
                        objMatrix.Columns.Item("DKey").Visible = False
                        objMatrix.AutoResizeColumns()
                    End If
                    'Comment
                    If pVal.ItemUID = "Import" And pVal.BeforeAction = False Then
                        Dim newThread As Thread = New Thread(New ThreadStart(AddressOf ShowFolderBrowser))
                        newThread.SetApartmentState(ApartmentState.STA)
                        newThread.Start()
                    End If
                    If pVal.ItemUID = "Refresh" And pVal.BeforeAction = False Then

                        objMatrix = objForm.Items.Item("Mtx").Specific
                        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim Str As String = "SELECT T0.""Code"", T0.""Name"", T0.""U_MAC"" FROM ""@SBO_DEVICE"" T0 Where T0.""Code"" Not In (SELECT ifnull(T1.""U_Code"",'') FROM ""@SBO_LICENSE_C1""  T1) and T0.""Code"" is not null and T0.""Code""<>''"
                        oRs.DoQuery(Str)
                        For i As Integer = 1 To oRs.RecordCount
                            objMatrix.AddRow()
                            oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                            oDS1.SetValue("U_Code", oDS1.Offset, oRs.Fields.Item("Code").Value)
                            oDS1.SetValue("U_DName", oDS1.Offset, oRs.Fields.Item("Name").Value)
                            oDS1.SetValue("U_MAC", oDS1.Offset, oRs.Fields.Item("U_MAC").Value)
                            oDS1.SetValue("U_Sts", oDS1.Offset, "")
                            oDS1.SetValue("U_DKey", oDS1.Offset, "")
                            objMatrix.SetLineData(objMatrix.VisualRowCount)
                            'End If
                            oRs.MoveNext()
                        Next
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        'objMain.objApplication.ActivateMenuItem("1290")
                    ElseIf pVal.ItemUID = "1" And pVal.BeforeAction = True Then
                        objMatrix = objForm.Items.Item("Mtx").Specific
                        Dim TotalLicenses As Integer = CInt(objForm.Items.Item("Total").Specific.Value)
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
                        End If
                        '-----------------------------
                        Dim License, Database, Company, HKey, TotalLicense As String
                        Dim StartDate, ExpDate As Date
                        License = objForm.Items.Item("License").Specific.Value.ToString().Trim()
                        Database = objForm.Items.Item("DB").Specific.Value.ToString().Trim()
                        Company = objForm.Items.Item("Company").Specific.Value.ToString().Trim()
                        HKey = objForm.Items.Item("HKey").Specific.Value.ToString().Trim()
                        TotalLicense = objForm.Items.Item("Total").Specific.Value.ToString().Trim()
                        Dim SDate As String = objForm.Items.Item("SDate").Specific.Value.ToString().Trim()
                        Dim EDate As String = objForm.Items.Item("EDate").Specific.Value.ToString().Trim()
                        SDate = SDate.Insert(4, "/")
                        SDate = SDate.Insert(7, "/")
                        EDate = EDate.Insert(4, "/")
                        EDate = EDate.Insert(7, "/")
                        StartDate = CDate(SDate)
                        ExpDate = CDate(EDate)
                        Dim uniKey As String = "MICROEXCEL_SAP_B!_"


                        'Dim strKey As String = StartDate.ToString("dd-MMM-yyyy") & "|" & ExpDate.ToString("dd-MMM-yyyy") & "|" & Company & "|" & TotalLicenses & "|" & Database & "|" & HKey
                        'Comment
                        Dim strKey As String = StartDate.ToString("dd-MMM-yyyy") & "|" + ExpDate.ToString("dd-MMM-yyyy") & "|" + Company & "|" & TotalLicenses & "|" + Database & "|" + HKey
                        Dim TestLicense As String = GetSHA1HashData(strKey + GetSHA1HashData(uniKey))
                        Dim GenLicense As String = GetSHA1HashData(strKey + GetSHA1HashData(uniKey))

                        If License <> GetSHA1HashData(strKey + GetSHA1HashData(uniKey)) Then
                            objMain.objApplication.StatusBar.SetText("Invalid license file", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        'Dim TestLicense As String = (strKey + uniKey)

                        'If License <> (strKey + uniKey) Then
                        '    objMain.objApplication.StatusBar.SetText("Invalid license file", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If


                        ''---------------------------------
                    End If

            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    'Public MustInherit Class SHA1
    '    Inherits HashAlgorithm

    'End Class
    'Public Shared Function Create() As SHA1

    'End Function
    'Public Shared Function Create(hashName As String) As SHA1

    'End Function
    '

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
            Dim License, Database, Company, HKey, TotalLicenses As String
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
                    End If

                End While
                System.Windows.Forms.Application.ExitThread()
                objForm.EnableMenu("1282", True)
                '---------------------------------------------------
                Dim uniKey As String = "MICROEXCEL_SAP_B!_"
                'Dim strKey As String = StartDate.ToString("dd-MMM-yyyy") & "|" + ExpDate.ToString("dd-MMM-yyyy") & "|" + Company & "|" + TotalLicenses & "|" + Database & "|" + HKey
                Dim strKey As String = StartDate.ToString("dd-MMM-yyyy") & "|" + ExpDate.ToString("dd-MMM-yyyy") & "|" + Company & "|" & TotalLicenses & "|" + Database & "|" + HKey
                'Comment
                Dim TestLicense As String = GetSHA1HashData(strKey + GetSHA1HashData(uniKey))
                If License <> GetSHA1HashData(strKey + GetSHA1HashData(uniKey)) Then
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
                '------------------------------------------
                objMatrix = objForm.Items.Item("Mtx").Specific
                Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim StrQuery As String = "SELECT T0.""Code"", T0.""Name"", T0.""U_MAC"" FROM ""@SBO_DEVICE""  T0"
                oRs.DoQuery(StrQuery)
                objMatrix.Clear()
                For i As Integer = 1 To oRs.RecordCount
                    objMatrix.AddRow()
                    oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                    oDS1.SetValue("U_Code", oDS1.Offset, oRs.Fields.Item("Code").Value)
                    oDS1.SetValue("U_DName", oDS1.Offset, oRs.Fields.Item("Name").Value)
                    oDS1.SetValue("U_MAC", oDS1.Offset, oRs.Fields.Item("U_MAC").Value)
                    oDS1.SetValue("U_Sts", oDS1.Offset, "")
                    oDS1.SetValue("U_DKey", oDS1.Offset, "")
                    objMatrix.SetLineData(objMatrix.VisualRowCount)
                    oRs.MoveNext()
                Next
                '------------------------------------------
                Dim Query As String = "SELECT * FROM ""@SBO_LICENSE_C"""
                rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rs.DoQuery(Query)
                If rs.RecordCount = 0 Then
                    ' Me.m_SBO_Form.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                Else
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If
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


End Class




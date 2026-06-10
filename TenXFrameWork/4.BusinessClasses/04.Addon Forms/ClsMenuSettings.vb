
Imports System.Security.Cryptography
Imports System.Xml

Imports SAPbobsCOM


Public Class ClsMenuSettings

    'Dim objForm As SAPbouiCOM.Form
    'Dim objMatrix, objMatrix2 As SAPbouiCOM.Matrix
    'Dim oDBs_Head, oDBs_Details, oDBs_Details2 As SAPbouiCOM.DBDataSource
    'Dim oGridUser, oGridUserGroup As SAPbouiCOM.Grid
    'Dim oDT As SAPbouiCOM.Grid

    'Public Sub New()

    'End Sub

    'Public Structure MenuItemData
    '    Public UniqueID As String
    '    Public Name As String
    '    Public FatherUID As String
    '    Public Positionn As Integer
    'End Structure


    'Public Class MenuItem
    '    Public Property FatherUID As String
    '    Public Property UniqueID As String
    '    Public Property Title As String
    '    Public Property Position As Integer
    '    Public Property Level As Integer ' For tree indentation
    'End Class

    'Public Function LoadMenuDataUpdate(xmlFilePath As String)

    '    Try

    '        objForm.Freeze(True)

    '        Dim menuList As New List(Of MenuItem)()
    '        Dim xmlDoc As New Xml.XmlDocument()
    '        xmlDoc.Load("Menu.xml") ' Full path if needed

    '        Dim menuNodes = xmlDoc.SelectNodes("//Menu")

    '        For Each node As Xml.XmlNode In menuNodes
    '            Dim fatherUID As String = node.Attributes("FatherUID")?.Value
    '            Dim uniqueID As String = node.Attributes("UniqueID")?.Value
    '            Dim title As String = node.Attributes("String")?.Value
    '            Dim positionStr As String = node.Attributes("Position")?.Value



    '            Dim position As Integer = 0
    '            If Integer.TryParse(positionStr, position) Then
    '                menuList.Add(New MenuItem With {
    '        .FatherUID = fatherUID,
    '        .UniqueID = uniqueID,
    '        .Title = title,
    '        .Position = position
    '    })
    '            End If
    '        Next


    '        Dim sortedMenus = menuList.
    'OrderBy(Function(m) m.FatherUID).
    'ThenBy(Function(m) m.UniqueID).
    'ToList()

    '        Dim rsExcluded As SAPbobsCOM.Recordset
    '        rsExcluded = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        Dim query As String = "SELECT DISTINCT ""U_MID"" FROM ""@TENXMENUC0"""
    '        rsExcluded.DoQuery(query)


    '        Dim excludedIDs As New List(Of String)
    '        While Not rsExcluded.EoF
    '            excludedIDs.Add(rsExcluded.Fields.Item(0).Value.ToString())
    '            rsExcluded.MoveNext()
    '        End While

    '        System.Runtime.InteropServices.Marshal.ReleaseComObject(rsExcluded)



    '        Dim groupedMenus = From m In menuList Where Not excludedIDs.Contains(m.UniqueID)
    '                           Group m By Key = New With {m.FatherUID, m.Title} Into Group
    '                           Let sortedGroup = Group.OrderBy(Function(x) x.UniqueID).ToList()
    '                           Order By Key.FatherUID
    '                           Select New With {
    '                   .FatherUID = Key.FatherUID,
    '                   .Title = Key.Title,
    '                   .Menus = sortedGroup
    '               }


    '        Dim dt As New DataTable()
    '        dt.Columns.Add("FatherUID")
    '        dt.Columns.Add("UniqueID")
    '        dt.Columns.Add("Title")
    '        dt.Columns.Add("Position", GetType(Integer))

    '        Dim ChangeStat As String = "NO"

    '        For Each group In groupedMenus.Distinct

    '            For Each item In group.Menus

    '                dt.Rows.Add(group.FatherUID, item.UniqueID, item.Title, item.Position)

    '                ChangeStat = "Yes"

    '                objMatrix.AddRow()

    '                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
    '                oDBs_Details.SetValue("U_PID", oDBs_Details.Offset, group.FatherUID)
    '                oDBs_Details.SetValue("U_PMN", oDBs_Details.Offset, group.Title)

    '                oDBs_Details.SetValue("U_MID", oDBs_Details.Offset, item.UniqueID)
    '                oDBs_Details.SetValue("U_MNM", oDBs_Details.Offset, item.Title)

    '                oDBs_Details.SetValue("U_VSBL", oDBs_Details.Offset, "Y")


    '                objMatrix.SetLineData(objMatrix.VisualRowCount)


    '            Next

    '        Next

    '        If ChangeStat.Trim.ToUpper = "YES" Then
    '            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
    '                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
    '                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
    '            End If

    '        End If

    '        objMatrix.AutoResizeColumns()
    '        objMatrix2.AutoResizeColumns()


    '        objForm.Freeze(False)
    '    Catch ex As Exception
    '        objForm.Freeze(False)
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Function

    'Public Function LoadMenuData(xmlFilePath As String)

    '    Try

    '        objForm.Freeze(True)

    '        Dim menuList As New List(Of MenuItem)()
    '        Dim xmlDoc As New Xml.XmlDocument()
    '        xmlDoc.Load("Menu.xml") ' Full path if needed

    '        Dim menuNodes = xmlDoc.SelectNodes("//Menu")

    '        For Each node As Xml.XmlNode In menuNodes
    '            Dim fatherUID As String = node.Attributes("FatherUID")?.Value
    '            Dim uniqueID As String = node.Attributes("UniqueID")?.Value
    '            Dim title As String = node.Attributes("String")?.Value
    '            Dim positionStr As String = node.Attributes("Position")?.Value



    '            Dim position As Integer = 0
    '            If Integer.TryParse(positionStr, position) Then
    '                menuList.Add(New MenuItem With {
    '        .FatherUID = fatherUID,
    '        .UniqueID = uniqueID,
    '        .Title = title,
    '        .Position = position
    '    })
    '            End If
    '        Next


    '        Dim sortedMenus = menuList.
    'OrderBy(Function(m) m.FatherUID).
    'ThenBy(Function(m) m.UniqueID).
    'ToList()


    '        Dim groupedMenus = From m In menuList
    '                           Group m By Key = New With {m.FatherUID, m.Title} Into Group
    '                           Let sortedGroup = Group.OrderBy(Function(x) x.UniqueID).ToList()
    '                           Order By Key.FatherUID
    '                           Select New With {
    '                   .FatherUID = Key.FatherUID,
    '                   .Title = Key.Title,
    '                   .Menus = sortedGroup
    '               }


    '        Dim dt As New DataTable()
    '        dt.Columns.Add("FatherUID")
    '        dt.Columns.Add("UniqueID")
    '        dt.Columns.Add("Title")
    '        dt.Columns.Add("Position", GetType(Integer))

    '        For Each group In groupedMenus.Distinct

    '            For Each item In group.Menus

    '                dt.Rows.Add(group.FatherUID, item.UniqueID, item.Title, item.Position)

    '                objMatrix.AddRow()

    '                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
    '                oDBs_Details.SetValue("U_PID", oDBs_Details.Offset, group.FatherUID)
    '                oDBs_Details.SetValue("U_PMN", oDBs_Details.Offset, group.Title)

    '                oDBs_Details.SetValue("U_MID", oDBs_Details.Offset, item.UniqueID)
    '                oDBs_Details.SetValue("U_MNM", oDBs_Details.Offset, item.Title)

    '                oDBs_Details.SetValue("U_VSBL", oDBs_Details.Offset, "Y")


    '                objMatrix.SetLineData(objMatrix.VisualRowCount)


    '            Next

    '        Next



    '        objMatrix.AutoResizeColumns()
    '        objMatrix2.AutoResizeColumns()

    '        objForm.Freeze(False)
    '    Catch ex As Exception
    '        objForm.Freeze(False)
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Function
    'Public Function LoadMenuData_Old1(xmlFilePath As String)

    '    Try



    '        Dim menuList As New List(Of MenuItem)()
    '        Dim xmlDoc As New Xml.XmlDocument()
    '        xmlDoc.Load("Menu.xml") ' Full path if needed

    '        Dim menuNodes = xmlDoc.SelectNodes("//Menu")

    '        For Each node As Xml.XmlNode In menuNodes
    '            Dim fatherUID As String = node.Attributes("FatherUID")?.Value
    '            Dim uniqueID As String = node.Attributes("UniqueID")?.Value
    '            Dim title As String = node.Attributes("String")?.Value
    '            Dim positionStr As String = node.Attributes("Position")?.Value



    '            Dim position As Integer = 0
    '            If Integer.TryParse(positionStr, position) Then
    '                menuList.Add(New MenuItem With {
    '        .FatherUID = fatherUID,
    '        .UniqueID = uniqueID,
    '        .Title = title,
    '        .Position = position
    '    })
    '            End If
    '        Next


    '        Dim sortedMenus = menuList.
    'OrderBy(Function(m) m.FatherUID).
    'ThenBy(Function(m) m.UniqueID).
    'ToList()


    '        Dim groupedMenus = From m In menuList
    '                           Group m By Key = New With {m.FatherUID, m.Title} Into Group
    '                           Let sortedGroup = Group.OrderBy(Function(x) x.UniqueID).ToList()
    '                           Order By Key.FatherUID
    '                           Select New With {
    '                   .FatherUID = Key.FatherUID,
    '                   .Title = Key.Title,
    '                   .Menus = sortedGroup
    '               }


    '        Dim dt As New DataTable()
    '        dt.Columns.Add("FatherUID")
    '        dt.Columns.Add("UniqueID")
    '        dt.Columns.Add("Title")
    '        dt.Columns.Add("Position", GetType(Integer))

    '        For Each group In groupedMenus.Distinct

    '            For Each item In group.Menus

    '                dt.Rows.Add(group.FatherUID, item.UniqueID, item.Title, item.Position)

    '                objMatrix.AddRow()

    '                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
    '                oDBs_Details.SetValue("U_PID", oDBs_Details.Offset, group.FatherUID)
    '                oDBs_Details.SetValue("U_PMN", oDBs_Details.Offset, group.Title)

    '                oDBs_Details.SetValue("U_MID", oDBs_Details.Offset, item.UniqueID)
    '                oDBs_Details.SetValue("U_MNM", oDBs_Details.Offset, item.Title)

    '                oDBs_Details.SetValue("U_VSBL", oDBs_Details.Offset, "Y")


    '                objMatrix.SetLineData(objMatrix.VisualRowCount)


    '            Next

    '        Next



    '        objMatrix.AutoResizeColumns()
    '        objMatrix2.AutoResizeColumns()

    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Function

    'Public Function LoadMenuData_old(xmlFilePath As String)

    '    Dim menuList As New List(Of MenuItemData)()
    '    Dim xmlDoc As New XmlDocument()
    '    xmlDoc.Load(xmlFilePath)

    '    Dim menuNodes As XmlNodeList = xmlDoc.SelectNodes("//Menu")
    '    For Each node As XmlNode In menuNodes
    '        If node.Attributes IsNot Nothing Then


    '            Dim name As String = node.Attributes("String")?.Value
    '            Dim uid As String = node.Attributes("UniqueID")?.Value
    '            Dim fatherUid As String = node.Attributes("FatherUID")?.Value
    '            Dim Positionn As Integer = node.Attributes("Position")?.Value

    '            If Not String.IsNullOrEmpty(name) AndAlso Not String.IsNullOrEmpty(uid) Then
    '                menuList.Add(New MenuItemData With {
    '                    .Name = name,
    '                    .UniqueID = uid,
    '                    .FatherUID = fatherUid,
    '                    .Positionn = Positionn
    '                })
    '            End If
    '        End If
    '    Next

    '    'Dim groupedByUID = From m In menuList
    '    '                   Group m By UID = m.UniqueID Into Group
    '    '                   Select UID, Items = Group.ToList()

    '    Dim groupedMenus = From m In menuList
    '                       Group m By Key = New With {m.FatherUID, m.UniqueID, m.Name} Into Group
    '                       Let sortedGroup = Group.OrderBy(Function(x) x.Positionn).ToList()
    '                       Select FatherUID = Key.FatherUID, UniqueID = Key.UniqueID, Name = Key.Name, Items = sortedGroup


    '    For Each group In groupedMenus

    '        Dim uid As String = group.FatherUID
    '        Dim unm As String = group.Name

    '        Dim items As List(Of MenuItemData) = group.Items





    '        For Each item In items


    '            objMatrix.AddRow()

    '            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)

    '            'oDBs_Details.SetValue("U_PID", oDBs_Details.Offset, uid)
    '            'oDBs_Details.SetValue("U_PMN", oDBs_Details.Offset, item.Name)
    '            'oDBs_Details.SetValue("U_MID", oDBs_Details.Offset, item.FatherUID)
    '            'oDBs_Details.SetValue("U_MNM", oDBs_Details.Offset, item.UniqueID)
    '            'oDBs_Details.SetValue("U_VSBL", oDBs_Details.Offset, "Y")

    '            oDBs_Details.SetValue("U_PID", oDBs_Details.Offset, uid)
    '            oDBs_Details.SetValue("U_PMN", oDBs_Details.Offset, unm)

    '            oDBs_Details.SetValue("U_MID", oDBs_Details.Offset, item.UniqueID)
    '            oDBs_Details.SetValue("U_MNM", oDBs_Details.Offset, item.Name)

    '            oDBs_Details.SetValue("U_VSBL", oDBs_Details.Offset, "Y")


    '            'If Row = 0 Then
    '            '    objMatrix.CommonSetting.SetRowBackColor(objMatrix.VisualRowCount, ColorTranslator.ToOle(Color.LightGreen))
    '            '    objMatrix.CommonSetting.SetRowFontStyle(objMatrix.VisualRowCount, SAPbouiCOM.BoFontStyle.fs_Bold)
    '            'End If

    '            'Row = Row + 1

    '            objMatrix.SetLineData(objMatrix.VisualRowCount)

    '        Next


    '    Next



    '    Dim dt As New DataTable()
    '    dt.Columns.Add("FatherUID", GetType(String))
    '    dt.Columns.Add("UniqueID", GetType(String))
    '    dt.Columns.Add("Name", GetType(String))
    '    dt.Columns.Add("Items", GetType(String)) ' Store serialized item list

    '    ' Loop through the groupedMenus and add to DataTable
    '    For Each group In groupedMenus
    '        Dim itemsList = group.Items.Select(Function(i) i.Name).ToList()
    '        Dim itemsString As String = String.Join(", ", itemsList)

    '        dt.Rows.Add(group.FatherUID, group.UniqueID, group.Name, itemsString)
    '    Next


    '    Dim sortedView As New DataView(dt)
    '    sortedView.Sort = "FatherUID, UniqueID"
    '    Dim sortedTable As DataTable = sortedView.ToTable()



    '    Return menuList
    'End Function
    'Sub CreateForm()
    '    Try
    '        objMain.objUtilities.LoadForm("MenuSettings.xml", "TENXMENU_Form", ResourceType.Embeded)
    '        objForm = objMain.objApplication.Forms.GetForm("TENXMENU_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
    '        objForm.Freeze(True)


    '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TENXMENU")
    '        oDBs_Details = objForm.DataSources.DBDataSources.Item("@TENXMENUC0")
    '        oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TENXMENUC1")

    '        objMatrix = objForm.Items.Item("Item_2").Specific
    '        objMatrix2 = objForm.Items.Item("Item_13").Specific

    '        oGridUser = objForm.Items.Item("Item_4").Specific
    '        oGridUserGroup = objForm.Items.Item("Item_5").Specific

    '        objForm.Items.Item("Item_8").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

    '        objForm.Items.Item("Item_10").TextStyle = 1
    '        objForm.Items.Item("Item_1").TextStyle = 1
    '        objForm.Items.Item("Item_0").TextStyle = 1


    '        objMain.objApplication.MetadataAutoRefresh = True



    '        Dim CheckDoc As String = "Select ""DocNum"" FROM ""@TENXMENU"" WHERE ""DocNum"" IS NOT NULL"
    '        Dim oRsCheckDoc As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
    '        oRsCheckDoc.DoQuery(CheckDoc)

    '        If oRsCheckDoc.RecordCount > 0 Then

    '            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
    '            objForm.Items.Item("Item_6").Enabled = True
    '            objForm.Items.Item("Item_6").Specific.Value = oRsCheckDoc.Fields.Item(0).Value
    '            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

    '            Me.LoadMenuDataUpdate(IO.Directory.GetParent(Application.ExecutablePath).ToString & "\Menu.xml")

    '            oGridUser.AutoResizeColumns()
    '            oGridUserGroup.AutoResizeColumns()
    '            objMatrix.AutoResizeColumns()
    '            objMatrix2.AutoResizeColumns()


    '            objForm.Freeze(False)

    '            Exit Try

    '        Else

    '            Me.LoadMenuData(IO.Directory.GetParent(Application.ExecutablePath).ToString & "\Menu.xml")

    '            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TENXMENU"))
    '        End If


    '        objForm.Freeze(False)
    '    Catch ex As Exception
    '        objForm.Freeze(False)
    '        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    End Try
    'End Sub

    'Sub SetDefault(ByVal FormUID As String)
    '    Try

    '        objForm.Freeze(True)

    '        objForm = objMain.objApplication.Forms.Item(FormUID)

    '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TENXMENU")
    '        oDBs_Details = objForm.DataSources.DBDataSources.Item("@TENXMENUC0")
    '        oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TENXMENUC1")

    '        objMatrix = objForm.Items.Item("Item_2").Specific
    '        objMatrix2 = objForm.Items.Item("Item_13").Specific




    '        objForm.Freeze(False)
    '    Catch ex As Exception
    '        objForm.Freeze(False)
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub

    'Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
    '    Try
    '        If pVal.MenuUID = "TENXMENU" And pVal.BeforeAction = False Then

    '            Me.CreateForm()

    '            Me.SetDefault(objForm.UniqueID)


    '        ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then


    '            objForm = objMain.objApplication.Forms.GetForm("TENXMENU_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
    '            Me.SetDefault(objForm.UniqueID)

    '        End If

    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    End Try
    'End Sub

    'Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    '    Try
    '        Select Case pVal.EventType


    '            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
    '                objForm = objMain.objApplication.Forms.Item(FormUID)

    '                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TENXMENU")
    '                oDBs_Details = objForm.DataSources.DBDataSources.Item("@TENXMENUC0")
    '                oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TENXMENUC1")

    '                objMatrix = objForm.Items.Item("Item_2").Specific
    '                objMatrix2 = objForm.Items.Item("Item_13").Specific

    '                If pVal.ItemUID = "Item_3" And pVal.BeforeAction = False And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
    '                    Try
    '                        objForm.Freeze(True)

    '                        Me.LoadMenuDataUpdate(IO.Directory.GetParent(Application.ExecutablePath).ToString & "\Menu.xml")

    '                        objForm.Freeze(False)
    '                    Catch ex As Exception
    '                        objForm.Freeze(False)
    '                        objMain.objApplication.StatusBar.SetText(ex.Message)
    '                    End Try
    '                End If

    '                If pVal.ItemUID = "Item_14" And pVal.BeforeAction = False Then
    '                    Try
    '                        objForm.Freeze(True)


    '                        Dim MatrixRow As Integer = objMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder)


    '                        Dim GridRow As Integer = oGridUser.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)

    '                        If MatrixRow < 1 Or GridRow < 0 Then
    '                            objMain.objApplication.StatusBar.SetText("Please select one user/group and one menu")
    '                            objForm.Freeze(False)
    '                            Exit Try
    '                        End If

    '                        Dim CurrentUser As String = oGridUser.DataTable.GetValue("User Code", oGridUser.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder))
    '                        Dim CurrentMenu As String = objMatrix.Columns.Item("Col_0").Cells.Item(objMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).Specific.Value


    '                        For k As Integer = 1 To objMatrix2.VisualRowCount

    '                            objMain.objApplication.StatusBar.SetText("Please Wait Menu validating...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

    '                            Dim LoopUser As String = objMatrix2.Columns.Item("Col_1").Cells.Item(k).Specific.Value
    '                            Dim LoopMenu As String = objMatrix2.Columns.Item("Col_3").Cells.Item(k).Specific.Value

    '                            If LoopUser.Trim = CurrentUser.Trim And CurrentMenu.Trim = LoopMenu.Trim Then
    '                                objMain.objApplication.StatusBar.SetText("Same User and menu already added")
    '                                objForm.Freeze(False)
    '                                Exit Try
    '                            End If


    '                        Next


    '                        objMatrix2.AddRow()
    '                        oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, objMatrix2.VisualRowCount)
    '                        oDBs_Details2.SetValue("U_USRID", oDBs_Details2.Offset, oGridUser.DataTable.GetValue("User Code", oGridUser.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)))
    '                        oDBs_Details2.SetValue("U_USRNM", oDBs_Details2.Offset, oGridUser.DataTable.GetValue("UserName", oGridUser.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)))
    '                        oDBs_Details2.SetValue("U_MENUID", oDBs_Details2.Offset, objMatrix.Columns.Item("Col_0").Cells.Item(objMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).Specific.Value)
    '                        oDBs_Details2.SetValue("U_MENUNM", oDBs_Details2.Offset, objMatrix.Columns.Item("Col_1").Cells.Item(objMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).Specific.Value)
    '                        oDBs_Details2.SetValue("U_DSBL", oDBs_Details2.Offset, "Y")

    '                        objMatrix2.SetLineData(objMatrix2.VisualRowCount)

    '                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
    '                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
    '                        End If

    '                        objForm.Freeze(False)
    '                    Catch ex As Exception
    '                        objForm.Freeze(False)
    '                        objMain.objApplication.StatusBar.SetText(ex.Message)
    '                    End Try
    '                End If



    '            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
    '                objForm = objMain.objApplication.Forms.Item(FormUID)

    '                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TENXMENU")
    '                oDBs_Details = objForm.DataSources.DBDataSources.Item("@TENXMENUC0")
    '                oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TENXMENUC1")

    '                objMatrix = objForm.Items.Item("Item_2").Specific
    '                objMatrix2 = objForm.Items.Item("Item_13").Specific


    '                Dim oCFL As SAPbouiCOM.ChooseFromList
    '                Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
    '                Dim CFL_Id As String
    '                CFL_Id = CFLEvent.ChooseFromListUID
    '                oCFL = objForm.ChooseFromLists.Item(CFL_Id)
    '                Dim oDT As SAPbouiCOM.DataTable
    '                oDT = CFLEvent.SelectedObjects
    '                objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)


    '                If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
    '                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
    '                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)



    '                End If


    '        End Select
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    End Try
    'End Sub


    'Sub SetNewLine(ByVal FormUID As String)
    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)

    '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TENXMENU")
    '        oDBs_Details = objForm.DataSources.DBDataSources.Item("@TENXMENUC0")
    '        oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TENXMENUC1")

    '        objMatrix = objForm.Items.Item("Item_2").Specific
    '        objMatrix2 = objForm.Items.Item("Item_13").Specific

    '        objForm.Freeze(True)



    '        objMatrix.AddRow()
    '        oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
    '        oDBs_Details.SetValue("U_PID", oDBs_Details.Offset, "")
    '        oDBs_Details.SetValue("U_PMN", oDBs_Details.Offset, "")
    '        oDBs_Details.SetValue("U_MID", oDBs_Details.Offset, "")
    '        oDBs_Details.SetValue("U_MNM", oDBs_Details.Offset, "")
    '        oDBs_Details.SetValue("U_VSBL", oDBs_Details.Offset, "")

    '        objMatrix.SetLineData(objMatrix.VisualRowCount)


    '        objForm.Freeze(False)

    '    Catch ex As Exception
    '        objForm.Freeze(False)
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub

End Class

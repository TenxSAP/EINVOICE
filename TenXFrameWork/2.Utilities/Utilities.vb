

Imports System.Reflection
Imports SAPbobsCOM


Public Class Utilities
    Public strLastErrorCode As String
    Public strLastError As String
    Private objForm As SAPbouiCOM.Form
    Dim oItem As SAPbouiCOM.Item
    Dim TempCode As String
    Dim BankFileName As String

#Region " Get Application "
    Public Function GetApplication() As SAPbouiCOM.Application
        Dim objApp As SAPbouiCOM.Application
        Try
            Dim objSboGuiApi As New SAPbouiCOM.SboGuiApi
            Dim strConnectionString As String = String.Empty
            If strConnectionString = "" Then
                If Environment.GetCommandLineArgs().Length = 1 Then
                    strConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
                Else
                    strConnectionString = Environment.GetCommandLineArgs.GetValue(1)
                End If
            End If
            objSboGuiApi = New SAPbouiCOM.SboGuiApi
            objSboGuiApi.Connect(strConnectionString)
            objApp = objSboGuiApi.GetApplication()
            Return objApp
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            System.Windows.Forms.Application.Exit()
            Return Nothing
        End Try
    End Function
#End Region

#Region " Get Company "
    Public Function GetCompany(ByVal SBOApplication As SAPbouiCOM.Application) As SAPbobsCOM.Company
        Dim objCompany As SAPbobsCOM.Company

        Dim strCookie As String
        Dim strCookieContext As String

        Try
            objCompany = New SAPbobsCOM.Company
            strCookie = objCompany.GetContextCookie
            strCookieContext = SBOApplication.Company.GetConnectionContext(strCookie)
            objCompany.SetSboLoginContext(strCookieContext)
            If objCompany.Connect <> 0 Then
                strLastError = "Connection Error"
                SBOApplication.StatusBar.SetText("GetCompany() Connection Error", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return Nothing
            End If
            Return objCompany
        Catch ex As Exception
            SBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return Nothing
        End Try
    End Function
#End Region

#Region " Load Form "
    Public Sub LoadForm(ByVal XMLFile As String, ByVal FormType As String, Optional ByVal FileType As ResourceType = ResourceType.Content)
        Try
            Dim AppAssemblty As Assembly = Assembly.GetExecutingAssembly()
            Dim sExecutingAssemblyNmae As String = AppAssemblty.GetName().Name.ToString()
            Dim xmldoc As New Xml.XmlDocument
            XMLFile = sExecutingAssemblyNmae + "." + XMLFile

            '// load the form to the SBO application in one batch
            Dim Streaming As System.IO.Stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(XMLFile)
            Dim StreamRead As New System.IO.StreamReader(Streaming, True)
            xmldoc.LoadXml(StreamRead.ReadToEnd)
            StreamRead.Close()

            If Not xmldoc.SelectSingleNode("//form") Is Nothing Then
                xmldoc.SelectSingleNode("//form").Attributes.GetNamedItem("uid").Value = xmldoc.SelectSingleNode("//form").Attributes.GetNamedItem("uid").Value & "_" & objMain.objApplication.Forms.Count
                Dim a As String = xmldoc.InnerXml
                objMain.objApplication.LoadBatchActions(xmldoc.InnerXml)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Add Column - ## Not Used "
    Private Sub addCol(ByVal strTab As String, ByVal strCol As String, ByVal strDesc As String, ByVal nType As Integer, Optional ByVal nEditSize As Integer = 10, Optional ByVal nSubType As Integer = 0)

        Dim oUFields As SAPbobsCOM.UserFieldsMD
        Dim nError As Integer

        oUFields = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUFields.TableName = strTab
        oUFields.Name = strCol
        oUFields.Type = nType
        oUFields.SubType = nSubType
        oUFields.Description = strDesc
        oUFields.EditSize = nEditSize
        nError = oUFields.Add()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
        GC.Collect()
        GC.WaitForPendingFinalizers()
        If nError <> 0 Then
            'MsgBox(strCol & " table could not be added")
        End If
    End Sub
#End Region

#Region "Create Table"
    Public Function CreateTable(ByVal TableName As String, ByVal TableDescription As String, ByVal TableType As SAPbobsCOM.BoUTBTableType) As Boolean
        Dim intRetCode As Integer
        Dim objUserTableMD As SAPbobsCOM.UserTablesMD
        objUserTableMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
        Try
            If (Not objUserTableMD.GetByKey(TableName)) Then
                objMain.objApplication.StatusBar.SetText("Creating table... [@" & TableName & "]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objUserTableMD.TableName = TableName
                objUserTableMD.TableDescription = TableDescription
                objUserTableMD.TableType = TableType
                intRetCode = objUserTableMD.Add()
                If (intRetCode = 0) Then
                    Return True
                Else
                    'Vj Added for testing///////////////
                    Dim lret As Integer
                    Dim sret As String = String.Empty
                    objMain.objCompany.GetLastError(lret, sret)
                    objMain.objApplication.MessageBox(lret & " : " & sret)
                    '//////////////////Done
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            objMain.objApplication.MessageBox(ex.Message)
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserTableMD)
            GC.Collect()
        End Try
    End Function

#End Region

    Public Function GetNextDocNumSeries(ByRef Objform As SAPbouiCOM.Form, ByVal UDOName As String, ByVal Series As String) As Integer
        Dim Str As String
        Dim oRs As SAPbobsCOM.Recordset
        Dim DocNum As Integer
        oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Str = "Select ""Series"" From NNM1 Where ""ObjectCode"" = '" & UDOName & "' And ""Series""= '" & Series & "'"
        Try
            oRs.DoQuery(Str)
            oRs.MoveFirst()
            If oRs.RecordCount > 0 Then
                DocNum = Objform.BusinessObject.GetNextSerialNumber(oRs.Fields.Item(0).Value, UDOName)
            End If
            If DocNum = 0 Then DocNum = 1
            Return DocNum
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("DN: " + ex.Message)
        End Try
    End Function

    Public Sub ComboBoxLoadValuesSeries(ByVal objCombo As SAPbouiCOM.ComboBox, ByVal QueryAsValueAndDescription As String)
        Try
            If (objCombo.ValidValues.Count <> 0) Then
                For R As Integer = objCombo.ValidValues.Count - 1 To 0 Step -1
                    Try
                        objCombo.ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                    Catch ex As Exception
                    End Try
                Next
            End If

            If objCombo.ValidValues.Count = 0 Then
                Dim objRecSet
                objRecSet = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecSet.DoQuery(QueryAsValueAndDescription)
                objRecSet.MoveFirst()
                'objCombo.ValidValues.Add("", "")
                While Not objRecSet.EoF
                    Try
                        objCombo.ValidValues.Add(objRecSet.Fields.Item(0).Value, objRecSet.Fields.Item(1).Value)
                    Catch ex As Exception
                    End Try
                    objRecSet.MoveNext()
                End While
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub ComboBoxLoadValues1(ByVal objCombo As SAPbouiCOM.ComboBox, ByVal QueryAsValueAndDescription As String)
        Try
            If (objCombo.ValidValues.Count <> 0) Then
                For R As Integer = objCombo.ValidValues.Count - 1 To 0 Step -1
                    Try
                        objCombo.ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                    Catch ex As Exception
                    End Try
                Next
            End If

            If objCombo.ValidValues.Count = 0 Then
                Dim objRecSet
                objRecSet = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecSet.DoQuery(QueryAsValueAndDescription)
                objRecSet.MoveFirst()
                While Not objRecSet.EoF
                    Try
                        objCombo.ValidValues.Add(objRecSet.Fields.Item(0).Value, objRecSet.Fields.Item(1).Value)
                    Catch ex As Exception
                    End Try
                    objRecSet.MoveNext()
                End While
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


#Region "Fields Creation"
    Public Sub AddAlphaField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, Optional ByVal DefaultValue As String = "")
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", DefaultValue)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Public Sub AddAlphaField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, Size, SAPbobsCOM.BoFldSubTypes.st_None, ValidValues, ValidDescriptions, SetValidValue)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        objUserFieldMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        Try
            strValue = ValidValues.Split(Convert.ToChar(","))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Invalid Valid Values")
            End If
            If (Not isColumnExist(TableName, ColumnName)) Then
                objMain.objApplication.StatusBar.SetText("Creating field...[" & ColumnName & "] of table [" & TableName & "]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName
                objUserFieldMD.Description = ColDescription
                objUserFieldMD.Type = FieldType
                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                objUserFieldMD.SubType = SubType
                objUserFieldMD.DefaultValue = SetValidValue
                If strValue.Length > 1 Then
                    For intLoop = 0 To strValue.GetLength(0) - 1
                        objUserFieldMD.ValidValues.Value = strValue(intLoop)
                        objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                        objUserFieldMD.ValidValues.Add()
                    Next
                End If
                If (objUserFieldMD.Add() <> 0) Then
                    objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
                End If
                'Else
                '    objMain.objApplication.StatusBar.SetText("Creating field...[" & ColumnName & "] of table [" & TableName & "]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                '    objUserFieldMD.TableName = TableName
                '    objUserFieldMD.Name = ColumnName
                '    objUserFieldMD.Description = ColDescription
                '    objUserFieldMD.Type = FieldType
                '    If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                '        objUserFieldMD.Size = Size
                '    Else
                '        objUserFieldMD.EditSize = Size
                '    End If
                '    objUserFieldMD.SubType = SubType
                '    objUserFieldMD.DefaultValue = SetValidValue
                '    For intLoop = 0 To strValue.GetLength(0) - 1
                '        objUserFieldMD.ValidValues.Value = strValue(intLoop)
                '        objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                '        objUserFieldMD.ValidValues.Add()
                '    Next
                '    If (objUserFieldMD.Update() <> 0) Then
                '        objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
                '    End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            GC.Collect()
        End Try
    End Sub
    Public Sub AddAlphaFieldLinkedTable(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, Optional ByVal DefaultValue As String = "", Optional ByVal sLinkedTable As String = "")
        Try
            addFieldLinkedTable(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", DefaultValue, sLinkedTable)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Public Sub addFieldLinkedTable(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String, ByVal sLinkedTable As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        objUserFieldMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        Try
            strValue = ValidValues.Split(Convert.ToChar(","))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Invalid Valid Values")
            End If

            If (Not isColumnExist(TableName, ColumnName)) Then
                objMain.objApplication.StatusBar.SetText("Creating field...[" & ColumnName & "] of table [" & TableName & "]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName
                objUserFieldMD.Description = ColDescription
                objUserFieldMD.Type = FieldType

                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                objUserFieldMD.SubType = SubType
                objUserFieldMD.DefaultValue = SetValidValue
                If strValue.Length > 1 Then
                    For intLoop = 0 To strValue.GetLength(0) - 1
                        objUserFieldMD.ValidValues.Value = strValue(intLoop)
                        objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                        objUserFieldMD.ValidValues.Add()

                    Next
                End If

                objUserFieldMD.LinkedTable = sLinkedTable

                If (objUserFieldMD.Add() <> 0) Then
                    objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
                End If

            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message)

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            GC.Collect()
        End Try
    End Sub

    Private Function isColumnExist(ByVal TableName As String, ByVal ColumnName As String) As Boolean
        Dim objRecordSet As SAPbobsCOM.Recordset
        objRecordSet = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            objRecordSet.DoQuery("SELECT COUNT(*) FROM CUFD WHERE ""TableID"" = '" & TableName & "' AND ""AliasID"" = '" & ColumnName.Trim & "'")


            If (Convert.ToInt16(objRecordSet.Fields.Item(0).Value) <> 0) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecordSet)
            GC.Collect()
        End Try
    End Function
    Public Sub AddFloatField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal SubType As SAPbobsCOM.BoFldSubTypes)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Float, 0, SubType, "", "", "")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Public Sub AddDateField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal SubType As SAPbobsCOM.BoFldSubTypes)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Date, 0, SubType, "", "", "")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Public Sub AddAlphaMemoField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer)

        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Memo, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Public Sub AddInteger(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal Size As Integer)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Numeric, Size, SubType, "", "", "")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Public Sub UniqueIDField(ByVal TableName As String, ByVal FieldName As String, ByVal oCompany As SAPbobsCOM.Company)

        '//****************************************************************************
        '// The UserKeysMD represents a meta-data object that allows you
        '// to add\remove user defined keys.
        '//****************************************************************************

        Dim oUserKeysMD As SAPbobsCOM.UserKeysMD

        '//flag
        Dim bFlagFirst As Boolean

        bFlagFirst = True

        '//****************************************************************************
        '// In any meta-data operation there should be no other object "alive"
        '// but the meta-data object, otherwise the operation will fail.
        '// This restriction is intended to prevent collisions.
        '//****************************************************************************

        '// The meta-data object must be initialized with a
        '// regular UserKeys object
        oUserKeysMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)

        '// Set the table name and the key name
        'oUserKeysMD.TableName = "OCRD" '// BP table
        'oUserKeysMD.KeyName = "BE_MyKey1"

        oUserKeysMD.TableName = TableName '// BP table
        oUserKeysMD.KeyName = FieldName


        '//*******************************************
        '// Add a column to a key button:
        '//-------------------------------------------
        '// To add an additional column to
        '// the key, an additional element must be
        '// created in the Elements collection.
        '// The Add method of the Elements collection
        '// must be used only as of the second element.

        '// Do not use the Add method for the first element
        If bFlagFirst = True Then
            bFlagFirst = False
        Else
            '// Add an item to the Elements collection
            oUserKeysMD.Elements.Add()
            strLastErrorCode = oCompany.GetLastErrorCode()
            strLastError = oCompany.GetLastErrorDescription()
        End If

        '// Set the column's alias
        oUserKeysMD.Elements.ColumnAlias = FieldName

        '// Determine whether the key is unique or not
        'oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES
        oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tNO

        '// Add the key
        oUserKeysMD.Add()
        strLastErrorCode = oCompany.GetLastErrorCode()
        strLastError = oCompany.GetLastErrorDescription()
        'If (oUserKeysMD <> DBNull.Value) Then
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD)

        'End If
    End Sub
    Public Sub AddLinkField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Memo, Size, SubType, "", "", "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub AddImageField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_Image, "", "", "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

#Region " Add Data to No Object Table"
    Public Function AddDataToNoObjectTable(ByVal TableName As String, ByVal Code As String, ByVal Name As String, Optional ByVal UDFName1 As String = "", Optional ByVal UDFValue1 As String = "")
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim lReturn As Integer
        Dim ErrorString As String
        oUserTable = objMain.objCompany.UserTables.Item(TableName)

        If oUserTable.GetByKey(Code) = False Then
            'Set default, mandatory fields
            oUserTable.Code = Code
            oUserTable.Name = Name

            'Set user field
            If UDFName1 <> String.Empty Then oUserTable.UserFields.Fields.Item(UDFName1).Value = UDFValue1
            oUserTable.Add()
            If lReturn <> 0 Then
                objMain.objCompany.GetLastError(lReturn, ErrorString)
                Return (ErrorString)
            End If
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable)
        Return ("")
    End Function
#End Region

#Region "add menus with xml"
    Public Sub LoadFromXML(ByRef FileName As String)

        Dim oXmlDoc As Xml.XmlDocument
        oXmlDoc = New Xml.XmlDocument
        '// load the content of the XML File
        Dim sPath As String
        sPath = IO.Directory.GetParent(Application.ExecutablePath).ToString
        oXmlDoc.Load(sPath & "\" & FileName)
        '// load the form to the SBO application in one batch
        objMain.objApplication.LoadBatchActions(oXmlDoc.InnerXml)
        sPath = objMain.objApplication.GetLastBatchResults()

    End Sub
#End Region

#Region " Check if Form Exists - ## Not Used "
    Public Function FormExist(ByVal FormUID As String) As Boolean
        Dim intLoop As Integer

        For intLoop = objMain.objApplication.Forms.Count - 1 To 0 Step -1
            If Trim(FormUID) = Trim(objMain.objApplication.Forms.Item(intLoop).UniqueID) Then
                Return True
            End If
        Next
        Return False
    End Function
#End Region

#Region " Get MaxCode "
    Public Function getMaxCode(ByVal sTable As String) As String
        Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim MaxCode As Integer
        Dim sCode As String
        Dim strSQL As String
        Try
            strSQL = "SELECT MAX(CAST(""Code"" AS INT)) AS Code From """ & sTable & """"
            oRS.DoQuery(strSQL)
            If Convert.ToString(oRS.Fields.Item(0).Value).Length > 0 Then
                MaxCode = oRS.Fields.Item(0).Value + 1
            Else
                MaxCode = 10000001
            End If
            sCode = MaxCode
            Return sCode

        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
        End Try
    End Function
#End Region

#Region " UDO Document Numbering "
    Public Function GetNextDocNum(ByRef Objform As SAPbouiCOM.Form, ByVal UDOName As String, Optional ByVal SeriesName As String = "Primary") As Integer
        Dim Str As String
        Dim oRs As SAPbobsCOM.Recordset
        Dim DocNum As Integer
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
            objMain.objApplication.StatusBar.SetText("DN: " + ex.Message)
        End Try
    End Function
#End Region
    Public Function GetDocumentEntry(ByVal sDocNum As String, ByVal sDocSeries As String, ByVal DocTable As String)
        Try

            Dim GetDocEntry As String = "Select ""DocEntry"" From """ & DocTable & """ Where ""DocNum"" = '" & sDocNum & "' And ""Series"" = '" & sDocSeries & "' "
            Dim oRsGetDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDocEntry.DoQuery(GetDocEntry)

            Return oRsGetDocEntry.Fields.Item("DocEntry").Value.ToString
        Catch ex As Exception
        End Try
    End Function

    Public Function GetUserName(ByVal UserCode As String)
        Try

            Dim UserName As String = "Select ""U_NAME"" From OUSR Where ""USER_CODE"" = '" & UserCode & "' "
            Dim oRsUserName As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsUserName.DoQuery(UserName)

            Return oRsUserName.Fields.Item("U_Name").Value.ToString
        Catch ex As Exception
        End Try
    End Function

#Region " Load DataSource from DB "
    Public Sub RefreshDatasourceFromDB(ByVal FormUID As String, ByRef oDBs_Head As SAPbouiCOM.DBDataSource, ByVal ConditionAlias As String, ByVal ConditionValue As String)
        Try
            Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)
            Dim oConditions As SAPbouiCOM.Conditions = New SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            oCondition = oConditions.Add()
            oCondition.Alias = ConditionAlias
            oCondition.ComparedAlias = ConditionAlias
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = ConditionValue
            oDBs_Head.Query(oConditions)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

#Region "      ***Load Values to ComboBoxes***       "
    Public Sub ComboBoxLoadValues(ByVal objCombo As SAPbouiCOM.ComboBox, ByVal QueryAsValueAndDescription As String)
        Try
            If (objCombo.ValidValues.Count <> 0) Then
                For R As Integer = objCombo.ValidValues.Count - 1 To 0 Step -1
                    Try
                        objCombo.ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                    Catch ex As Exception
                    End Try
                Next
            End If

            If objCombo.ValidValues.Count = 0 Then
                Dim objRecSet
                objRecSet = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecSet.DoQuery(QueryAsValueAndDescription)
                objRecSet.MoveFirst()
                objCombo.ValidValues.Add("", "")
                While Not objRecSet.EoF
                    Try
                        objCombo.ValidValues.Add(objRecSet.Fields.Item(0).Value, objRecSet.Fields.Item(1).Value)
                    Catch ex As Exception
                    End Try
                    objRecSet.MoveNext()
                End While
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Sub MatrixComboBoxValues(ByVal oColumn As SAPbouiCOM.Column, ByVal QueryAsValueAndDescription As String)
        Try
            If oColumn.ValidValues.Count = 0 Then
                Dim objRecSet
                objRecSet = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecSet.DoQuery(QueryAsValueAndDescription)
                objRecSet.MoveFirst()
                oColumn.ValidValues.Add("", "")
                While Not objRecSet.EoF
                    Try
                        oColumn.ValidValues.Add(objRecSet.Fields.Item(0).Value, objRecSet.Fields.Item(1).Value)
                    Catch ex As Exception
                    End Try
                    objRecSet.MoveNext()
                End While
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

#End Region

    Function SetAttachMentFile()
        Dim ShowFolderBrowserThread As Threading.Thread
        Try
            'Me.ShowFolderBrowser()
            ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFolderBrowser)
            If ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Unstarted Then
                ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA)
                ShowFolderBrowserThread.Start()
                'ElseIf ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Stopped Then
                '    ShowFolderBrowserThread.Start()
                '    ShowFolderBrowserThread.Join()
            End If
            While ShowFolderBrowserThread.ThreadState = Threading.ThreadState.Running
                System.Windows.Forms.Application.DoEvents()
            End While
            If BankFileName <> "" Then
                Return BankFileName
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'oApplication.MessageBox("FileFile Method Failed : " & ex.Message)
        End Try
        Return ""
    End Function
    Public Function AddOrUpdateQuery(company As SAPbobsCOM.Company, queryName As String, categoryId As Integer, sqlText As String, ByVal Type As String) As Boolean
        Try
            Dim oUserQuery As SAPbobsCOM.UserQueries = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries)

            ' Check if query already exists
            Dim rs As SAPbobsCOM.Recordset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim Qry As String = "SELECT ""IntrnalKey"", ""QCategory"" FROM OUQR WHERE ""QName"" = '" & queryName & "'"
            rs.DoQuery(Qry)
            Dim key As Integer = CInt(rs.Fields.Item("IntrnalKey").Value)
            Dim categoryID1 As Integer = CInt(rs.Fields.Item("QCategory").Value)

            If rs.RecordCount > 0 Then
                ' Update existing query


                If oUserQuery.GetByKey(key, categoryID1) Then
                    oUserQuery.Query = sqlText
                    If oUserQuery.Update() <> 0 Then
                        Throw New Exception("Error updating query: " & company.GetLastErrorDescription())
                    Else
                        If Type = "Alert" Then
                            Dim rs1 As SAPbobsCOM.Recordset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim Qry1 As String = "SELECT ""IntrnalKey"" FROM OUQR WHERE ""QName"" = '" & queryName & "'"
                            rs1.DoQuery(Qry1)
                            If rs1.RecordCount > 0 Then
                                Dim key1 As Integer = CInt(rs1.Fields.Item("IntrnalKey").Value)

                                Dim rsAlertCheck As SAPbobsCOM.Recordset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim alertCheckQuery As String = "SELECT T0.""Code"", T0.""Name"" FROM OALT T0 where T0.""Name"" = '" & queryName & "'"
                                rsAlertCheck.DoQuery(alertCheckQuery)

                                If rsAlertCheck.RecordCount = 0 Then

                                    CreateCustomAlert(queryName, key1, 1, 30, AlertManagementFrequencyType.atfi_Minutes)
                                End If
                            End If

                        End If
                    End If
                End If
            Else
                ' Add new query
                oUserQuery.QueryCategory = categoryId
                oUserQuery.Query = sqlText
                oUserQuery.QueryDescription = queryName
                If oUserQuery.Add() <> 0 Then
                    Throw New Exception("Error adding query: " & company.GetLastErrorDescription())
                Else
                    If Type = "Alert" Then
                        Dim rs1 As SAPbobsCOM.Recordset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim Qry1 As String = "SELECT ""IntrnalKey"" FROM OUQR WHERE ""QName"" = '" & queryName & "'"
                        rs1.DoQuery(Qry1)
                        If rs1.RecordCount > 0 Then
                            Dim key1 As Integer = CInt(rs1.Fields.Item("IntrnalKey").Value)

                            Dim rsAlertCheck As SAPbobsCOM.Recordset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim alertCheckQuery As String = "SELECT T0.""Code"", T0.""Name"" FROM OALT T0 where T0.""Name"" = '" & queryName & "'"
                            rsAlertCheck.DoQuery(alertCheckQuery)

                            If rsAlertCheck.RecordCount = 0 Then

                                CreateCustomAlert(queryName, key1, 1, 30, AlertManagementFrequencyType.atfi_Minutes)
                            End If
                        End If

                    End If
                End If
            End If
            ' Clean up
            ' Marshal.ReleaseComObject(rs)
            ' Marshal.ReleaseComObject(oUserQuery)
            GC.Collect()
            objMain.objApplication.StatusBar.SetText("Query successfully added or updated in Query Manager:" & queryName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error while adding query:" & ex.Message & "Query:" & queryName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Public Function CreateCustomAlert(ByVal alertName As String, ByVal queryId As Integer, ByVal recipientUserCode As String, ByVal frequencyInterval As Integer, ByVal frequencyType As AlertManagementFrequencyType) As Boolean
        Try
            ' Get the Company Service
            Dim oCmpSrv As CompanyService = objMain.objCompany.GetCompanyService()

            ' Get the Alert Management Service
            Dim oAlertService As AlertManagementService = DirectCast(oCmpSrv.GetBusinessService(ServiceTypes.AlertManagementService), AlertManagementService)

            ' Create a new Alert Management object
            Dim oAlertManagement As AlertManagement = DirectCast(oAlertService.GetDataInterface(AlertManagementServiceDataInterfaces.atsdiAlertManagement), AlertManagement)
            ' Set alert properties
            oAlertManagement.Name = alertName
            oAlertManagement.Priority = AlertManagementPriorityEnum.atp_Low ' You can change priority as needed
            oAlertManagement.Active = BoYesNoEnum.tYES
            oAlertManagement.QueryID = queryId ' Assign the ID of your user-defined query
            oAlertManagement.FrequencyInterval = frequencyInterval
            oAlertManagement.FrequencyType = frequencyType
            oAlertManagement.SaveHistory = BoYesNoEnum.tYES ' Example: Save alert history

            Dim oRecipient As AlertManagementRecipient = oAlertManagement.AlertManagementRecipients.Add()
            oRecipient.UserCode = recipientUserCode
            oRecipient.SendInternal = BoYesNoEnum.tYES ' Send as an internal message in SAP Business One
            oRecipient.SendEmail = BoYesNoEnum.tYES ' Send as email
            ' oRecipient.EmailAddress = "recipient@example.com" ' Optionally specify an email address

            'oAlertService.AddAlert(oAlertManagement)
            oAlertService.AddAlertManagement(oAlertManagement)

            objMain.objApplication.StatusBar.SetText("Query successfully added or updated in Query Manager", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            ' MessageBox.Show("Alert '" & alertName & "' created successfully!")
            Return True

        Catch ex As Exception
            ' MessageBox.Show("Error creating alert: " & ex.Message)
            Return False
        End Try
    End Function

    Public Sub AddAuthorization(ByVal authCode As String, ByVal authName As String, ByVal formUID As String, ByVal ParentForm As String, ByVal permission As SAPbobsCOM.BoUPTOptions)
        Dim oUserPermission As SAPbobsCOM.UserPermissionTree
        oUserPermission = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree), SAPbobsCOM.UserPermissionTree)

        ' Check if authorization already exists
        If oUserPermission.GetByKey(authCode) Then
            ' Optionally update existing
            oUserPermission.Name = authName
            oUserPermission.Options = permission
            oUserPermission.ParentID = ParentForm
            If oUserPermission.Update() <> 0 Then
                MsgBox("Failed to update auth: " & objMain.objCompany.GetLastErrorDescription())
            End If
        Else
            ' Create new authorization
            oUserPermission.PermissionID = authCode
            oUserPermission.Name = authName
            oUserPermission.Options = permission
            oUserPermission.ParentID = ParentForm
            oUserPermission.UserPermissionForms.FormType = formUID
            oUserPermission.UserPermissionForms.Add()

            If oUserPermission.Add() <> 0 Then
                MsgBox("Failed to add auth: " & objMain.objCompany.GetLastErrorDescription())
            End If
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserPermission)
        oUserPermission = Nothing
        GC.Collect()
    End Sub

    Public Sub ShowFolderBrowser()
        Dim MyProcs() As System.Diagnostics.Process
        BankFileName = ""
        Dim OpenFile As New OpenFileDialog
        Try
            OpenFile.Multiselect = False
            OpenFile.Filter = "All files(*.)|*.*" '   "|*.*"
            Dim filterindex As Integer = 0
            Try
                filterindex = 0
            Catch ex As Exception
            End Try
            OpenFile.FilterIndex = filterindex
            OpenFile.RestoreDirectory = True
            MyProcs = Process.GetProcessesByName("SAP Business One")
            'If MyProcs.Length = 2 Then
            'For i As Integer = 0 To MyProcs.Length - 2
            Dim MyWindow As New clsWindowWrapper(MyProcs(0).MainWindowHandle)
            Dim ret As DialogResult = OpenFile.ShowDialog(MyWindow)
            If ret = DialogResult.OK Then
                BankFileName = OpenFile.FileName
                OpenFile.Dispose()
            Else
                System.Windows.Forms.Application.ExitThread()
            End If
            'Next
            'End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            BankFileName = ""
        Finally
            OpenFile.Dispose()
        End Try
    End Sub

#Region " Checking DataType"
    Public Function IsAlpha(ByVal str As String) As Boolean
        Try
            Dim i As Integer
            For i = 0 To str.Length - 1
                If Not Char.IsLetter(str, i) Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception

        End Try
    End Function

    Public Function IsNumeric(ByVal str As String) As Boolean
        Try
            Dim i As Integer
            If str.Contains(".") = True Then
                Return False
            End If
            For i = 0 To str.Length - 1
                If Not Char.IsNumber(str, i) Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception

        End Try
    End Function

    Public Function IsFloat(ByVal str As String) As Boolean
        Try
            Dim i As Integer
            If str.Substring(0, 1) = "." Then
                Return False
            End If
            If str.Contains(".") = False Then
                Return False
            Else
                str = str.Remove(str.LastIndexOfAny("."), 1)
            End If
            For i = 0 To str.Length - 1
                If Not Char.IsNumber(str, i) Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception

        End Try
    End Function

    Function IsPercentage(ByVal str As String) As Boolean
        Try
            Dim i As Integer
            If str.Contains(".") = True Then
                str = str.Remove(str.LastIndexOfAny("."), 1)
                If str.Contains(".") = True Then
                    Return False
                End If
            End If
            For i = 0 To str.Length - 1
                If Not Char.IsNumber(str, i) Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception

        End Try
    End Function
#End Region

#Region " Adding Items To Forms"

    Public Sub AddLabel(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer,
                        ByVal iWidth As Integer, ByVal iCaption As String, ByVal iLink As String, Optional ByVal iHeight As Integer = 14,
                        Optional ByVal iFromPane As Integer = 0, Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.LinkTo = iLink
        oItem.Specific.Caption = iCaption

    End Sub


    Public Sub AddEditBox(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer,
                          ByVal iWidth As Integer, ByVal TableName As String, ByVal UdFName As String,
                          ByVal LinkTo As String, Optional ByVal iHeight As Integer = 14, Optional ByVal iFromPane As Integer = 0,
                          Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.Height = iHeight
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.LinkTo = LinkTo
        oItem.Height = iHeight
        oItem.Specific.DataBind.SetBound(True, TableName, UdFName)
    End Sub

    Public Sub AddExtendedEditBox(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer,
                                  ByVal iLeft As Integer, ByVal iWidth As Integer, ByVal TableName As String,
                                  ByVal UdFName As String, ByVal LinkTo As String, Optional ByVal iFromPane As Integer = 0,
                                  Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.Height = 80
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.LinkTo = LinkTo
        oItem.Specific.DataBind.SetBound(True, TableName, UdFName)
    End Sub
    Public Sub AddComboBox(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer,
                             ByVal iWidth As Integer, ByVal iHight As Integer, ByVal LinkTo As String, ByVal TableName As String, ByVal UdFName As String, Optional ByVal iFromPane As Integer = 0,
                             Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.Height = iHight
        oItem.LinkTo = LinkTo
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.Specific.DataBind.SetBound(True, TableName, UdFName)
    End Sub
    Public Sub AddComboBox1(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer,
                           ByVal iWidth As Integer, ByVal TableName As String, ByVal UdFName As String, ByVal Link As String, Optional ByVal iFromPane As Integer = 0,
                           Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.LinkTo = Link
        oItem.Specific.DataBind.SetBound(True, TableName, UdFName)
    End Sub


    Public Sub AddFolder(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer, ByVal iWidth As Integer,
                         ByVal UdFName As String, ByVal Caption As String, ByVal AliasName As String,
                         ByVal GroupItem As String)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_FOLDER)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        Dim oFolder As SAPbouiCOM.Folder
        oFolder = oItem.Specific
        oFolder.Caption = Caption
        'oFolder.DataBind.SetBound(True, "", AliasName)
        oFolder.GroupWith(GroupItem)
    End Sub

    Public Sub AddButton(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer, ByVal iWidth As Integer,
                       ByVal iHight As Integer, ByVal LinkTo As String, ByVal Caption As String, Optional ByVal iFromPane As Integer = 0,
                                  Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_BUTTON)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.Height = iHight
        oItem.LinkTo = LinkTo
        Dim btn As SAPbouiCOM.Button = objForm.Items.Item(ItemUID).Specific
        btn.Caption = Caption
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
    End Sub


    Public Sub AddLinkButton(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer,
                             ByVal iLeft As Integer, ByVal iLinkTo As String, ByVal LinkedObject As Integer)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
        oItem.Top = iTop
        oItem.Left = iLeft
        Dim LinkBtn As SAPbouiCOM.LinkedButton
        LinkBtn = oItem.Specific
        oItem.LinkTo = iLinkTo
        If LinkedObject <> 0 Then
            LinkBtn.LinkedObject = LinkedObject
        End If
    End Sub

    Public Sub AddCheckBox(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer,
                           ByVal iWidth As Integer, ByVal TableName As String, ByVal UdFName As String, ByVal iCaption As String, Optional ByVal iHeight As Integer = 14, Optional ByVal iFromPane As Integer = 0,
                                  Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.Height = iHeight
        oItem.Specific.DataBind.SetBound(True, TableName, UdFName)
        oItem.Specific.Caption = iCaption
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
    End Sub
#End Region

#Region "Conversions"
    Function RupeesToWord(ByVal MyNumber As String) As String
        Dim Temp As String
        Dim Rupees As String = String.Empty
        Dim Paisa As String = String.Empty
        Dim DecimalPlace As String = String.Empty
        Dim iCount As String = String.Empty
        Dim Hundred As String = String.Empty
        Dim Words As String = String.Empty

        Dim ValidateNumber As String = MyNumber

        Dim place(9) As String
        place(0) = " Thousand "
        place(2) = " Lakh "
        place(4) = " Crore "
        place(6) = " Hundred "
        place(8) = " Kharab "
        If ValidateNumber.Length > 9 Then
            If ValidateNumber.Length = 10 And ValidateNumber.Substring(1, ValidateNumber.Length - 1) = "0" Then
                place(4) = " Crore "
                place(6) = " Hundred Crore "
            ElseIf ValidateNumber.Length = 11 And (ValidateNumber.Substring(2, ValidateNumber.Length - 2) = "0") Then
                place(4) = " Crore "
                place(6) = " Hundred "
            Else
                place(4) = " Crore "
                place(6) = " Hundred "
            End If
        End If

        On Error Resume Next
        ' Convert MyNumber to a string, trimming extra spaces.
        MyNumber = Trim(Str(MyNumber))

        ' Find decimal place.
        DecimalPlace = InStr(MyNumber, ".")

        ' If we find decimal place...
        If DecimalPlace > 0 Then
            ' Convert Paisa
            Temp = Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2)
            Paisa = " and " & ConvertTens(Temp) & " Paisa"

            ' Strip off paisa from remainder to convert.
            MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
        End If

        '===============================================================
        Dim TM As String  ' If MyNumber between Rs.1 To 99 Only.
        TM = Right(MyNumber, 2)

        If Len(MyNumber) > 0 And Len(MyNumber) <= 2 Then
            If Len(TM) = 1 Then
                Words = ConvertDigit(TM)
                RupeesToWord = "Rupees " & Words & Paisa & " Only"

                Exit Function

            Else
                If Len(TM) = 2 Then
                    Words = ConvertTens(TM)
                    RupeesToWord = "Rupees " & Words & Paisa & " Only"
                    Exit Function

                End If
            End If
        End If
        '===============================================================


        ' Convert last 3 digits of MyNumber to ruppees in word.
        Hundred = ConvertHundreds(Right(MyNumber, 3))
        ' Strip off last three digits
        MyNumber = Left(MyNumber, Len(MyNumber) - 3)

        iCount = 0
        Do While MyNumber <> ""
            'Strip last two digits
            Temp = Right(MyNumber, 2)
            If Len(MyNumber) = 1 Then


                If Trim(Words) = "Thousand" Or
                Trim(Words) = "Lakh  Thousand" Or
                Trim(Words) = "Lakh" Or
                Trim(Words) = "Crore" Or
                Trim(Words) = "Crore  Lakh  Thousand" Or
                Trim(Words) = "Hundred  Crore  Lakh  Thousand" Or
                Trim(Words) = "Hundred" Or
                Trim(Words) = "Kharab  Hundred  Crore  Lakh  Thousand" Or
                Trim(Words) = "Kharab" Then

                    Words = ConvertDigit(Temp) & place(iCount)
                    MyNumber = Left(MyNumber, Len(MyNumber) - 1)

                Else

                    Words = ConvertDigit(Temp) & place(iCount) & Words
                    MyNumber = Left(MyNumber, Len(MyNumber) - 1)

                End If
            Else

                If Trim(Words) = "Thousand" Or
                   Trim(Words) = "Lakh  Thousand" Or
                   Trim(Words) = "Lakh" Or
                   Trim(Words) = "Crore" Or
                   Trim(Words) = "Crore  Lakh  Thousand" Or
                   Trim(Words) = "Hundred  Crore  Lakh  Thousand" Or
                   Trim(Words) = "Hundred" Then


                    Words = ConvertTens(Temp) & place(iCount)


                    MyNumber = Left(MyNumber, Len(MyNumber) - 2)
                Else

                    '=================================================================
                    ' if only Lakh, Crore, Arab, Kharab

                    If Trim(ConvertTens(Temp) & place(iCount)) = "Lakh" Or
                       Trim(ConvertTens(Temp) & place(iCount)) = "Crore" Or
                       Trim(ConvertTens(Temp) & place(iCount)) = "Hundred" Then

                        Words = Words
                        MyNumber = Left(MyNumber, Len(MyNumber) - 2)
                    Else
                        Words = ConvertTens(Temp) & place(iCount) & Words
                        MyNumber = Left(MyNumber, Len(MyNumber) - 2)
                    End If

                End If
            End If

            iCount = iCount + 2
        Loop

        RupeesToWord = "Rupees " & Words & Hundred & Paisa & " Only"

    End Function

    Private Function ConvertHundreds(ByVal MyNumber As String) As String
        Dim Result As String = String.Empty
        'Return String.Empty
        ' Exit if there is nothing to convert.
        If Val(MyNumber) = 0 Then
            Return Nothing
            'Exit Function
        End If



        ' Append leading zeros to number.
        MyNumber = Right("000" & MyNumber, 3)

        ' Do we have a hundreds place digit to convert?
        If Left(MyNumber, 1) <> "0" Then
            Result = ConvertDigit(Left(MyNumber, 1)) & " Hundred "
        End If

        ' Do we have a tens place digit to convert?
        If Mid(MyNumber, 2, 1) <> "0" Then
            Result = Result & ConvertTens(Mid(MyNumber, 2))
        Else
            ' If not, then convert the ones place digit.
            Result = Result & ConvertDigit(Mid(MyNumber, 3))
        End If

        ConvertHundreds = Trim(Result)
    End Function

    Private Function ConvertTens(ByVal MyTens As String) As String
        Dim Result As String = String.Empty

        ' Is value between 10 and 19?
        If Val(Left(MyTens, 1)) = 1 Then
            Select Case Val(MyTens)
                Case 10 : Result = "Ten"
                Case 11 : Result = "Eleven"
                Case 12 : Result = "Twelve"
                Case 13 : Result = "Thirteen"
                Case 14 : Result = "Fourteen"
                Case 15 : Result = "Fifteen"
                Case 16 : Result = "Sixteen"
                Case 17 : Result = "Seventeen"
                Case 18 : Result = "Eighteen"
                Case 19 : Result = "Nineteen"
                Case Else
            End Select
        Else
            ' .. otherwise it's between 20 and 99.
            Select Case Val(Left(MyTens, 1))
                Case 2 : Result = "Twenty "
                Case 3 : Result = "Thirty "
                Case 4 : Result = "Forty "
                Case 5 : Result = "Fifty "
                Case 6 : Result = "Sixty "
                Case 7 : Result = "Seventy "
                Case 8 : Result = "Eighty "
                Case 9 : Result = "Ninety "
                Case Else
            End Select

            ' Convert ones place digit.
            Result = Result & ConvertDigit(Right(MyTens, 1))
        End If

        ConvertTens = Result
    End Function

    Private Function ConvertDigit(ByVal MyDigit As String) As String
        Select Case Val(MyDigit)
            Case 1 : ConvertDigit = "One"
            Case 2 : ConvertDigit = "Two"
            Case 3 : ConvertDigit = "Three"
            Case 4 : ConvertDigit = "Four"
            Case 5 : ConvertDigit = "Five"
            Case 6 : ConvertDigit = "Six"
            Case 7 : ConvertDigit = "Seven"
            Case 8 : ConvertDigit = "Eight"
            Case 9 : ConvertDigit = "Nine"
            Case Else : ConvertDigit = ""
        End Select
    End Function
#End Region

    Public Sub AddLinkTableField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, ByVal LinkedTableName As String)
        Try

            Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
            objUserFieldMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            Try

                If (Not isColumnExist(TableName, ColumnName)) Then
                    objMain.objApplication.StatusBar.SetText("Creating field...[" & ColumnName & "] of table [" & TableName & "]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objUserFieldMD.TableName = TableName
                    objUserFieldMD.Name = ColumnName
                    objUserFieldMD.Description = ColDescription
                    objUserFieldMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
                    objUserFieldMD.LinkedTable = LinkedTableName

                    If (objUserFieldMD.Add() <> 0) Then
                        objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
                GC.Collect()
            End Try
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub AddElements(ByVal FormUID As String, ByVal TableName As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oItem = objForm.Items.Add("lbl", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 5
            oItem.Left = objForm.Items.Item("38").Left
            oItem.Width = objForm.Items.Item("3").Width
            oItem.Height = 15
            Dim lbl As SAPbouiCOM.StaticText = objForm.Items.Item("lbl").Specific
            lbl.Caption = "Total Quantity"

            oItem = objForm.Items.Add("TtlQty", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 5
            oItem.Left = objForm.Items.Item("38").Left + objForm.Items.Item("lbl").Width
            oItem.Width = objForm.Items.Item("3").Width
            oItem.Height = 15

            oItem = objForm.Items.Add("Src", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 5
            oItem.Left = objForm.Items.Item("38").Left + objForm.Items.Item("TtlQty").Width + objForm.Items.Item("lbl").Width + 4
            oItem.Width = 16
            oItem.Height = 15

            Dim btn As SAPbouiCOM.Button = objForm.Items.Item("Src").Specific
            btn.Caption = "..."

            Dim oEditText As SAPbouiCOM.EditText = objForm.Items.Item("TtlQty").Specific
            oEditText.DataBind.SetBound(True, TableName, "U_QUANTITY")
            objForm.Items.Item("TtlQty").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("TtlQty").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.DataSources.UserDataSources.Add("U_DflWhs", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            objForm.DataSources.UserDataSources.Add("U_Branch", SAPbouiCOM.BoDataType.dt_LONG_TEXT)

            oItem = objForm.Items.Add("lblDflWhs", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 22
            oItem.Left = objForm.Items.Item("38").Left
            oItem.Width = objForm.Items.Item("3").Width
            oItem.Height = 15
            Dim lbl1 As SAPbouiCOM.StaticText = objForm.Items.Item("lblDflWhs").Specific
            lbl1.Caption = "Warehouse"

            oItem = objForm.Items.Add("edtBranch", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oItem.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 22
            oItem.Left = objForm.Items.Item("38").Left + objForm.Items.Item("lblDflWhs").Width
            oItem.Width = objForm.Items.Item("4").Width
            oItem.Height = 15
            Dim oComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("edtBranch").Specific
            oComboBox.DataBind.SetBound(True, "", "U_Branch")
            objMain.objUtilities.ComboBoxLoadValues(oComboBox, "Select ""BPLId"" ,  ""BPLName"" From OBPL")
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            oItem.DisplayDesc = True

            oItem = objForm.Items.Add("edtDflWhs", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oItem.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 22
            oItem.Left = objForm.Items.Item("edtBranch").Left + objForm.Items.Item("edtBranch").Width + 5
            oItem.Width = objForm.Items.Item("4").Width
            oItem.Height = 15
            oComboBox = objForm.Items.Item("edtDflWhs").Specific
            oComboBox.DataBind.SetBound(True, "", "U_DflWhs")
            objMain.objUtilities.ComboBoxLoadValues(oComboBox, "Select ""WhsCode"" , ""WhsName"" From ""OWHS"" ")
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            oItem.DisplayDesc = True

            oItem = objForm.Items.Add("Src1", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 22
            oItem.Left = objForm.Items.Item("edtDflWhs").Left + objForm.Items.Item("edtDflWhs").Width + 2
            oItem.Width = 16
            oItem.LinkTo = "edtDflWhs"
            oItem.Height = 15
            Dim btn1 As SAPbouiCOM.Button = objForm.Items.Item("Src1").Specific
            btn1.Caption = "..."
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Function GetBranch(ByVal WhsCode As String) As String
        Try
            Dim BranchDtl As String = "Select ""BPLId"" , ""BPLName"" ,  ""BPLFrName"" , ""DfltResWhs"" From OBPL Where ""DfltResWhs"" = '" & WhsCode & "' "
            Dim oRsBranchDtl As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsBranchDtl.DoQuery(BranchDtl)
            If oRsBranchDtl.RecordCount > 0 Then
                Return oRsBranchDtl.Fields.Item("BPLFrName").Value.ToString
            Else
                Return Nothing
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return Nothing
        End Try
    End Function


    Sub CreateMySimpleForm1()
        Try
            'Dim sIssueSupplier As String = oDBDSHeader.GetValue("CardCode", 0).Trim
            'If sIssueSupplier.Equals("") Then
            '    oCBF.StatusBarErrorMsg("Supplier Code should be left empty ....")
            '    Exit Try
            'End If
            'Dim sIssueProject As String = oDBDSHeader.GetValue("Project", 0).Trim
            'If sIssueProject.Equals("") Then
            '    oCBF.StatusBarErrorMsg("Project Code should be left empty ....")
            '    Exit Try
            'End If
            Dim CP As SAPbouiCOM.FormCreationParams = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)

            Dim oItem As SAPbouiCOM.Item
            Dim oGrid As SAPbouiCOM.Grid
            Dim oButton As SAPbouiCOM.Button
            Dim ocheckbox As SAPbouiCOM.CheckBox

            CP.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            CP.UniqueID = "SOPOMRDetails"
            CP.FormType = "200"

            objForm = objMain.objApplication.Forms.AddEx(CP)
            ' Set form width and height 
            objForm.Height = 300
            objForm.Width = 600
            objForm.Title = "PO ItemList"
            '' Add a Grid item to the form 
            oItem = objForm.Items.Add("MyGrid", SAPbouiCOM.BoFormItemTypes.it_GRID)
            ' Set the grid dimentions and position 
            oItem.Left = 20
            oItem.Top = 20
            oItem.Width = 550
            oItem.Height = 200
            ' Set the grid data 
            oGrid = oItem.Specific
            objForm.DataSources.DataTables.Add("MyDataTable")
            ' Add OK Button 
            oItem = objForm.Items.Add("ok", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 20
            oItem.Top = 230
            oItem.Width = 65
            oItem.Height = 20
            oButton = oItem.Specific
            oButton.Caption = "Ok"
            ' Add CANCEL Button 
            oItem = objForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 90
            oItem.Top = 230
            oItem.Width = 65
            oItem.Height = 20
            oButton = oItem.Specific
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            objForm.DataSources.UserDataSources.Add("Accept", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oItem = objForm.Items.Add("c_select", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 20
            oItem.Top = 5
            oItem.Width = 120
            oItem.Height = 15
            oItem.Enabled = True
            oItem.Visible = True
            oItem.DisplayDesc = False
            oItem.AffectsFormMode = False
            ocheckbox = oItem.Specific
            ocheckbox.Caption = "Select All"
            ocheckbox.ValOn = "Y"
            ocheckbox.ValOff = "N"
            ocheckbox.DataBind.SetBound(True, "", "Accept")
            objForm.Items.Item("c_select").Visible = True

            Dim sQuery As String = " Select 'N' Selected, a.Code,c.ItemName,a.Quantity,b.Project from ITT1 a inner join OITT B on a.Father= b.Code " &
                    " inner join OITM c on a.Code=c.ItemCode where a.U_MakTyp='P' and B.U_Stages='5' and " &
                    " b.Project='" & objForm.Items.Item("157").Specific.value & "' and a.code not in " &
                    " (Select a.ItemCode from POR1 a inner join OPOR b on a.Docentry=b.Docentry " &
                    " where b.Project='" & objForm.Items.Item("157").Specific.value & "' and b.Canceled='N' And b.DocType <>'S') "
            objForm.DataSources.DataTables.Item(0).ExecuteQuery(sQuery)

            oGrid.DataTable = objForm.DataSources.DataTables.Item("MyDataTable")
            oGrid.AutoResizeColumns()
            oGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(2).Editable = False
            oGrid.Columns.Item(3).Editable = True
            oGrid.Columns.Item(4).Editable = False
            objForm.Visible = True
        Catch ex As Exception
            objForm.Close()
            objMain.objApplication.StatusBar.SetText("Item Details Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    'SQl Script
    Public Function ExecuteSQLScript(company As SAPbobsCOM.Company, sqlScript As String) As Boolean
        Try
            Dim rs As SAPbobsCOM.Recordset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            rs.DoQuery(sqlScript)


            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
            rs = Nothing
            GC.Collect()
            Return True
        Catch ex As Exception
            ' Log the error or handle accordingly
            'MsgBox("Error executing SQL: " & ex.Message)
            Return False
        End Try
    End Function


    Public Function LoadEmbeddedSQL(resourceName As String) As String
        Dim assembly = System.Reflection.Assembly.GetExecutingAssembly()
        Dim sExecutingAssemblyNmae As String = assembly.GetName().Name.ToString()
        Dim xmldoc As New Xml.XmlDocument
        resourceName = sExecutingAssemblyNmae + "." + resourceName

        Using stream = assembly.GetManifestResourceStream(resourceName)
            If stream Is Nothing Then Throw New Exception("Embedded SQL resource not found: " & resourceName)
            Using reader As New IO.StreamReader(stream)
                Return reader.ReadToEnd()
            End Using
        End Using
    End Function
End Class


Public Enum ResourceType
    Embeded
    Content
End Enum




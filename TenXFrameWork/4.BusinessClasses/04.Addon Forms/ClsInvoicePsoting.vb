Imports SAPbobsCOM
Imports System
Imports System.IO
Imports System.Text
Imports System.Net.Http
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates

Imports ZXing
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Xml.XmlReader
Imports System.Xml.XmlWriter

Imports System.Xml.Serialization

Imports Newtonsoft.Json
Imports TenXFrameWork.Zatca_CSID.Models
Imports TenXFrameWork.Zatca.Models

Imports Org.BouncyCastle.Math.EC
Imports FormatException = System.FormatException
Imports System.Security.Cryptography.Xml
Imports Org.BouncyCastle.Crypto
Imports Org.BouncyCastle.Security
Imports Org.BouncyCastle.Crypto.Signers
Imports Org.BouncyCastle.Crypto.Parameters
Imports Org.BouncyCastle.OpenSsl
Imports TenXFrameWork.FlickINVModel
Imports System.Net.Http.Headers

Public Class ClsInvoicePsoting
    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix

    Dim objutilities As Utilities
    Public rs, RsNum As SAPbobsCOM.Recordset

    'Form Creation
    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("InvoicingConfiguration.xml", "CONFT", ResourceType.Embeded) 'Loading Form.xml file
            objForm = objMain.objApplication.Forms.GetForm("CONFT", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_INVF")
            Me.Load()
            FillCompanyDBCombo()
            Dim label As SAPbouiCOM.StaticText = CType(objForm.Items.Item("TEST").Specific, SAPbouiCOM.StaticText)
            label.Caption = "𝐓𝐄𝐒𝐓 :"
            Dim label1 As SAPbouiCOM.StaticText = CType(objForm.Items.Item("LIVE").Specific, SAPbouiCOM.StaticText)
            label1.Caption = "𝐋𝐈𝐕𝐄 :"
            Dim label2 As SAPbouiCOM.StaticText = CType(objForm.Items.Item("TEST1").Specific, SAPbouiCOM.StaticText)
            label2.Caption = "𝐓𝐄𝐒𝐓 :"
            Dim label3 As SAPbouiCOM.StaticText = CType(objForm.Items.Item("LIVE1").Specific, SAPbouiCOM.StaticText)
            label3.Caption = "𝐋𝐈𝐕𝐄 :"
            Dim label4 As SAPbouiCOM.StaticText = CType(objForm.Items.Item("B2C").Specific, SAPbouiCOM.StaticText)
            label4.Caption = "𝐁𝟐𝐂 :"
            Dim label5 As SAPbouiCOM.StaticText = CType(objForm.Items.Item("B2B").Specific, SAPbouiCOM.StaticText)
            label5.Caption = "𝐁𝟐𝐁 :"
            Dim label6 As SAPbouiCOM.StaticText = CType(objForm.Items.Item("LB2C").Specific, SAPbouiCOM.StaticText)
            label6.Caption = "𝐁𝟐𝐂 :"
            Dim label7 As SAPbouiCOM.StaticText = CType(objForm.Items.Item("TB2B").Specific, SAPbouiCOM.StaticText)
            label7.Caption = "𝐁𝟐𝐁 :"

            Me.DisableCountryTabs()

            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub FillCompanyDBCombo()

        Try
            Dim oCombo As SAPbouiCOM.ComboBox
            Dim compList As SAPbobsCOM.Recordset

            oCombo = objForm.Items.Item("DBNE").Specific

            'Clear values
            For i As Integer = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
            Next

            'Get company list
            compList = objMain.objCompany.GetCompanyList()

            If compList.RecordCount > 0 Then

                compList.MoveFirst()

                While compList.EoF = False

                    oCombo.ValidValues.Add(
                    compList.Fields.Item(0).Value.ToString(),
                    compList.Fields.Item(1).Value.ToString())

                    compList.MoveNext()

                End While

                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

            End If

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
            ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
    Private Sub DisableCountryTabs()

        Try
            Dim rs As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            objForm.Freeze(True)

            rs.DoQuery("SELECT TOP 1 ""Country"" FROM OADM")

            If rs.EoF Then Exit Sub

            Dim country As String =
        rs.Fields.Item("Country").Value.ToString().Trim().ToUpper()

            '====================================
            ' SAUDI ARABIA
            '====================================
            If country = "SA" OrElse country.Contains("SAUDI") Then

                objForm.PaneLevel = 1

                Try
                    objForm.Items.Item("Item_0").Enabled = True
                Catch
                End Try

                Try
                    objForm.Items.Item("UAEB").Enabled = False
                Catch
                End Try

                '====================================
                ' UAE
                '====================================
            ElseIf country = "AE" OrElse country.Contains("EMIRATES") Then

                objForm.PaneLevel = 2

                Try
                    objForm.Items.Item("UAEB").Enabled = True
                Catch
                End Try

                Try
                    objForm.Items.Item("Item_0").Enabled = False
                Catch
                End Try

            End If

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
        ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        Finally
            objForm.Freeze(False)
        End Try

    End Sub

    'Private Sub DisableCountryTabs()

    '    Try

    '        Dim rs As SAPbobsCOM.Recordset =
    '        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        '✔ Correct query
    '        rs.DoQuery("SELECT TOP 1 ""Country"" FROM OADM")

    '        '✔ Proper check
    '        If rs.EoF Then Exit Sub

    '        rs.MoveFirst()

    '        Dim country As String = ""
    '        country = rs.Fields.Item("Country").Value.ToString().Trim().ToUpper()

    '        'Debug output
    '        objMain.objApplication.StatusBar.SetText(
    '        "Country = " & country,
    '        SAPbouiCOM.BoMessageTime.bmt_Short,
    '        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

    '        '=============================
    '        ' SAUDI ARABIA LOGIC
    '        '=============================
    '        If country = "SA" Or country.Contains("SA") Then

    '            objForm.PaneLevel = 1

    '            Try
    '                objForm.Items.Item("Item_0").Enabled = True   'SA Tab
    '            Catch
    '            End Try

    '            Try
    '                objForm.Items.Item("UAEB").Enabled = False    'UAE Tab
    '            Catch
    '            End Try

    '            '=============================
    '            ' UAE LOGIC
    '            '=============================
    '        ElseIf country = "AE" Then

    '            objForm.PaneLevel = 2

    '            Try
    '                objForm.Items.Item("UAEB").Enabled = True
    '            Catch
    '            End Try

    '            Try
    '                objForm.Items.Item("Item_0").Enabled = False
    '            Catch
    '            End Try

    '        End If

    '    Catch ex As Exception

    '        objMain.objApplication.StatusBar.SetText(
    '        ex.Message,
    '        SAPbouiCOM.BoMessageTime.bmt_Short,
    '        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '    End Try

    'End Sub



    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "Invoice_Posting" And pVal.BeforeAction = False Then
                Me.CreateForm()
            End If

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub Load()
        Try
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_INVF")

            Dim sql As String =
"SELECT TOP 1 " &
"""Code"", " &
"""Name"", " &
"""DocEntry"", " &
"""U_IPPRL"", " &
"""U_IQRCG"", " &
"""U_UUIDB"", " &
"""U_CMPRL"", " &
"""U_CMQRC"", " &
"""U_CMURL"", " &
"""U_BPID"", " &
"""U_APIK"", " &
"""U_IPURLB"", " &
"""U_IQRCGB"", " &
"""U_CMPRLB"", " &
"""U_CMQRB"", " &
"""U_CMURLB"", " &
"""U_BPIDB"", " &
"""U_APIKB"", " &
"""U_TYPE"", " &
"""U_TIPPRL"", " &
"""U_TIQRCG"", " &
"""U_TUUIDB"", " &
"""U_TCMPRL"", " &
"""U_TCMQRC"", " &
"""U_TCMURL"", " &
"""U_TBPID"", " &
"""U_TAPIK"", " &
"""U_TIPURLB"", " &
"""U_TIQRCGB"", " &
"""U_TCMPRLB"", " &
"""U_TCMQRB"", " &
"""U_TCMURLB"", " &
"""U_TBPIDB"", " &
"""U_TAPIKB"", " &
"""U_LTYPE"", " &
"""U_AIPPRL"", " &
"""U_AIUUID"", " &
"""U_ACMQRC"", " &
"""U_ABPID"", " &
"""U_ALIQRCG"", " &
"""U_ALCMPRL"", " &
"""U_ALCMURL"", " &
"""U_ALAPIK"", " &
"""U_LIPPRL"", " &
"""U_LIUUID"", " &
"""U_LCMQRC"", " &
"""U_LBPID"", " &
"""U_LIQRCG"", " &
"""U_LCMPRL"", " &
"""U_LCMURL"", " &
"""U_LAPIK"", " &
"""U_DBNE"", " &
"""U_UTA"", " &
"""U_GURL"", " &
"""U_BPI"", " &
"""U_IPU"", " &
"""U_CPL"", " &
"""U_USM"", " &
"""U_CMRL"", " &
"""U_APID"", " &
"""U_LUTA"", " &
"""U_LGURL"", " &
"""U_LBPI"", " &
"""U_LIPU"", " &
"""U_LCPL"", " &
"""U_LUSM"", " &
"""U_LCMRL"", " &
"""U_LAPID"" " &
"FROM ""@TNX_INVF"" " &
"ORDER BY ""DocEntry"" DESC"



            Dim rs As SAPbobsCOM.Recordset =
            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            rs.DoQuery(sql)

            If rs.RecordCount > 0 Then
                Dim docEntry As String = rs.Fields.Item("DocEntry").Value.ToString()
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                objForm.Items.Item("DocEntry").Specific.Value = docEntry

                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular) ' Click Find button
                objForm.Freeze(False)
                ' Existing configuration → LOAD & UPDATE
                'oDBs_Head.SetValue("Code", 0, rs.Fields.Item("Code").Value)
                'oDBs_Head.SetValue("Name", 0, rs.Fields.Item("Name").Value)
                'oDBs_Head.SetValue("U_IPPRL", 0, rs.Fields.Item("U_IPPRL").Value)
                'oDBs_Head.SetValue("U_IQRCG", 0, rs.Fields.Item("U_IQRCG").Value)
                'oDBs_Head.SetValue("U_IUUID", 0, rs.Fields.Item("U_IUUID").Value)
                'oDBs_Head.SetValue("U_CMPRL", 0, rs.Fields.Item("U_CMPRL").Value)
                'oDBs_Head.SetValue("U_CMQRC", 0, rs.Fields.Item("U_CMQRC").Value)
                'oDBs_Head.SetValue("U_CMURL", 0, rs.Fields.Item("U_CMURL").Value)
                'oDBs_Head.SetValue("U_BPID", 0, rs.Fields.Item("U_BPID").Value)
                'oDBs_Head.SetValue("U_APIK", 0, rs.Fields.Item("U_APIK").Value)
                'oDBs_Head.SetValue("DocEntry", 0, rs.Fields.Item("DocEntry").Value)

                ' objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            Else
                ' First time → ADD MODE
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                oDBs_Head.SetValue("Code", 0, "1")
                'oDBs_Head.SetValue("Name", 0, "InvoiceConfig")
                oDBs_Head.SetValue("DocEntry", 0, 1)
                oDBs_Head.SetValue("U_IPPRL", oDBs_Head.Offset, rs.Fields.Item("U_IPPRL").Value)
                'oDBs_Head.SetValue("U_CTY", oDBs_Head.Offset, rs.Fields.Item("U_CTY").Value)
                oDBs_Head.SetValue("U_IQRCG", oDBs_Head.Offset, rs.Fields.Item("U_IQRCG").Value)
                oDBs_Head.SetValue("U_UUIDB", oDBs_Head.Offset, rs.Fields.Item("U_UUIDB").Value)
                oDBs_Head.SetValue("U_CMPRL", oDBs_Head.Offset, rs.Fields.Item("U_CMPRL").Value)
                oDBs_Head.SetValue("U_CMQRC", oDBs_Head.Offset, rs.Fields.Item("U_CMQRC").Value)
                oDBs_Head.SetValue("U_CMURL", oDBs_Head.Offset, rs.Fields.Item("U_CMURL").Value)
                oDBs_Head.SetValue("U_BPID", oDBs_Head.Offset, rs.Fields.Item("U_BPID").Value)
                oDBs_Head.SetValue("U_APIK", oDBs_Head.Offset, rs.Fields.Item("U_APIK").Value)
                'B2B
                oDBs_Head.SetValue("U_IPURLB", oDBs_Head.Offset, rs.Fields.Item("U_IPURLB").Value)
                oDBs_Head.SetValue("U_IQRCGB", oDBs_Head.Offset, rs.Fields.Item("U_IQRCGB").Value)
                oDBs_Head.SetValue("U_UUIDB", oDBs_Head.Offset, rs.Fields.Item("U_UUIDB").Value)
                oDBs_Head.SetValue("U_CMPRLB", oDBs_Head.Offset, rs.Fields.Item("U_CMPRLB").Value)
                oDBs_Head.SetValue("U_CMQRB", oDBs_Head.Offset, rs.Fields.Item("U_CMQRB").Value)
                oDBs_Head.SetValue("U_CMURLB", oDBs_Head.Offset, rs.Fields.Item("U_CMURLB").Value)
                oDBs_Head.SetValue("U_BPIDB", oDBs_Head.Offset, rs.Fields.Item("U_BPIDB").Value)
                oDBs_Head.SetValue("U_APIKB", oDBs_Head.Offset, rs.Fields.Item("U_APIKB").Value)
                oDBs_Head.SetValue("U_TYPE", oDBs_Head.Offset, rs.Fields.Item("U_TYPE").Value)

                oDBs_Head.SetValue("U_TIPPRL", oDBs_Head.Offset, rs.Fields.Item("U_TIPPRL").Value)
                oDBs_Head.SetValue("U_TIQRCG", oDBs_Head.Offset, rs.Fields.Item("U_TIQRCG").Value)
                oDBs_Head.SetValue("U_TUUIDB", oDBs_Head.Offset, rs.Fields.Item("U_TUUIDB").Value)
                oDBs_Head.SetValue("U_TCMPRL", oDBs_Head.Offset, rs.Fields.Item("U_TCMPRL").Value)
                oDBs_Head.SetValue("U_TCMQRC", oDBs_Head.Offset, rs.Fields.Item("U_TCMQRC").Value)
                oDBs_Head.SetValue("U_TCMURL", oDBs_Head.Offset, rs.Fields.Item("U_TCMURL").Value)
                oDBs_Head.SetValue("U_TBPID", oDBs_Head.Offset, rs.Fields.Item("U_TBPID").Value)
                oDBs_Head.SetValue("U_TAPIK", oDBs_Head.Offset, rs.Fields.Item("U_TAPIK").Value)
                'B2B
                oDBs_Head.SetValue("U_TIPURLB", oDBs_Head.Offset, rs.Fields.Item("U_TIPURLB").Value)
                oDBs_Head.SetValue("U_TIQRCGB", oDBs_Head.Offset, rs.Fields.Item("U_TIQRCGB").Value)
                oDBs_Head.SetValue("U_TUUIDB", oDBs_Head.Offset, rs.Fields.Item("U_TUUIDB").Value)
                oDBs_Head.SetValue("U_TCMPRLB", oDBs_Head.Offset, rs.Fields.Item("U_TCMPRLB").Value)
                oDBs_Head.SetValue("U_TCMQRB", oDBs_Head.Offset, rs.Fields.Item("U_TCMQRB").Value)
                oDBs_Head.SetValue("U_TCMURLB", oDBs_Head.Offset, rs.Fields.Item("U_TCMURLB").Value)
                oDBs_Head.SetValue("U_TBPIDB", oDBs_Head.Offset, rs.Fields.Item("U_TBPIDB").Value)
                oDBs_Head.SetValue("U_TAPIKB", oDBs_Head.Offset, rs.Fields.Item("U_TAPIKB").Value)
                oDBs_Head.SetValue("U_LTYPE", oDBs_Head.Offset, rs.Fields.Item("U_LTYPE").Value)

                oDBs_Head.SetValue("U_AIPPRL", oDBs_Head.Offset, rs.Fields.Item("U_AIPPRL").Value)
                oDBs_Head.SetValue("U_AIUUID", oDBs_Head.Offset, rs.Fields.Item("U_AIUUID").Value)
                oDBs_Head.SetValue("U_ACMQRC", oDBs_Head.Offset, rs.Fields.Item("U_ACMQRC").Value)
                oDBs_Head.SetValue("U_ABPID", oDBs_Head.Offset, rs.Fields.Item("U_ABPID").Value)
                oDBs_Head.SetValue("U_ALIQRCG", oDBs_Head.Offset, rs.Fields.Item("U_ALIQRCG").Value)
                oDBs_Head.SetValue("U_ALCMPRL", oDBs_Head.Offset, rs.Fields.Item("U_ALCMPRL").Value)
                oDBs_Head.SetValue("U_ALCMURL", oDBs_Head.Offset, rs.Fields.Item("U_ALCMURL").Value)
                oDBs_Head.SetValue("U_ALAPIK", oDBs_Head.Offset, rs.Fields.Item("U_ALAPIK").Value)

                oDBs_Head.SetValue("U_LIPPRL", oDBs_Head.Offset, rs.Fields.Item("U_LIPPRL").Value)
                oDBs_Head.SetValue("U_LIUUID", oDBs_Head.Offset, rs.Fields.Item("U_LIUUID").Value)
                oDBs_Head.SetValue("U_LCMQRC", oDBs_Head.Offset, rs.Fields.Item("U_LCMQRC").Value)
                oDBs_Head.SetValue("U_LBPID", oDBs_Head.Offset, rs.Fields.Item("U_LBPID").Value)
                oDBs_Head.SetValue("U_LIQRCG", oDBs_Head.Offset, rs.Fields.Item("U_LIQRCG").Value)
                oDBs_Head.SetValue("U_LCMPRL", oDBs_Head.Offset, rs.Fields.Item("U_LCMPRL").Value)
                oDBs_Head.SetValue("U_LCMURL", oDBs_Head.Offset, rs.Fields.Item("U_LCMURL").Value)
                oDBs_Head.SetValue("U_LAPIK", oDBs_Head.Offset, rs.Fields.Item("U_LAPIK").Value)

                oDBs_Head.SetValue("U_UTA", oDBs_Head.Offset, rs.Fields.Item("U_UTA").Value)
                oDBs_Head.SetValue("U_GURL", oDBs_Head.Offset, rs.Fields.Item("U_GURL").Value)
                oDBs_Head.SetValue("U_BPI", oDBs_Head.Offset, rs.Fields.Item("U_BPI").Value)
                oDBs_Head.SetValue("U_IPU", oDBs_Head.Offset, rs.Fields.Item("U_IPU").Value)
                oDBs_Head.SetValue("U_CPL", oDBs_Head.Offset, rs.Fields.Item("U_CPL").Value)
                oDBs_Head.SetValue("U_USM", oDBs_Head.Offset, rs.Fields.Item("U_USM").Value)
                oDBs_Head.SetValue("U_CMRL", oDBs_Head.Offset, rs.Fields.Item("U_CMRL").Value)
                oDBs_Head.SetValue("U_APID", oDBs_Head.Offset, rs.Fields.Item("U_APID").Value)

                oDBs_Head.SetValue("U_LUTA", oDBs_Head.Offset, rs.Fields.Item("U_LUTA").Value)
                oDBs_Head.SetValue("U_LGURL", oDBs_Head.Offset, rs.Fields.Item("U_LGURL").Value)
                oDBs_Head.SetValue("U_LBPI", oDBs_Head.Offset, rs.Fields.Item("U_LBPI").Value)
                oDBs_Head.SetValue("U_LIPU", oDBs_Head.Offset, rs.Fields.Item("U_LIPU").Value)
                oDBs_Head.SetValue("U_LCPL", oDBs_Head.Offset, rs.Fields.Item("U_LCPL").Value)
                oDBs_Head.SetValue("U_LUSM", oDBs_Head.Offset, rs.Fields.Item("U_LUSM").Value)
                oDBs_Head.SetValue("U_LCMRL", oDBs_Head.Offset, rs.Fields.Item("U_LCMRL").Value)
                oDBs_Head.SetValue("U_LAPID", oDBs_Head.Offset, rs.Fields.Item("U_LAPID").Value)

                oDBs_Head.SetValue("U_DBNE", oDBs_Head.Offset, rs.Fields.Item("U_DBNE").Value)
                oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, rs.Fields.Item("DocEntry").Value)

            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

  





End Class

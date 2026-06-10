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
Imports Zatca.EInvoice.SDK
Imports Zatca.EInvoice.SDK.Contracts.Models
Imports Newtonsoft.Json
Imports TenXFrameWork.Zatca_CSID.Models
Imports TenXFrameWork.Zatca.Models
Imports uuids
Imports Org.BouncyCastle.Math.EC
Imports ECCurve = System.Security.Cryptography.ECCurve
Imports FormatException = System.FormatException
Imports System.Security.Cryptography.Xml
Imports Org.BouncyCastle.Crypto
Imports Org.BouncyCastle.Security
Imports Org.BouncyCastle.Crypto.Signers
Imports Org.BouncyCastle.Crypto.Parameters
Imports Org.BouncyCastle.OpenSsl
Imports TenXFrameWork.FlickINVModel
Imports System.Net.Http.Headers

Public Class clsARDownPayment

#Region "        Declaration        "
    Dim objForm, AltForm, objMIForm As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim SplitCombo As SAPbouiCOM.ComboBox
    Dim SplitFlgCombo As SAPbouiCOM.ComboBox
    '  Public TLV As TLV

    Private EditText0 As SAPbouiCOM.EditText

    Private EditText1 As SAPbouiCOM.EditText

#End Region

#Region "Declaration"
    Dim oitem As SAPbouiCOM.Item
#End Region

    Public Function GetFormattedTime(t As String) As List(Of String)
        Dim lstTime As New List(Of String)()

        If t.Length < 6 Then
            Dim t2 = "0" & t
            lstTime.Add(Convert.ToString(t2(0)) & Convert.ToString(t2(1)))
            lstTime.Add(Convert.ToString(t2(2)) & Convert.ToString(t2(3)))
            lstTime.Add(Convert.ToString(t2(4)) & Convert.ToString(t2(5)))
        Else
            Dim t1 = t
            lstTime.Add(Convert.ToString(t1(0)) & Convert.ToString(t1(1)))
            lstTime.Add(Convert.ToString(t1(2)) & Convert.ToString(t1(3)))
            lstTime.Add(Convert.ToString(t1(4)) & Convert.ToString(t1(5)))
        End If

        Return lstTime
    End Function

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix = objForm.Items.Item("38").Specific

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix = objForm.Items.Item("38").Specific
                    If pVal.ItemUID = "Split" And pVal.BeforeAction = True And objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Dim sErrMsg As String = Nothing
                        Dim lErrCode As Integer = 0

                        Try

                            Dim recinvoice As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(300), SAPbobsCOM.Recordset)
                            Dim DocNum As String = objForm.Items.Item("8").Specific.Value.ToString()
                            Dim Series As String = objForm.Items.Item("88").Specific.Value.ToString()
                            Dim Qry As String = "SELECT ""DocEntry"" FROM ODPI T0 WHERE T0.""DocNum"" = '" & DocNum & "' and T0.""Series""='" & Series & "'"
                            recinvoice.DoQuery(Qry)

                            Dim DocEntry As String = recinvoice.Fields.Item(0).Value().ToString()
                            Me.FlickINVPosting(DocEntry)


                            Exit Sub
                        Catch exception As System.Exception
                            Dim ex As System.Exception = exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        End Try

                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix = objForm.Items.Item("38").Specific
                    If pVal.BeforeAction = False Then

                        Me.AddItems(objForm.UniqueID)
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("65300", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Try
                            Dim GetInv As String = " Select Max(""DocEntry"")  From ODPI"
                            Dim oRsGetInv As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetInv.DoQuery(GetInv)

                            If oRsGetInv.RecordCount > 0 Then
                                Dim DocEntry As String = oRsGetInv.Fields.Item(0).Value
                                Me.FlickINVPosting(DocEntry)
                            End If
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Try
                            Dim GetInv As String = " Select ""U_STATUS""  From ODPI Where ""DocNum""='" & objForm.Items.Item("8").Specific.Value & "' AND ""Series""='" & objForm.Items.Item("88").Specific.Value & "'"
                            Dim oRsGetInv As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetInv.DoQuery(GetInv)
                            If oRsGetInv.RecordCount > 0 Then
                                If oRsGetInv.Fields.Item(0).Value = "success" Then
                                    objForm.Items.Item("Split").Enabled = False
                                End If
                            End If
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub AddItems(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim objMatrix As SAPbouiCOM.Matrix
            objMatrix = objForm.Items.Item("38").Specific

            oitem = objForm.Items.Add("Split", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oitem.Top = objForm.Items.Item("2").Top
            oitem.Left = objForm.Items.Item("2").Left + objForm.Items.Item("2").Width + 5
            oitem.Width = objForm.Items.Item("2").Width + 40
            oitem.Height = objForm.Items.Item("2").Height
            oitem.FromPane = objForm.Items.Item("2").FromPane
            oitem.ToPane = objForm.Items.Item("2").ToPane
            Dim btn As SAPbouiCOM.Button = objForm.Items.Item("Split").Specific
            btn.Caption = "ZATCA Posting"


            objMain.objUtilities.AddLabel(objForm.UniqueID, "Sts", objForm.Items.Item("86").Top + objForm.Items.Item("86").Height + 5, objForm.Items.Item("86").Left, objForm.Items.Item("86").Width, "ZATCA Status", "86")
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "U_STATUS", objForm.Items.Item("46").Top + objForm.Items.Item("46").Height + 5, objForm.Items.Item("46").Left, objForm.Items.Item("46").Width, "ODPI", "U_STATUS", "46", objForm.Items.Item("46").Height)


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FlickINVPosting(ByVal INVEntry As String)
        Try
            'Dim invQry As String = "SELECT T0.""DocEntry"",T0.""DocNum"", T0.""CardCode"", T0.""CardName"", T0.""DocDate"" " &
            '   ",IFNULL(TO_CHAR((TO_TIME(LPAD(T0.""DocTime"",4,0))),'HH24:MI:SS'),'0000') As ""DocTime"",T0.""QRCodeSrc"" " &
            '  " , T0.""NumAtCard"", T0.""VatPercent"",T0.""VatSum"", T0.""VatSumFC"",T0.""DocTotal""-T0.""VatSum"" As ""DocTotalBefVat"" " &
            '  " , T0.""DiscSum"", T0.""DocCur"", T0.""VATRegNum"", T1.""LineNum"", T1.""ItemCode"",T0.""DocTotal"", " &
            '  " T1.""Dscription"",T.""FrgnName"", T1.""Quantity"", T1.""Price"", T1.""LineTotal"", T1.""VatGroup"" " &
            '  " ,C.""CardCode"" As ""CusCode"", C.""CardName"" , C.""Address"" As ""CusAddress"", C.""ZipCode"" As ""CusZipCode"" " &
            '  " , C.""LicTradNum"" As ""CusVAT"",C.""Block"" As ""CusBlock"" " &
            '  " ,C.""StreetNo"" As ""CusStreet"", C.""City"" As ""CustCity"", C.""Country"" As ""CusCountry"",C.""VatIdUnCmp"" As ""CusVAT13"" " &
            '  " ,C.""CardFName"",C.""U_StreetAr"" As ""StreetAr"",C.""U_SubDivAr"" As ""SubDivAr"",C.""U_CityAr"" As ""CityAr"" " &
            '  " , T1.""TaxStatus"", T1.""VatSum"" FROM ODPI T0  INNER JOIN INV1 T1 " &
            '  " On T0.""DocEntry"" = T1.""DocEntry"" INNER JOIN OCRD C On C.""CardCode""=T0.""CardCode"" " &
            '  "  INNER JOIN OITM T On T.""ItemCode""=T1.""ItemCode"" WHERE T0.""DocEntry"" = '" & INVEntry & "' "
            Dim invQry As String = "Call ""TNX_ZATCA_ODPI""('" & INVEntry & "')"
            Dim rsInvoice As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            rsInvoice.DoQuery(invQry)

            Dim dtInvoiceDt As Date = rsInvoice.Fields.Item("DocDate").Value
            Dim IssueDate As String = dtInvoiceDt.ToString("yyyy-MM-dd")
            Dim IssueTime As String = rsInvoice.Fields.Item("DocTime").Value
            Dim InvDocNum As String = rsInvoice.Fields.Item("DocNum").Value
            Dim QR As String = rsInvoice.Fields.Item("QRCodeSrc").Value
            Dim InvRef = rsInvoice.Fields.Item("NumAtCard").Value
            Dim CusCode As String = rsInvoice.Fields.Item("CusCode").Value
            Dim CustNameAr As String = rsInvoice.Fields.Item("CardFName").Value

            Dim DocType As String = rsInvoice.Fields.Item("DocType").Value
            Dim InvType As String = rsInvoice.Fields.Item("InvType").Value
            Dim DocCurrency As String = rsInvoice.Fields.Item("DocCur").Value

            If CustNameAr = "" Then
                objMain.objApplication.StatusBar.SetText("Customer Foreign Name should not be empty")
                Exit Sub
            End If
            Dim CusStreetName As String = rsInvoice.Fields.Item("CusStreet").Value
            Dim CusBuilding As String = rsInvoice.Fields.Item("CusBuilding").Value
            If CusBuilding = "" Then
                objMain.objApplication.StatusBar.SetText("Customer Building should not be empty")
                Exit Sub
            End If
            Dim CusCity As String = rsInvoice.Fields.Item("CustCity").Value
            Dim CusCountry As String = rsInvoice.Fields.Item("CusCountry").Value
            Dim CusRegistrationName As String = rsInvoice.Fields.Item("CardName").Value
            Dim CusTaxID As String = rsInvoice.Fields.Item("CusVAT13").Value
            If CusTaxID = "" Then
                objMain.objApplication.StatusBar.SetText("Customer Federal Tax ID should not be empty")
                Exit Sub
            End If
            Dim CusZipCode As String = rsInvoice.Fields.Item("CusZipCode").Value
            If CusZipCode = "" Then
                objMain.objApplication.StatusBar.SetText("Customer Address ZipCode should not be empty")
                Exit Sub
            End If
            Dim CityAr As String = rsInvoice.Fields.Item("CityAr").Value
            If CityAr = "" Then
                objMain.objApplication.StatusBar.SetText("Customer Address Ciry Arabic should not be empty")
                Exit Sub
            End If
            Dim StreetAr As String = rsInvoice.Fields.Item("StreetAr").Value
            If StreetAr = "" Then
                objMain.objApplication.StatusBar.SetText("Customer Address Street Arabic should not be empty")
                Exit Sub
            End If
            Dim SubDivisionAr As String = rsInvoice.Fields.Item("SubDivAr").Value
            If SubDivisionAr = "" Then
                objMain.objApplication.StatusBar.SetText("Customer Address Sub Division Arabic should not be empty")
                Exit Sub
            End If

            'Dim DocTotal As Double = rsInvoice.Fields.Item("DocTotal").Value
            'Dim DocTotalBefVat As Double = rsInvoice.Fields.Item("DocTotalBefVat").Value
            'Dim VatTotal As Double = rsInvoice.Fields.Item("VatSum").Value

            'Dim LineId As Integer = rsInvoice.Fields.Item("LineNum").Value
            'Dim Qty As Double = rsInvoice.Fields.Item("Quantity").Value
            'Dim LineTotal As Double = rsInvoice.Fields.Item("LineTotal").Value
            'Dim vatSum As Double = rsInvoice.Fields.Item("VatSum").Value
            'Dim ItemDescription As String = rsInvoice.Fields.Item("Dscription").Value.ToString()
            'Dim Price As Double = rsInvoice.Fields.Item("Price").Value


            Dim oInv As FlickINVModel.FlickINV = New FlickINVModel.FlickINV
            oInv.invoice_ref_number = InvDocNum
            oInv.issue_date = IssueDate
            oInv.issue_time = IssueTime
            oInv.doc_type = DocType
            oInv.inv_type = InvType
            oInv.currency = DocCurrency

            Dim oParty As Party_Details = New Party_Details
            oParty.party_name_ar = CustNameAr
            oParty.party_name_en = CusRegistrationName
            oParty.party_vat = CusTaxID
            oParty.city_ar = CityAr
            oParty.city_en = CusCity
            oParty.city_subdivision_ar = SubDivisionAr
            oParty.city_subdivision_en = CusStreetName
            oParty.street_ar = StreetAr
            oParty.street_en = CusCountry
            oParty.building = CusBuilding
            oParty.postal_zone = CusZipCode
            oInv.party_details = oParty

            Dim oDeliveryDetails As Delivery_Details = New Delivery_Details
            oDeliveryDetails.actual_delivery = IssueDate
            oDeliveryDetails.latest_delivery = IssueDate
            oInv.delivery_details = oDeliveryDetails

            Dim oNotesDetails As notes_details = New notes_details
            oNotesDetails.related_doc_reference = ""
            oNotesDetails.reason_text = ""
            oInv.notes_details = oNotesDetails

            Dim oLines As IList(Of Lineitem) = New List(Of Lineitem)
            For i As Integer = 1 To rsInvoice.RecordCount
                Dim Lineitem As Lineitem = New Lineitem()
                Dim Quantity As Double = rsInvoice.Fields.Item("Quantity").Value
                Dim Price As Double = rsInvoice.Fields.Item("Price").Value

                If rsInvoice.Fields.Item("FrgnName").Value.ToString() = "" Then
                    objMain.objApplication.StatusBar.SetText("Item Foreign Name should not be empty for " & rsInvoice.Fields.Item("Dscription").Value.ToString())
                    Exit Sub
                End If
                Lineitem.name_ar = rsInvoice.Fields.Item("FrgnName").Value
                Lineitem.name_en = rsInvoice.Fields.Item("Dscription").Value
                Lineitem.quantity = rsInvoice.Fields.Item("Quantity").Value
                Lineitem.tax_category = "S"
                Lineitem.tax_exclusive_price = rsInvoice.Fields.Item("Price").Value
                Lineitem.tax_percentage = rsInvoice.Fields.Item("VatPercent").Value
                oLines.Add(Lineitem)
                rsInvoice.MoveNext()

            Next
            oInv.lineitems = oLines
            Dim jsonString As String = JsonConvert.SerializeObject(oInv)

            '-----Request------
            'Dim url As String = "https://sandbox-sa.flick.network/api/einvoice/generate/invoice"
            Dim url As String = ""
            Try
                Dim response = MakeRequestAsync(url, jsonString).GetAwaiter().GetResult()

                Dim oInvoice As SAPbobsCOM.Documents = DirectCast(objMain.objCompany.GetBusinessObject(203), SAPbobsCOM.Documents)
                oInvoice.GetByKey(Integer.Parse(INVEntry))

                If response.message Is Nothing Then
                    oInvoice.UserFields.Fields.Item("U_MSG").Value = "Posted"
                Else
                    oInvoice.UserFields.Fields.Item("U_MSG").Value = response.message
                End If

                oInvoice.UserFields.Fields.Item("U_INVHASH").Value = response.data.invoice_hash
                oInvoice.UserFields.Fields.Item("U_UUID").Value = response.data.uuid
                oInvoice.UserFields.Fields.Item("U_STATUS").Value = response.status

                oInvoice.CreateQRCodeFrom = (response.data.qr_code)

                Dim lRetCode As Integer = oInvoice.Update()
                If (lRetCode = 0) Then
                    objForm.Refresh()
                    objMain.objApplication.StatusBar.SetText("Posted Suscessfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
                End If

            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText("Error: " & ex.Message)
            End Try
            Dim Test As String = ""
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    ' Define a class for the JSON response

    Public Async Function MakeRequestAsync(url As String, jsonData As String) As Task(Of ApiResponse)
        Using client As New HttpClient()
            ' Add two custom headers
            'client.DefaultRequestHeaders.Add("egs_uuid", "0afab0bf-66ad-47b3-b60b-82b4c0eff8c1")
            'client.DefaultRequestHeaders.Add("x-flick-auth-key", "3776dfaec0a070ab5c1aa89f5bb2616194796612b625b75a36de58089a84da7c")

            ' Set the content type to JSON
            Dim content As New StringContent(jsonData, Encoding.UTF8, "application/json")

            ' Make the POST request
            Dim response As HttpResponseMessage = Await client.PostAsync(url, content)

            ' Ensure success status code
            response.EnsureSuccessStatusCode()

            ' Read and deserialize the JSON response
            Dim jsonResponse As String = Await response.Content.ReadAsStringAsync()
            Dim deserializedResponse As ApiResponse = JsonConvert.DeserializeObject(Of ApiResponse)(jsonResponse)

            Return deserializedResponse
        End Using
    End Function

    Public Class ApiResponse
        Public Property status As String
        Public Property message As String
        Public Property data As Data
    End Class
    Public Class Data
        Public Property uuid As String
        Public Property invoice_hash As String
        Public Property qr_code As String
    End Class



End Class

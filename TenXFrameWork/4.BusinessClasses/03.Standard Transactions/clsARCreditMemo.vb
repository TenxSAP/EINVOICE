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

Public Class clsARCreditMemo

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
                            Dim Qry As String = "SELECT ""DocEntry"" FROM ORIN T0 WHERE T0.""DocNum"" = '" & DocNum & "' and T0.""Series""='" & Series & "'"
                            recinvoice.DoQuery(Qry)

                            Dim DocEntry As String = recinvoice.Fields.Item(0).Value().ToString()
                            Me.E_CreditMemoPosting(DocEntry)


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

            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    objForm = objMain.objApplication.Forms.GetForm("179", objMain.objApplication.Forms.ActiveForm.TypeCount)
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Try
                            Dim GetInv As String = " Select Max(""DocEntry"")  From ORIN"
                            Dim oRsGetInv As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetInv.DoQuery(GetInv)

                            If oRsGetInv.RecordCount > 0 Then
                                Dim DocEntry As String = oRsGetInv.Fields.Item(0).Value
                                Me.E_CreditMemoPosting(DocEntry)
                            End If
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    objForm = objMain.objApplication.Forms.GetForm("179", objMain.objApplication.Forms.ActiveForm.TypeCount)
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Try
                            Dim GetInv As String = " Select ""U_STATUS1""  From ORIN Where ""DocNum""='" & objForm.Items.Item("8").Specific.Value & "' AND ""Series""='" & objForm.Items.Item("88").Specific.Value & "'"
                            Dim oRsGetInv As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetInv.DoQuery(GetInv)
                            If oRsGetInv.RecordCount > 0 Then
                                Dim status As String = oRsGetInv.Fields.Item("U_STATUS1").Value.ToString().Trim()

                                If status <> "" Then
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
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "U_STATUS1", objForm.Items.Item("46").Top + objForm.Items.Item("46").Height + 5, objForm.Items.Item("46").Left, objForm.Items.Item("46").Width, "ORIN", "U_STATUS1", "46", objForm.Items.Item("46").Height)


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub E_CreditMemoPosting(ByVal INVEntry As String)
        Try
            Dim rs As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            Dim sql As String = "SELECT TOP 1 ""U_TYPE"" FROM ""@TNX_INVF"""

            rs.DoQuery(sql)
            If rs.RecordCount = 0 Then
                Throw New Exception("No configuration found in @TNX_INVF")
            End If

            Dim Type As String = rs.Fields.Item("U_TYPE").Value.ToString()

            If Type = "2" Then
                E_CreditMemoPosting_B2C(INVEntry)
            ElseIf Type = "1" Then
                E_CreditMemoPosting_B2B(INVEntry)

            Else
                Throw New Exception("Invalid TYPE in @TNX_INVF (1=B2B, 2=B2C)")
            End If

        Catch ex As Exception
            ErrorLog(ex.Message)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub E_CreditMemoPosting_B2C(ByVal CRNEntry As String)

        Try
            ' Calling Procedure 
            Dim crQry As String = "Call ""TNX_ZATCA_ORIN""('" & CRNEntry & "')"

            Dim rsCredit As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            rsCredit.DoQuery(crQry)

            If rsCredit.RecordCount = 0 Then
                Throw New Exception("No Credit Memo data found")
            End If
            'Header Fields
            Dim CreditNoteDate As String = CType(rsCredit.Fields.Item("credit_note_date").Value, Date).ToString("yyyy-MM-dd")
            Dim CreditNoteNumber As String = rsCredit.Fields.Item("credit_note_number").Value.ToString()
            Dim InvoiceNumber As String = rsCredit.Fields.Item("invoice_number").Value.ToString()
            Dim InvoiceDate As String = CType(rsCredit.Fields.Item("invoice_date").Value, Date).ToString("yyyy-MM-dd")
            Dim Reason As String = rsCredit.Fields.Item("reason_for_issuance").Value.ToString()
            Dim PaymentTerms As String = rsCredit.Fields.Item("payment_terms").Value.ToString()
            'Dim InvRef = rsCredit.Fields.Item("NumAtCard").Value
            'Dim CusCode As String = rsCredit.Fields.Item("CusCode").Value
            Dim CustName As String = rsCredit.Fields.Item("name").Value
            ' Dim notes As String = rsCredit.Fields.Item("notes").Value
            'Dim DocEntry As Integer = rsCredit.Fields.Item("DocEntry").Value
            'Dim DocType As String = rsCredit.Fields.Item("DocType").Value
            'Dim InvType As String = rsCredit.Fields.Item("InvType").Value

            Dim Phone As String = rsCredit.Fields.Item("phone").Value
            Dim Email As String = rsCredit.Fields.Item("email").Value
            Dim Country As String = rsCredit.Fields.Item("country").Value
            Dim CountryCode As String = rsCredit.Fields.Item("country_code").Value
            Dim OtherIDType As String = rsCredit.Fields.Item("other_id_type").Value
            Dim OtherID As String = rsCredit.Fields.Item("other_id").Value

            Dim ItemCode As String = rsCredit.Fields.Item("item_id").Value
            Dim ItemName As String = rsCredit.Fields.Item("item_name").Value
            Dim Quantity As Decimal = Convert.ToDecimal(rsCredit.Fields.Item("quantity").Value)
            Dim UnitRate As Decimal = Convert.ToDecimal(rsCredit.Fields.Item("unit_rate").Value)
            ' Dim LineTotal As Decimal = Convert.ToDecimal(rsCredit.Fields.Item("LineTotal").Value)
            Dim VatCategory As String = rsCredit.Fields.Item("vat_category_code").Value
            Dim VatRate As Decimal = Convert.ToDecimal(rsCredit.Fields.Item("vat_rate").Value)
            Dim Uom As String = rsCredit.Fields.Item("unit_of_measure").Value.ToString()
            Dim PaymentMeans As String = rsCredit.Fields.Item("payment_means").Value.ToString()
            'Dim LogoUrl As String = rsCredit.Fields.Item("logo_url").Value.ToString()
            Dim Type As String = rsCredit.Fields.Item("type").Value.ToString()
            Dim ItemPriceDiscount As Decimal = Convert.ToDecimal(rsCredit.Fields.Item("item_price_discount").Value)
            Dim PriceAllowanceIndicator As String = rsCredit.Fields.Item("price_allowance_indicator").Value.ToString()
            Dim VatExemptionCode As String = rsCredit.Fields.Item("vat_exemption_code").Value.ToString()
            'Dim VatExemptionReason As String = rsCredit.Fields.Item("vat_exemption_reason").Value.ToString()
            ' Dim ServiceCode As String = rsCredit.Fields.Item("service_code").Value.ToString()
            ' Dim HeaderDiscount As Decimal = Convert.ToDecimal(rsCredit.Fields.Item("header_discount").Value)
            Dim InvoiceCurrency As String = rsCredit.Fields.Item("invoice_currency").Value.ToString()
            Dim BusinessKind As String = rsCredit.Fields.Item("business_kind").Value.ToString()

            'Dim DiscountApplicable As String = rsCredit.Fields.Item("is_discount_applicable").Value.ToString()
            '  Credit Note Object
            '---------------------------------------------
            Dim oCRN As New FlickINVModel.E_CreditNoteRequest()

            oCRN.type = rsCredit.Fields.Item("type").Value.ToString()
            oCRN.credit_note_number = CreditNoteNumber
            oCRN.credit_note_date = CreditNoteDate
            oCRN.invoice_number = InvoiceNumber
            oCRN.invoice_date = InvoiceDate
            oCRN.reason_for_issuance = Reason
            oCRN.payment_terms = PaymentTerms
            oCRN.payment_means = PaymentMeans
            oCRN.invoice_currency = InvoiceCurrency
            'oCRN.total_in_words_arabic = rsCredit.Fields.Item("business_kind").Value.ToString()
            oCRN.business_kind = BusinessKind
            oCRN.document_number_type = rsCredit.Fields.Item("document_number_type").Value.ToString()


            ' Billed To
            '---------------------------------------------
            Dim oBilledTo As New FlickINVModel.BilledTo()

            oBilledTo.name = CustName
            oBilledTo.phone = Phone
            oBilledTo.country = Country
            oBilledTo.country_code = CountryCode

            oCRN.billed_to = oBilledTo

            ' Item Details
            '---------------------------------------------
            Dim oItems As New List(Of ItemDetail)

            rsCredit.MoveFirst()
            While Not rsCredit.EoF

                Dim oItem As New ItemDetail()
                oItem.other_id_type = OtherIDType
                oItem.other_id = OtherID
                oItem.item_id = ItemCode
                oItem.item_name = ItemName
                oItem.quantity = Quantity
                oItem.unit_rate = UnitRate
                oItem.unit_of_measure = Uom

                oItem.vat_category_code = VatCategory
                oItem.vat_exemption_code = VatExemptionCode
                oItem.vat_rate = VatRate

                oItem.item_price_discount = ItemPriceDiscount

                oItem.price_allowance_indicator = PriceAllowanceIndicator

                'oItem.vat_exemption_reason = rsCredit.Fields.Item("vat_exemption_reason").Value.ToString()
                'oItem.service_code = rsCredit.Fields.Item("service_code").Value.ToString()
                oItem.is_discount_applicable = (oItem.item_price_discount > 0)


                oItems.Add(oItem)
                rsCredit.MoveNext()

            End While

            oCRN.item_details = oItems

            ' Converting above data into json format
            '---------------------------------------------
            Dim jsonString As String = JsonConvert.SerializeObject(oCRN, Newtonsoft.Json.Formatting.Indented)


            ' Get API Configuration
            '---------------------------------------------
            Dim rsCfg As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            'To Get Configured Data from the Invoicing Posting Screen
            Dim Query As String = "SELECT ""U_IPPRL"", ""U_BPID"", ""U_CMPRL"",""U_CMQRC"",""U_APIK"" FROM ""@TNX_INVF"" ORDER BY ""Code"" LIMIT 1"
            rsCfg.DoQuery(Query)

            Dim baseUrl As String = rsCfg.Fields.Item("U_CMPRL").Value.ToString.Trim    'Credit Memo Url
            Dim bpid As String = rsCfg.Fields.Item("U_BPID").Value.ToString.Trim    'Business Profile Id
            Dim QRGenerationUrl As String = rsCfg.Fields.Item("U_CMQRC").Value.ToString.Trim 'QRGeneration Url
            Dim apiKey As String = rsCfg.Fields.Item("U_APIK").Value.ToString.Trim 'Api Key

            Dim finalUrl As String = baseUrl & "/" & bpid

            'URl Post Result
            Dim apiResult = MakeRequestAsync(finalUrl, jsonString, apiKey).GetAwaiter().GetResult()
            'Fields Updating
            UpdateCreditMemoAfterResponse(CRNEntry, apiResult)

            QrCodeUpdation(CRNEntry, QRGenerationUrl, apiKey)
        Catch ex As Exception
            ErrorLog(ex.Message)
            objMain.objApplication.StatusBar.SetText("Credit Memo Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Sub E_CreditMemoPosting_B2B(ByVal CRNEntry As String)

        Try
            ' Calling Procedure 
            Dim crQry As String = "Call ""TNX_ZATCA_B2B_ORIN""('" & CRNEntry & "')"

            Dim rsCredit As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            rsCredit.DoQuery(crQry)

            If rsCredit.RecordCount = 0 Then
                Throw New Exception("No Credit Memo data found")
            End If
            'Header Fields
            Dim CreditNoteDate As String = CType(rsCredit.Fields.Item("credit_note_date").Value, Date).ToString("yyyy-MM-dd")
            Dim CreditNoteNumber As String = rsCredit.Fields.Item("credit_note_number").Value.ToString()
            Dim InvoiceNumber As String = rsCredit.Fields.Item("invoice_number").Value.ToString()
            Dim InvoiceDate As String = CType(rsCredit.Fields.Item("invoice_date").Value, Date).ToString("yyyy-MM-dd")
            Dim Reason As String = rsCredit.Fields.Item("reason_for_issuance").Value.ToString()
            Dim PaymentTerms As String = rsCredit.Fields.Item("payment_terms").Value.ToString()
            'Dim InvRef = rsCredit.Fields.Item("NumAtCard").Value
            'Dim CusCode As String = rsCredit.Fields.Item("CusCode").Value
            Dim CustName As String = rsCredit.Fields.Item("name").Value
            ' Dim notes As String = rsCredit.Fields.Item("notes").Value
            'Dim DocEntry As Integer = rsCredit.Fields.Item("DocEntry").Value
            'Dim DocType As String = rsCredit.Fields.Item("DocType").Value
            'Dim InvType As String = rsCredit.Fields.Item("InvType").Value

            Dim Phone As String = rsCredit.Fields.Item("phone").Value
            Dim Email As String = rsCredit.Fields.Item("email").Value
            Dim Country As String = rsCredit.Fields.Item("country").Value
            Dim CountryCode As String = rsCredit.Fields.Item("country_code").Value
            Dim OtherIDType As String = rsCredit.Fields.Item("other_id_type").Value
            Dim OtherID As String = rsCredit.Fields.Item("other_id").Value

            Dim ItemCode As String = rsCredit.Fields.Item("item_id").Value
            Dim ItemName As String = rsCredit.Fields.Item("item_name").Value
            Dim Quantity As Decimal = Convert.ToDecimal(rsCredit.Fields.Item("quantity").Value)
            Dim UnitRate As Decimal = Convert.ToDecimal(rsCredit.Fields.Item("unit_rate").Value)
            ' Dim LineTotal As Decimal = Convert.ToDecimal(rsCredit.Fields.Item("LineTotal").Value)
            Dim VatCategory As String = rsCredit.Fields.Item("vat_category_code").Value
            Dim VatRate As Decimal = Convert.ToDecimal(rsCredit.Fields.Item("vat_rate").Value)
            Dim Uom As String = rsCredit.Fields.Item("unit_of_measure").Value.ToString()
            Dim PaymentMeans As String = rsCredit.Fields.Item("payment_means").Value.ToString()
            'Dim LogoUrl As String = rsCredit.Fields.Item("logo_url").Value.ToString()
            Dim Type As String = rsCredit.Fields.Item("type").Value.ToString()
            Dim ItemPriceDiscount As Decimal = Convert.ToDecimal(rsCredit.Fields.Item("item_price_discount").Value)
            Dim PriceAllowanceIndicator As String = rsCredit.Fields.Item("price_allowance_indicator").Value.ToString()
            Dim VatExemptionCode As String = rsCredit.Fields.Item("vat_exemption_code").Value.ToString()
            'Dim VatExemptionReason As String = rsCredit.Fields.Item("vat_exemption_reason").Value.ToString()
            ' Dim ServiceCode As String = rsCredit.Fields.Item("service_code").Value.ToString()
            ' Dim HeaderDiscount As Decimal = Convert.ToDecimal(rsCredit.Fields.Item("header_discount").Value)
            Dim InvoiceCurrency As String = rsCredit.Fields.Item("invoice_currency").Value.ToString()
            Dim BusinessKind As String = rsCredit.Fields.Item("business_kind").Value.ToString()



            Dim name As String = rsCredit.Fields.Item("name").Value.ToString()
            Dim company_name As String = rsCredit.Fields.Item("company_name").Value.ToString()
            Dim logo_url As String = rsCredit.Fields.Item("logo_url").Value.ToString()
            Dim address1 As String = rsCredit.Fields.Item("address1").Value.ToString()
            Dim city As String = rsCredit.Fields.Item("city").Value.ToString()
            Dim district As String = rsCredit.Fields.Item("district").Value.ToString()
            Dim postal_code As String = rsCredit.Fields.Item("postal_code").Value.ToString()
            Dim company_reg_number As String = rsCredit.Fields.Item("company_reg_number").Value.ToString()
            Dim customer_number As String = rsCredit.Fields.Item("customer_number").Value.ToString()

            Dim gst_no As String = rsCredit.Fields.Item("gst_no").Value.ToString()
            Dim grp_vat_no As String = rsCredit.Fields.Item("grp_vat_no").Value.ToString()
            Dim other_id As String = rsCredit.Fields.Item("other_id").Value.ToString()
            Dim other_id_type As String = rsCredit.Fields.Item("other_id_type").Value.ToString()

            Dim building_number As String = rsCredit.Fields.Item("building_number").Value.ToString()
            Dim country_code As String = rsCredit.Fields.Item("country_code").Value.ToString()
            Dim neighbourhood As String = rsCredit.Fields.Item("neighbourhood").Value.ToString()

            Dim company_name_local As String = rsCredit.Fields.Item("company_name_local").Value.ToString()
            Dim iso_certificate_url As String = rsCredit.Fields.Item("iso_certificate_url").Value.ToString()

            Dim profile_type As String = rsCredit.Fields.Item("profile_type").Value.ToString()
            Dim is_saudi_vat_registered As Boolean = Convert.ToBoolean(rsCredit.Fields.Item("is_saudi_vat_registered").Value)

            Dim payment_card_last_four_digits As String = rsCredit.Fields.Item("payment_card_last_four_digits").Value.ToString()
            Dim address_in_local_language As String = rsCredit.Fields.Item("address_in_local_language").Value.ToString()
            'Dim DiscountApplicable As String = rsCredit.Fields.Item("is_discount_applicable").Value.ToString()
            '  Credit Note Object
            '---------------------------------------------
            Dim oCRN As New FlickINVModel.E_CreditNoteRequestB2B()

            oCRN.type = rsCredit.Fields.Item("type").Value.ToString()
            oCRN.credit_note_number = CreditNoteNumber
            oCRN.credit_note_date = CreditNoteDate
            oCRN.invoice_number = InvoiceNumber
            oCRN.invoice_date = InvoiceDate
            oCRN.reason_for_issuance = Reason
            oCRN.payment_terms = PaymentTerms
            oCRN.payment_means = PaymentMeans
            oCRN.invoice_currency = InvoiceCurrency
            'oCRN.total_in_words_arabic = rsCredit.Fields.Item("business_kind").Value.ToString()
            oCRN.business_kind = BusinessKind
            oCRN.document_number_type = rsCredit.Fields.Item("document_number_type").Value.ToString()


            ' Billed To
            '---------------------------------------------
            Dim oBilledTo As New FlickINVModel.BilledToB2B()
            oBilledTo.name = name
            oBilledTo.company_name = company_name
            oBilledTo.logo_url = logo_url
            oBilledTo.email = Email
            oBilledTo.phone = Phone

            oBilledTo.address1 = address1
            oBilledTo.city = city
            oBilledTo.district = district
            oBilledTo.postal_code = postal_code

            oBilledTo.country = Country
            oBilledTo.company_reg_number = company_reg_number
            oBilledTo.customer_number = customer_number

            oBilledTo.gst_no = gst_no
            oBilledTo.grp_vat_no = grp_vat_no
            oBilledTo.other_id = other_id
            oBilledTo.other_id_type = other_id_type

            oBilledTo.building_number = building_number
            oBilledTo.country_code = country_code
            oBilledTo.neighbourhood = neighbourhood

            oBilledTo.company_name_local = company_name_local
            oBilledTo.iso_certificate_url = iso_certificate_url

            oBilledTo.profile_type = profile_type
            oBilledTo.is_saudi_vat_registered = is_saudi_vat_registered

            oBilledTo.payment_card_last_four_digits = payment_card_last_four_digits
            oBilledTo.address_in_local_language = address_in_local_language


            oCRN.billed_to = oBilledTo

            ' Item Details
            '---------------------------------------------
            Dim oItems As New List(Of ItemDetailB2B)

            rsCredit.MoveFirst()
            While Not rsCredit.EoF

                Dim oItem As New ItemDetailB2B()
                oItem.other_id_type = OtherIDType
                oItem.other_id = OtherID
                oItem.item_id = ItemCode
                oItem.item_name = ItemName
                oItem.quantity = Quantity
                oItem.unit_rate = UnitRate
                oItem.unit_of_measure = Uom

                oItem.vat_category_code = VatCategory
                oItem.vat_exemption_code = VatExemptionCode
                oItem.vat_rate = VatRate

                oItem.item_price_discount = ItemPriceDiscount

                oItem.price_allowance_indicator = PriceAllowanceIndicator

                'oItem.vat_exemption_reason = rsCredit.Fields.Item("vat_exemption_reason").Value.ToString()
                'oItem.service_code = rsCredit.Fields.Item("service_code").Value.ToString()
                oItem.is_discount_applicable = (oItem.item_price_discount > 0)


                oItems.Add(oItem)
                rsCredit.MoveNext()

            End While

            oCRN.item_details = oItems

            ' Converting above data into json format
            '---------------------------------------------
            Dim jsonString As String = JsonConvert.SerializeObject(oCRN, Newtonsoft.Json.Formatting.Indented)


            ' Get API Configuration
            '---------------------------------------------
            Dim rsCfg As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            'To Get Configured Data from the Invoicing Posting Screen
            Dim Query As String = "SELECT ""U_IPPRL"", ""U_BPIDB"", ""U_CMPRLB"",""U_CMQRB"",""U_APIKB"" FROM ""@TNX_INVF"" ORDER BY ""Code"" LIMIT 1"
            rsCfg.DoQuery(Query)

            Dim baseUrl As String = rsCfg.Fields.Item("U_CMPRLB").Value.ToString.Trim    'Credit Memo Url
            Dim bpid As String = rsCfg.Fields.Item("U_BPIDB").Value.ToString.Trim    'Business Profile Id
            Dim QRGenerationUrl As String = rsCfg.Fields.Item("U_CMQRB").Value.ToString.Trim 'QRGeneration Url
            Dim apiKey As String = rsCfg.Fields.Item("U_APIKB").Value.ToString.Trim 'Api Key

            Dim finalUrl As String = baseUrl & "/" & bpid

            'URl Post Result
            Dim apiResult = MakeRequestAsync(finalUrl, jsonString, apiKey).GetAwaiter().GetResult()
            'Fields Updating
            UpdateCreditMemoAfterResponse(CRNEntry, apiResult)

            QrCodeUpdation(CRNEntry, QRGenerationUrl, apiKey)
        Catch ex As Exception
            ErrorLog(ex.Message)
            objMain.objApplication.StatusBar.SetText("Credit Memo Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
    Private Sub UpdateCreditMemoAfterResponse(CRNEntry As String, apiResult As ApiCallResult)

        Dim oCredit As SAPbobsCOM.Documents = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes), SAPbobsCOM.Documents)

        oCredit.GetByKey(Integer.Parse(CRNEntry))

        oCredit.UserFields.Fields.Item("U_RESPONSE").Value = apiResult.RawResponse
        oCredit.UserFields.Fields.Item("U_MSG").Value = apiResult.Message
        oCredit.Update()
        If apiResult.IsSuccess AndAlso apiResult.ParsedResponse IsNot Nothing Then

            If apiResult.ParsedResponse.zatca_response IsNot Nothing AndAlso apiResult.ParsedResponse.zatca_response.error = False Then
                oCredit.UserFields.Fields.Item("U_STATUS1").Value = apiResult.ParsedResponse.zatca_response.message
            Else
                oCredit.UserFields.Fields.Item("U_STATUS1").Value = "ZATCA_FAILED"
            End If

            If apiResult.ParsedResponse.details IsNot Nothing Then
                oCredit.UserFields.Fields.Item("U_UUID").Value = apiResult.ParsedResponse.details.document_id
            End If

            oCredit.Update()

            objMain.objApplication.StatusBar.SetText("ZATCA Credit Memo Post Success", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Else
            ErrorLog(apiResult.Message)
            objMain.objApplication.StatusBar.SetText("ZATCA Error: " & apiResult.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If

    End Sub


    Public Sub QrCodeUpdation(ByVal CRNEntry As Integer, ByVal QRGenerationUrl As String, ByVal apiKey As String)

        Try

            Dim oCredit As SAPbobsCOM.Documents = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes), SAPbobsCOM.Documents)

            If Not oCredit.GetByKey(CRNEntry) Then
                objMain.objApplication.StatusBar.SetText("Credit Memo not found", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            'Fetching UUID from the ORIN table
            Dim sUUID As String = oCredit.UserFields.Fields.Item("U_UUID").Value.ToString().Trim()

            If String.IsNullOrEmpty(sUUID) Then
                objMain.objApplication.StatusBar.SetText("UUID not available in Credit Memo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            'Fetching QRGenerationUrl from the Invoice Posting Screen  along with we are passing UUID to get the Response
            Dim getUrl As String = QRGenerationUrl & "/" & Uri.EscapeDataString(sUUID)

            Dim response As GetApiResponse = GetRequestAsync(getUrl, apiKey).GetAwaiter().GetResult()

            If response Is Nothing Then
                objMain.objApplication.StatusBar.SetText("Empty response from ZATCA", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            'Updating XML in HANA -
            If Not String.IsNullOrEmpty(response.xml) Then
                oCredit.UserFields.Fields.Item("U_XML").Value = response.xml
            End If

            'Updating QRCodeField in HANA --QRCodeSrc
            If Not String.IsNullOrEmpty(response.qr) Then
                oCredit.CreateQRCodeFrom = response.qr

            End If


            Dim lRetCode As Integer = oCredit.Update()

            If lRetCode = 0 Then
                objMain.objApplication.StatusBar.SetText("Credit Memo ZATCA XML & QR Fields are updated successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

        Catch ex As Exception
            ErrorLog(ex.Message)
            objMain.objApplication.StatusBar.SetText("Credit Memo ZATCA Update Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub



    Public Async Function MakeRequestAsync(url As String, jsonData As String, apiKey As String) As Task(Of ApiCallResult)

        Using client As New HttpClient()

            client.DefaultRequestHeaders.Clear()
            client.DefaultRequestHeaders.Add("api-key", apiKey)

            Dim content As New StringContent(jsonData, Encoding.UTF8, "application/json")

            Dim httpResponse As HttpResponseMessage =
            Await client.PostAsync(url, content)

            Dim rawJson As String =
            Await httpResponse.Content.ReadAsStringAsync()

            Dim msg As String = "API error occurred"
            Dim parsed As ApiResponse = Nothing
            ' 🔹 Extract "message" from JSON (success or error)
            Try
                parsed = JsonConvert.DeserializeObject(Of ApiResponse)(rawJson)
                If parsed IsNot Nothing AndAlso Not String.IsNullOrEmpty(parsed.message) Then
                    msg = parsed.message
                End If
            Catch
                msg = rawJson
            End Try

            Return New ApiCallResult With {
            .RawResponse = rawJson,
            .Message = msg,
            .IsSuccess = httpResponse.IsSuccessStatusCode,
            .ParsedResponse = parsed
        }

        End Using

    End Function


    'get method
    Public Async Function GetRequestAsync(url As String, apiKey As String) As Task(Of GetApiResponse)
        Using client As New HttpClient()
            client.DefaultRequestHeaders.Add("api-key", apiKey)
            Dim response As HttpResponseMessage = Await client.GetAsync(url)
            response.EnsureSuccessStatusCode()

            Dim jsonResponse As String = Await response.Content.ReadAsStringAsync()
            Dim result As GetApiResponse = JsonConvert.DeserializeObject(Of GetApiResponse)(jsonResponse)

            Return result
        End Using
    End Function
    Private Sub ErrorLog(ByVal errorMsg As String)

        Try
            Dim oTbl As SAPbobsCOM.UserTable =
            objMain.objCompany.UserTables.Item("TNX_ERRORLOG")

            ' Unique key
            Dim key As String = Guid.NewGuid().ToString("N")

            oTbl.Code = key
            oTbl.Name = key

            ' Date & Time
            oTbl.UserFields.Fields.Item("U_Date").Value = Date.Now
            oTbl.UserFields.Fields.Item("U_Time").Value = Date.Now.ToString("HHmmss")

            ' Error message (max 254 chars)
            oTbl.UserFields.Fields.Item("U_Error").Value =
            errorMsg.Substring(0, Math.Min(254, errorMsg.Length))

            oTbl.Add()

        Catch
            ' ❗ Never throw error from error logging
        End Try

    End Sub


    Public Class ApiResponse
        Public Property message As String
        Public Property details As ApiDetails
        Public Property zatca_response As ZatcaResponse
    End Class

    Public Class ApiDetails
        Public Property document_id As String
        Public Property doc_type As String
        Public Property view_type As String
    End Class

    Public Class ZatcaResponse
        Public Property message As String
        Public Property [error] As Boolean
    End Class
    Public Class GetApiResponse
        Public Property xml As String
        Public Property qr As String
        Public Property message As String
    End Class
    Public Class ApiCallResult
        Public Property RawResponse As String
        Public Property Message As String
        Public Property IsSuccess As Boolean
        Public Property ParsedResponse As ApiResponse   ' ⭐ ADD THIS

    End Class



End Class

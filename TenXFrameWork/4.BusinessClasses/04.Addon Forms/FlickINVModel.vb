Imports Newtonsoft.Json

Public Class FlickINVModel

    Public Class FlickINV
        Public Property invoice_number As String
        Public Property invoice_date As String
        Public Property issue_time As String
        Public Property doc_type As String
        Public Property inv_type As String
        Public Property payment_terms As String

        Public Property notes As String

        Public Property invoice_ref_number As String
        Public Property issue_date As String

        Public Property currency As String
        Public Property party_details As Party_Details
        Public Property delivery_details As Delivery_Details
        'Public Property lineitems() As Lineitem
        Public Property notes_details As notes_details
        Public lineitems As IList(Of Lineitem) = New List(Of Lineitem)()

    End Class

    Public Class Party_Details

        Public Property name As String
        Public Property phone As String
        Public Property country As String
        Public Property country_code As String
        Public Property email As String

        Public Property party_name_ar As String
        Public Property party_name_en As String
        Public Property party_vat As String
        Public Property city_ar As String
        Public Property city_en As String
        Public Property city_subdivision_ar As String
        Public Property city_subdivision_en As String
        Public Property street_ar As String
        Public Property street_en As String
        Public Property building As String
        Public Property postal_zone As String

    End Class

    Public Class Delivery_Details
        Public Property actual_delivery As String
        Public Property latest_delivery As String
    End Class

    Public Class notes_details
        Public Property related_doc_reference As String
        Public Property reason_text As String
    End Class

    Public Class Lineitem
        Public Property item_name As String
        Public Property item_id As String
        Public Property unit_of_measure As String
        Public Property quantity As Decimal
        Public Property unit_rate As Decimal
        Public Property vat_category_code As String
        Public Property vat_rate As Decimal


        Public Property name_en As String
        Public Property name_ar As String
        Public Property tax_category As String
        Public Property tax_exclusive_price As Integer
        Public Property tax_percentage As Single
    End Class

    Public Class BilledTo
        Public Property name As String


        Public Property email As String
        Public Property phone As String



        Public Property country As String

        Public Property country_code As String





    End Class

    Public Class ItemDetail
        Public Property vat_category_code As String
        Public Property vat_exemption_code As String
        Public Property vat_exemption_reason As String
        Public Property vat_rate As Decimal
        Public Property item_name As String
        Public Property item_price_discount As Decimal
        Public Property price_allowance_indicator As String
        Public Property quantity As Decimal
        Public Property unit_rate As Decimal
        Public Property service_code As String
        Public Property is_discount_applicable As Boolean
        Public Property unit_of_measure As String
        Public Property item_id As String
        Public Property other_id As String
        Public Property other_id_type As String
    End Class
    Public Class E_InvoiceRequest
        Public Property type As String
        Public Property invoice_date As String
        Public Property due_date As String
        Public Property invoice_number As String
        Public Property logo_url As String
        Public Property payment_terms As String
        Public Property billed_to As BilledTo
        Public Property item_details As List(Of ItemDetail)
        Public Property notes As String
        Public Property payment_means As String
        Public Property header_discount As Decimal
        Public Property invoice_currency As String
        Public Property business_kind As String
        Public Property invoice_transaction_code As String
    End Class
    Public Class E_CreditNoteRequest

        Public Property type As String
        Public Property credit_note_number As String
        Public Property credit_note_date As String

        Public Property invoice_number As String
        Public Property invoice_date As String

        Public Property reason_for_issuance As String
        Public Property payment_terms As String

        Public Property billed_to As BilledTo
        Public Property item_details As List(Of ItemDetail)

        Public Property payment_means As String
        Public Property invoice_currency As String
        Public Property total_in_words_arabic As String
        Public Property business_kind As String

        Public Property document_number_type As String

    End Class


    Public Class E_InvoiceRequestB2B
        Public Property type As String
        Public Property invoice_date As String
        Public Property due_date As String
        Public Property invoice_number As String
        Public Property logo_url As String
        Public Property payment_terms As String
        Public Property billed_to As BilledToB2B
        Public Property item_details As List(Of ItemDetailB2B)
        Public Property notes As String
        Public Property payment_means As String
        Public Property header_discount As Decimal
        Public Property invoice_currency As String
        Public Property business_kind As String
        Public Property invoice_transaction_code As String
    End Class
    Public Class E_CreditNoteRequestB2B

        Public Property type As String
        Public Property credit_note_number As String
        Public Property credit_note_date As String

        Public Property invoice_number As String
        Public Property invoice_date As String

        Public Property reason_for_issuance As String
        Public Property payment_terms As String

        Public Property billed_to As BilledToB2B
        Public Property item_details As List(Of ItemDetailB2B)

        Public Property payment_means As String
        Public Property invoice_currency As String
        Public Property total_in_words_arabic As String
        Public Property business_kind As String

        Public Property document_number_type As String

    End Class


    Public Class ItemDetailB2B
        Public Property vat_category_code As String
        Public Property vat_exemption_code As String
        Public Property vat_exemption_reason As String
        Public Property vat_rate As Decimal
        Public Property item_name As String
        Public Property item_price_discount As Decimal
        Public Property price_allowance_indicator As String
        Public Property quantity As Decimal
        Public Property unit_rate As Decimal
        Public Property service_code As String
        Public Property is_discount_applicable As Boolean
        Public Property unit_of_measure As String
        Public Property item_id As String
        Public Property other_id As String
        Public Property other_id_type As String
    End Class
    Public Class BilledToB2B
        Public Property name As String
        Public Property company_name As String
        Public Property logo_url As String
        Public Property email As String
        Public Property phone As String
        Public Property address1 As String
        Public Property city As String
        Public Property district As String
        Public Property postal_code As String
        Public Property country As String
        Public Property company_reg_number As String
        Public Property customer_number As String
        Public Property gst_no As String
        Public Property grp_vat_no As String
        Public Property other_id As String
        Public Property other_id_type As String
        Public Property building_number As String
        Public Property country_code As String
        Public Property neighbourhood As String
        Public Property company_name_local As String
        Public Property iso_certificate_url As String
        Public Property profile_type As String
        Public Property is_saudi_vat_registered As Boolean
        Public Property payment_card_last_four_digits As String
        Public Property address_in_local_language As String

    End Class
    Public Class InvoiceModel

        Public Property ReferenceNumber As String
        Public Property FinancialYear As String
        Public Property edelivery As String
        Public Property InvoiceID As String
        Public Property IssueDate As String
        Public Property DueDate As String
        Public Property InvoiceTypeCode As String

        Public Property Note As String
        Public Property DocumentCurrencyCode As String
        Public Property VATCurrencyCode As String
        Public Property AccountingCost As String
        Public Property BuyerReference As String

        Public Property InvoicePeriod As InvoicePeriodModel

        Public Property VATtaxpointdate As String
        Public Property PrincipleID As String
        Public Property BeneficiaryID As String
        Public Property Currencyexchangerate As String
        Public Property Salesorderreference As String
        Public Property Despatchadvicereference As String

        Public Property Receivingadvicereference As String

        Public Property Customsreferencenumber As String

        Public Property Purchaseorderreference As String
        Public Property FrequencyofBilling As String


        Public Property BillingReference As List(Of BillingReferenceModel)
        Public Property AdditionalDocumentDetails As List(Of AdditionalDocumentModel)

        Public Property AccountingSupplierParty As AccountingSupplierPartyModel
        Public Property AccountingCustomerParty As AccountingCustomerPartyModel

        Public Property Delivery As DeliveryModel

        Public Property AllowanceCharge As List(Of AllowanceChargeModel)

        Public Property LegalMonetaryTotal As TotalModel

        Public Property PaymentMeans As PaymentMeansModel
        Public Property PaymentTerms As PaymentTermsModel

        Public Property InvoiceLine As List(Of InvoiceLineModel)

    End Class
    Public Class DeliveryModel

        <JsonProperty("Delivery.ActualDeliveryDate")>
        Public Property ActualDeliveryDate As String

        <JsonProperty("Delivery.DeliveryLocation")>
        Public Property DeliveryLocation As DeliveryLocationModel

        <JsonProperty("Delivery.DeliveryParty")>
        Public Property DeliveryParty As DeliveryPartyModel

    End Class

    Public Class DeliveryLocationModel
        <JsonProperty("Delivery.Name")>
        Public Property Name As String

        <JsonProperty("Delivery.ID")>
        Public Property ID As DeliveryIDModel

        <JsonProperty("Delivery.PostalAddress")>
        Public Property PostalAddress As DeliveryAddressModel
    End Class

    Public Class DeliveryIDModel
        <JsonProperty("Delivery.schemeID")>
        Public Property schemeID As String

        <JsonProperty("Delivery.DeliveryIdentifier")>
        Public Property DeliveryIdentifier As String
    End Class

    Public Class DeliveryAddressModel
        <JsonProperty("Delivery.StreetName")>
        Public Property StreetName As String

        Public Property AdditionalStreetName As String


        <JsonProperty("Delivery.CityName")>
        Public Property CityName As String

        <JsonProperty("Delivery.PostalZone")>
        Public Property PostalZone As String

        <JsonProperty("Delivery.Country")>
        Public Property Country As DeliveryCountryModel
    End Class

    Public Class DeliveryCountryModel
        <JsonProperty("Delivery.IdentificationCode")>
        Public Property IdentificationCode As String
    End Class

    Public Class DeliveryPartyModel
        <JsonProperty("Delivery.PartyName")>
        Public Property PartyName As String
    End Class
    Public Class BillingReferenceModel

        <JsonProperty("BillingReference.InvoiceDocumentReference.ID")>
        Public Property ID As String

        <JsonProperty("BillingReference.InvoiceDocumentReference.IssueDate")>
        Public Property IssueDate As String

        Public Property Creditnotereasoncode As String

    End Class
    Public Class AdditionalDocumentModel

        Public Property AdditionalDocumentReferenceID As String
        Public Property SchemeID As String
        Public Property DocumentTypecode As String
        Public Property DocumentDescription As String
        Public Property EmbeddedDocumentBinaryObject As String
        Public Property mimeCode As String
        Public Property FileName As String
        Public Property ExternalReferenceID As String

    End Class
    Public Class InvoicePeriodModel
        <JsonProperty("InvoicePeriod.StartDate")>
        Public Property StartDate As String

        <JsonProperty("InvoicePeriod.EndDate")>
        Public Property EndDate As String
    End Class
    Public Class AllowanceChargeModel

        <JsonProperty("AllowanceCharge.AllowanceID")>
        Public Property AllowanceID As String

        Public Property ChargeIndicator As Boolean
        Public Property AllowanceChargeReasonCode As String
        Public Property MultiplierFactorNumeric As String

        <JsonProperty("AllowanceCharge.Amount")>
        Public Property Amount As AmountModel

        <JsonProperty("AllowanceCharge.BaseAmount")>
        Public Property BaseAmount As BaseAmountModel
        <JsonProperty("AllowanceCharge.TaxCategory")>
        Public Property TaxCategory As TaxCategoryModel

    End Class
    Public Class TaxCategoryModel

        <JsonProperty("ClassifiedTaxCategory.ID")>
        Public Property ID As String

        <JsonProperty("ClassifiedTaxCategory.Percent")>
        Public Property Percent As String

        <JsonProperty("ClassifiedTaxCategory.TaxScheme")>
        Public Property TaxScheme As TaxSchemesModel
        <JsonProperty("ClassifiedTaxCategory.TaxSchemeId")>
        Public Property TaxSchemeId As String
        <JsonProperty("ClassifiedTaxCategory.TaxExemptionReasonCode")>
        Public Property TaxExemptionReasonCode As String

        <JsonProperty("ClassifiedTaxCategory.TaxExemptionReason")>
        Public Property TaxExemptionReason As String

    End Class
    Public Class TaxSchemesModel

        <JsonProperty("AllowanceCharge.ID")>
        Public Property ID As String
        Public Property TaxExemptionReasonCode As String
        Public Property TaxExemptionReason As String


    End Class
    Public Class PaymentMeansModel

        Public Property PayeeName As String
        Public Property PaymentInstructions As String
        Public Property PaymentDueDate As String
        Public Property PaymentMeansCode As String

        Public Property PayeeFinancialAccount As FinancialAccountModel

    End Class

    Public Class FinancialAccountModel
        <JsonProperty("PaymentMeans.ID")>
        Public Property ID As String

        Public Property Name As String

        Public Property FinancialBranch As BranchModel
    End Class

    Public Class BranchModel
        <JsonProperty("PaymentMeans.FinancialBranch.ID")>
        Public Property ID As String
    End Class
    Public Class AmountModel
        <JsonProperty("AllowanceCharge.AllowanceAmount")>
        Public Property AllowanceAmount As String
    End Class

    Public Class BaseAmountModel
        <JsonProperty("AllowanceCharge.BaseAmount.AllowanceDiscountBaseAmount")>
        Public Property AllowanceDiscountBaseAmount As String
    End Class
    Public Class AccountingSupplierPartyModel
        Public Property SellerCode As String

        <JsonProperty("AccountingSupplierParty.EndpointID")>
        Public Property EndpointID As SupplierEndpointModel

        <JsonProperty("AccountingSupplierParty.PartyIdentification")>
        Public Property PartyIdentification As List(Of SupplierPartyIdentificationModel)

        <JsonProperty("AccountingSupplierParty.PostalAddress")>
        Public Property PostalAddress As SupplierAddressModel

        <JsonProperty("AccountingSupplierParty.TaxScheme")>
        Public Property TaxScheme As SupplierTaxModel

        <JsonProperty("AccountingSupplierParty.PartyLegalEntity")>
        Public Property PartyLegalEntity As SupplierLegalModel

        <JsonProperty("AccountingSupplierParty.Contact")>
        Public Property Contact As SupplierContactModel

    End Class
    Public Class SupplierEndpointModel
        <JsonProperty("AccountingSupplierParty.schemeID")>
        Public Property schemeID As String

        <JsonProperty("AccountingSupplierParty.SellerElectronicAddress")>
        Public Property SellerElectronicAddress As String
    End Class
    Public Class SupplierAddressModel
        <JsonProperty("AccountingSupplierParty.StreetName")>
        Public Property StreetName As String
        <JsonProperty("AccountingSupplierParty.CitySubdivisionName")>
        Public Property CitySubdivisionName As String

        <JsonProperty("AccountingSupplierParty.AdditionalStreetName")>
        Public Property AdditionalStreetName As String

        <JsonProperty("AccountingSupplierParty.CityName")>
        Public Property CityName As String

        <JsonProperty("AccountingSupplierParty.PostalZone")>
        Public Property PostalZone As String

        <JsonProperty("AccountingSupplierParty.Country")>
        Public Property Country As SupplierCountryModel


    End Class

    Public Class SupplierCountryModel
        <JsonProperty("AccountingSupplierParty.IdentificationCode")>
        Public Property IdentificationCode As String
    End Class
    Public Class SupplierLegalModel
        <JsonProperty("AccountingSupplierParty.RegistrationName")>
        Public Property RegistrationName As String

        <JsonProperty("AccountingSupplierParty.CompanyID")>
        Public Property CompanyID As String
    End Class
    Public Class SupplierTaxModel
        <JsonProperty("SellerTaxidentifier")>
        Public Property SellerTaxidentifier As String

        <JsonProperty("AccountingSupplierParty.ID")>
        Public Property ID As String
    End Class
    Public Class SupplierContactModel
        <JsonProperty("AccountingSupplierParty.Name")>
        Public Property Name As String

        <JsonProperty("AccountingSupplierParty.Telephone")>
        Public Property Telephone As String

        <JsonProperty("AccountingSupplierParty.EmailID")>
        Public Property EmailID As String
    End Class
    Public Class SupplierPartyIdentificationModel

        Public Property SellerLegalSchemeIdentifier As String
        Public Property SellerLegalRegistrationType As String
        Public Property SellerCommercial_Tradelicense As String
        Public Property SellerEmiratesID As String
        Public Property SellerPassport As String
        Public Property SellerCabinetDecision As String
        Public Property SellerAuthorityname As String
        Public Property SellerPassportIssuingCountrycode As String

    End Class
    Public Class CustomerPartyIdentificationModel

        Public Property ID As String
        Public Property BuyerSchemeidentifier As String

    End Class
    Public Class CountryModel
        Public Property IdentificationCode As String
    End Class
    Public Class TaxSchemeModel
        Public Property ID As String

        ' Optional fields (based on your JSON)
        Public Property Buyerlegalregistrationidentifiertype As String
        Public Property BuyerCommercialorTradelicense As String
        Public Property BuyerEmiratesID As String
        Public Property BuyerPassport As String
        Public Property BuyerCabinetDecision As String
        Public Property BuyerAuthorityName As String
        Public Property BuyerPassportIssuingCountrycode As String
    End Class
    Public Class CustomerContactModel
        Public Property Name As String
        Public Property Telephone As String
        Public Property ElectronicEmail As String
    End Class
    Public Class CustomerCompanyModel
        Public Property schemeID As String
        Public Property RegistrationIdentifier As String
    End Class
    Public Class AccountingCustomerPartyModel

        Public Property BuyerCode As String
        Public Property EndpointID As CustomerEndpointModel
        Public Property PartyIdentification As CustomerPartyIdentificationModel
        Public Property PostalAddress As CustomerAddressModel
        Public Property PartyTaxScheme As CustomerTaxModel

    End Class
    Public Class CustomerEndpointModel
        Public Property SchemeID As String
        Public Property BuyerElectronicAddress As String
    End Class
    Public Class CustomerAddressModel
        Public Property CitySubdivisionName As String
        Public Property StreetName As String
        Public Property AdditionalStreetName As String
        Public Property CityName As String
        Public Property PostalZone As String
        Public Property Country As CountryModel
    End Class
    Public Class CustomerTaxModel
        Public Property CompanyID As String
        Public Property TaxScheme As TaxSchemeModel
        Public Property PartyLegalEntity As CustomerLegalModel
        Public Property Contact As CustomerContactModel
    End Class
    Public Class CustomerLegalModel
        Public Property RegistrationName As String
        Public Property CompanyID As CustomerCompanyModel
    End Class



    Public Class PaymentTermsModel


        Public Property Note As String


        Public Property PayeeAccountID As String


    End Class
    Public Class TotalModel


        Public Property TaxAmount As Decimal

        Public Property LineExtensionAmount As String

        Public Property TaxExclusiveAmount As String

        Public Property TaxInclusiveAmount As String

        Public Property ChargeTotalAmount As String

        Public Property PrepaidAmount As String

        Public Property PayableRoundingAmount As String

        Public Property PayableAmount As String

        Public Property TotalDiscountAmount As String


    End Class
    Public Class InvoiceLineModel

        <JsonProperty("InvoiceLine.InvoiceLineIdentifier")>
        Public Property InvoiceLineIdentifier As String

        <JsonProperty("InvoiceLine.InvoicedQuantity")>
        Public Property InvoicedQuantity As InvoicedQuantityModel

        <JsonProperty("LineExtensionAmount")>
        Public Property LineExtensionAmount As LineExtensionAmountModel

        <JsonProperty("InvoiceLine.LineItem")>
        Public Property LineItem As LineItemModel

    End Class
    Public Class LineExtensionAmountModel
        <JsonProperty("InvoiceLine.TaxExclusiveAmount")>
        Public Property TaxExclusiveAmount As String
    End Class

    Public Class InvoicedQuantityModel

        <JsonProperty("InvoiceLine.unitCode")>
        Public Property unitCode As String

        <JsonProperty("InvoiceLine.Quantity")>
        Public Property Quantity As String

    End Class
    Public Class LineItemModel
        <JsonProperty("InvoiceLine.Description")>
        Public Property Description As String

        <JsonProperty("InvoiceLine.Name")>
        Public Property Name As String

        Public Property Typeofgoodsorservices As String
        Public Property ServiceAccountingcode As String
        Public Property ServiceAccountingSchemeidentifier As String
        Public Property Itemtype As String

        Public Property StandardItemIdentification As StandardItemIdentificationModel

        Public Property CommodityClassification As CommodityClassificationModel
        Public Property ClassifiedTaxCategory As TaxCategoryModel

        Public Property LineAllowanceCharge As List(Of LineAllowanceChargeModel)

        Public Property Price As PriceModel

    End Class
    Public Class StandardItemIdentificationModel

        <JsonProperty("InvoiceLine.StandardItemIdentification.schemeID")>
        Public Property schemeID As String

        <JsonProperty("InvoiceLine.StandardItemIdentification.SchemeIDCode")>
        Public Property SchemeIDCode As String

    End Class
    Public Class CommodityClassificationModel

        <JsonProperty("CommodityClassification.listID")>
        Public Property listID As String

        <JsonProperty("CommodityClassification.Itemcode")>
        Public Property Itemcode As String

    End Class
    Public Class ClassifiedTaxCategoryModel

        <JsonProperty("ClassifiedTaxCategory.ID")>
        Public Property ID As String

        <JsonProperty("ClassifiedTaxCategory.Percent")>
        Public Property Percent As String

        <JsonProperty("ClassifiedTaxCategory.TaxSchemeId")>
        Public Property TaxSchemeId As String

        Public Property TaxExemptionReasonCode As String
        Public Property TaxExemptionReason As String

    End Class
    Public Class LineAllowanceChargeModel

        <JsonProperty("InvoiceLine.LineAllowanceCharge.LineAllowanceCharge.LineAllowanceID")>
        Public Property LineAllowanceID As String

        <JsonProperty("LineAllowanceCharge.ChargeIndicator")>
        Public Property ChargeIndicator As String

        <JsonProperty("LineAllowanceCharge.AllowanceChargeReasonCode")>
        Public Property AllowanceChargeReasonCode As String


        <JsonProperty("LineAllowanceCharge.MultiplierFactorNumeric")>
        Public Property MultiplierFactorNumeric As String

        Public Property Amount As LineAllowanceAmountModel
        Public Property BaseAmount As LineAllowanceBaseModel

    End Class
    Public Class LineAllowanceAmountModel

        <JsonProperty("LineAllowanceCharge.CurrencyID")>
        Public Property CurrencyID As String

        <JsonProperty("LineAllowanceCharge.AllowanceAmount")>
        Public Property AllowanceAmount As String

    End Class
    Public Class LineAllowanceBaseModel

        <JsonProperty("LineAllowanceCharge.Base.CurrencyID")>
        Public Property CurrencyID As String

        <JsonProperty("LineAllowanceCharge.Base.AllowanceChargeBaseAmount")>
        Public Property AllowanceChargeBaseAmount As String

    End Class

    Public Class PriceModel
        Public Property BaseQuantity As String
        Public Property BaseQuantityUOM As String
        Public Property PriceAmount As PriceAmountModel
        Public Property AlwChg As AlwChgModel

    End Class
    Public Class AlwChgModel

        Public Property Amt As String
        Public Property BaseAmt As String

    End Class
    Public Class PriceAmountModel

        <JsonProperty("Price.PriceAmount")>
        Public Property Amount As String

    End Class
    Public Class InvoiceResponse
        Public Property Invoice As InvoiceData
    End Class

    Public Class InvoiceData
        Public Property ID As String
        Public Property UUID As String
        Public Property ErrorMessage As String
    End Class

End Class

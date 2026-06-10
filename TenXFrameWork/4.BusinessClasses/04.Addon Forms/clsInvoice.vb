Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks


Namespace Zatca_CSID.Models
    Public Class clsInvoice
        Public Property ID As String
        Public Property IssueDate As String
        Public Property IssueTime As String
        Public Property InvoiceTypeCode As Integer
        Public Property InvoiceTypeCodeName As String
        Public Property Cur_Code As String
        Public Property Tax_Cur_Code As String
        Public Property Doc_Ref_ID As String
        Public Property Adtl_Ref_UUID As Integer = 0
        Public Property Adtl_Ref_ICV As String
        Public Property Bill_ref As String
        Public Property DeliveryDate As String
        Public Property Actual_DeliveryDate As String
        Public Property PaymentMeansCode As String
        Public Property PaymentNotes As String
        Public allowancecharge As _AllowanceCharge = New _AllowanceCharge()
        Public invline As IList(Of _InvoiceLine) = New List(Of _InvoiceLine)()
        Public supplier As _SupCustInfo = New _SupCustInfo()
        Public customer As _SupCustInfo = New _SupCustInfo()
        Public authKey As AuthKey = New AuthKey()
    End Class

    Public Class _SupCustInfo
        Public Property ID As String
        Public Property SchemaID As String
        Public Property StreeName As String
        Public Property Building_Num As String
        Public Property PlotNum As String
        Public Property City As String
        Public Property Pincode As String
        Public Property SubDivision As String
        Public Property Country As String
        Public Property RegistrationName As String
        Public Property TaxID As String
    End Class

    Public Class _InvoiceLine
        Public Property Qty As Decimal = 0
        Public Property Tax_ID As String = String.Empty
        Public Property Taxper As Decimal = 0
        Public Property ExcludeVat As Boolean = False
        Public Property ItemName As String = String.Empty
        Public Property Price As Decimal = 0
        Public Property Discount_Amt As Decimal = 0
        Public Property Discount_Reason As String = String.Empty
        Public Property Tax_Excempt_reason As String = String.Empty
        Public Property Tax_Excempt_ReasonCode As String = String.Empty
    End Class

    Public Class AuthKey
        Public Cer As String = String.Empty
        Public Key As String = String.Empty
        Public Pih As String = String.Empty
    End Class

    Public Class EncodedInvoiceModel
        Public Property Cer As String = String.Empty
        Public Property OTP As String = String.Empty
        Public Property InvoiceTypeCode As Integer = 0
        Public Property EncodedInvoice As String
        Public Property InvoiceHash As String = String.Empty
        Public Property UUID As String = String.Empty
        Public Property IsProduction As Boolean = False
    End Class

    Public Class _AllowanceCharge
        Public Property ID As String = String.Empty
        Public Property MultiplierFactorNumeric As Decimal = 0
        Public Property Amount As Decimal = 0
        Public Property BaseAmount As Decimal = 0
        Public Property ChargeIndicator As Boolean = False
        Public Property TaxID As String = String.Empty
        Public Property Percent As Decimal = 0
        Public Property TaxExemptionReason As String = String.Empty
        Public Property TaxExemptionReasonCode As String = String.Empty
        Public Property AllowanceChargeReason As String = String.Empty
    End Class
End Namespace


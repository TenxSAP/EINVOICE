Imports System
Imports System.Collections.Generic
    Imports System.Linq
    Imports System.Text
    Imports System.Threading.Tasks


Namespace Zatca.Models

    ' NOTE: Generated code may require at least .NET Framework 4.5 or .NET Core/Standard 2.0.
    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:Invoice-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:Invoice-2", IsNullable:=False)>
    Partial Public Class clsInvoiceModel

        Private uBLExtensionsField As UBLExtensions

        Private profileIDField As String
        'Naresh ID to string
        Private idField As String

        Private uUIDField As String

        Private issueDateField As String

        Private issueTimeField As String

        Private invoiceTypeCodeField As InvoiceTypeCode

        Private noteField As Note

        Private documentCurrencyCodeField As String

        Private taxCurrencyCodeField As String

        Private additionalDocumentReferenceField() As AdditionalDocumentReference

        Private signatureField As Signature1

        Private accountingSupplierPartyField As AccountingSupplierParty

        Private accountingCustomerPartyField As AccountingCustomerParty

        Private paymentMeansField As PaymentMeans

        Private allowanceChargeField As AllowanceCharge

        Private taxTotalField() As TaxTotal

        Private legalMonetaryTotalField As LegalMonetaryTotal

        Private invoiceLineField() As InvoiceLine

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2")>
        Public Property UBLExtensions() As UBLExtensions
            Get
                Return Me.uBLExtensionsField
            End Get
            Set
                Me.uBLExtensionsField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property ProfileID() As String
            Get
                Return Me.profileIDField
            End Get
            Set
                Me.profileIDField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property ID() As String
            Get
                Return Me.idField
            End Get
            Set
                Me.idField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property UUID() As String
            Get
                Return Me.uUIDField
            End Get
            Set
                Me.uUIDField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2", DataType:="date")>
        Public Property IssueDate() As String
            Get
                Return Me.issueDateField
            End Get
            Set
                Me.issueDateField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2", DataType:="time")>
        Public Property IssueTime() As String
            Get
                Return Me.issueTimeField
            End Get
            Set
                Me.issueTimeField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property InvoiceTypeCode() As InvoiceTypeCode
            Get
                Return Me.invoiceTypeCodeField
            End Get
            Set
                Me.invoiceTypeCodeField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property Note() As Note
            Get
                Return Me.noteField
            End Get
            Set
                Me.noteField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property DocumentCurrencyCode() As String
            Get
                Return Me.documentCurrencyCodeField
            End Get
            Set
                Me.documentCurrencyCodeField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property TaxCurrencyCode() As String
            Get
                Return Me.taxCurrencyCodeField
            End Get
            Set
                Me.taxCurrencyCodeField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute("AdditionalDocumentReference", [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
        Public Property AdditionalDocumentReference() As AdditionalDocumentReference()
            Get
                Return Me.additionalDocumentReferenceField
            End Get
            Set
                Me.additionalDocumentReferenceField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
        Public Property Signature() As Signature1
            Get
                Return Me.signatureField
            End Get
            Set
                Me.signatureField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
        Public Property AccountingSupplierParty() As AccountingSupplierParty
            Get
                Return Me.accountingSupplierPartyField
            End Get
            Set
                Me.accountingSupplierPartyField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
        Public Property AccountingCustomerParty() As AccountingCustomerParty
            Get
                Return Me.accountingCustomerPartyField
            End Get
            Set
                Me.accountingCustomerPartyField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
        Public Property PaymentMeans() As PaymentMeans
            Get
                Return Me.paymentMeansField
            End Get
            Set
                Me.paymentMeansField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
        Public Property AllowanceCharge() As AllowanceCharge
            Get
                Return Me.allowanceChargeField
            End Get
            Set
                Me.allowanceChargeField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute("TaxTotal", [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
        Public Property TaxTotal() As TaxTotal()
            Get
                Return Me.taxTotalField
            End Get
            Set
                Me.taxTotalField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
        Public Property LegalMonetaryTotal() As LegalMonetaryTotal
            Get
                Return Me.legalMonetaryTotalField
            End Get
            Set
                Me.legalMonetaryTotalField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute("InvoiceLine", [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
        Public Property InvoiceLine() As InvoiceLine()
            Get
                Return Me.invoiceLineField
            End Get
            Set
                Me.invoiceLineField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2", IsNullable:=False)>
    Partial Public Class UBLExtensions

        Private uBLExtensionField As UBLExtensionsUBLExtension

        '''<remarks/>
        Public Property UBLExtension() As UBLExtensionsUBLExtension
            Get
                Return Me.uBLExtensionField
            End Get
            Set
                Me.uBLExtensionField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2")>
    Partial Public Class UBLExtensionsUBLExtension

        Private extensionURIField As String

        Private extensionContentField As UBLExtensionsUBLExtensionExtensionContent

        '''<remarks/>
        Public Property ExtensionURI() As String
            Get
                Return Me.extensionURIField
            End Get
            Set
                Me.extensionURIField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property ExtensionContent() As UBLExtensionsUBLExtensionExtensionContent
            Get
                Return Me.extensionContentField
            End Get
            Set
                Me.extensionContentField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2")>
    Partial Public Class UBLExtensionsUBLExtensionExtensionContent

        Private uBLDocumentSignaturesField As UBLDocumentSignatures

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonSignatureComponents-2")>
        Public Property UBLDocumentSignatures() As UBLDocumentSignatures
            Get
                Return Me.uBLDocumentSignaturesField
            End Get
            Set
                Me.uBLDocumentSignaturesField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonSignatureComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonSignatureComponents-2", IsNullable:=False)>
    Partial Public Class UBLDocumentSignatures

        Private signatureInformationField As SignatureInformation

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:SignatureAggregateComponents-2")>
        Public Property SignatureInformation() As SignatureInformation
            Get
                Return Me.signatureInformationField
            End Get
            Set
                Me.signatureInformationField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:SignatureAggregateComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:SignatureAggregateComponents-2", IsNullable:=False)>
    Partial Public Class SignatureInformation

        Private idField As String

        Private referencedSignatureIDField As String

        Private signatureField As Signature

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property ID() As String
            Get
                Return Me.idField
            End Get
            Set
                Me.idField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:SignatureBasicComponents-2")>
        Public Property ReferencedSignatureID() As String
            Get
                Return Me.referencedSignatureIDField
            End Get
            Set
                Me.referencedSignatureIDField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="http://www.w3.org/2000/09/xmldsig#")>
        Public Property Signature() As Signature
            Get
                Return Me.signatureField
            End Get
            Set
                Me.signatureField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2", IsNullable:=False)>
    Partial Public Class ID

        Public schemeIDField As String

        Public schemeAgencyIDField As Byte

        Public schemeAgencyIDFieldSpecified As Boolean

        Public valueField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property schemeID() As String
            Get
                Return Me.schemeIDField
            End Get
            Set
                Me.schemeIDField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property schemeAgencyID() As String
            Get
                Return Me.schemeAgencyIDField
            End Get
            Set
                Me.schemeAgencyIDField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlIgnoreAttribute()>
        Public Property schemeAgencyIDSpecified() As Boolean
            Get
                Return Me.schemeAgencyIDFieldSpecified
            End Get
            Set
                Me.schemeAgencyIDFieldSpecified = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlTextAttribute()>
        Public Property Value() As String
            Get
                Return Me.valueField
            End Get
            Set
                Me.valueField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.w3.org/2000/09/xmldsig#"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="http://www.w3.org/2000/09/xmldsig#", IsNullable:=False)>
    Partial Public Class Signature

        Private signedInfoField As SignatureSignedInfo

        Private signatureValueField As String

        Private keyInfoField As SignatureKeyInfo

        Private objectField As SignatureObject

        Private idField As String

        '''<remarks/>
        Public Property SignedInfo() As SignatureSignedInfo
            Get
                Return Me.signedInfoField
            End Get
            Set
                Me.signedInfoField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property SignatureValue() As String
            Get
                Return Me.signatureValueField
            End Get
            Set
                Me.signatureValueField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property KeyInfo() As SignatureKeyInfo
            Get
                Return Me.keyInfoField
            End Get
            Set
                Me.keyInfoField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property [Object]() As SignatureObject
            Get
                Return Me.objectField
            End Get
            Set
                Me.objectField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property Id() As String
            Get
                Return Me.idField
            End Get
            Set
                Me.idField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.w3.org/2000/09/xmldsig#")>
    Partial Public Class SignatureSignedInfo

        Private canonicalizationMethodField As SignatureSignedInfoCanonicalizationMethod

        Private signatureMethodField As SignatureSignedInfoSignatureMethod

        Private referenceField() As SignatureSignedInfoReference

        '''<remarks/>
        Public Property CanonicalizationMethod() As SignatureSignedInfoCanonicalizationMethod
            Get
                Return Me.canonicalizationMethodField
            End Get
            Set
                Me.canonicalizationMethodField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property SignatureMethod() As SignatureSignedInfoSignatureMethod
            Get
                Return Me.signatureMethodField
            End Get
            Set
                Me.signatureMethodField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute("Reference")>
        Public Property Reference() As SignatureSignedInfoReference()
            Get
                Return Me.referenceField
            End Get
            Set
                Me.referenceField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.w3.org/2000/09/xmldsig#")>
    Partial Public Class SignatureSignedInfoCanonicalizationMethod

        Private algorithmField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property Algorithm() As String
            Get
                Return Me.algorithmField
            End Get
            Set
                Me.algorithmField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.w3.org/2000/09/xmldsig#")>
    Partial Public Class SignatureSignedInfoSignatureMethod

        Private algorithmField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property Algorithm() As String
            Get
                Return Me.algorithmField
            End Get
            Set
                Me.algorithmField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.w3.org/2000/09/xmldsig#")>
    Partial Public Class SignatureSignedInfoReference

        Private transformsField() As SignatureSignedInfoReferenceTransform

        Private digestMethodField As SignatureSignedInfoReferenceDigestMethod

        Private digestValueField As String

        Private idField As String

        Private uRIField As String

        Private typeField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlArrayItemAttribute("Transform", IsNullable:=False)>
        Public Property Transforms() As SignatureSignedInfoReferenceTransform()
            Get
                Return Me.transformsField
            End Get
            Set
                Me.transformsField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property DigestMethod() As SignatureSignedInfoReferenceDigestMethod
            Get
                Return Me.digestMethodField
            End Get
            Set
                Me.digestMethodField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property DigestValue() As String
            Get
                Return Me.digestValueField
            End Get
            Set
                Me.digestValueField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property Id() As String
            Get
                Return Me.idField
            End Get
            Set
                Me.idField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property URI() As String
            Get
                Return Me.uRIField
            End Get
            Set
                Me.uRIField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property Type() As String
            Get
                Return Me.typeField
            End Get
            Set
                Me.typeField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.w3.org/2000/09/xmldsig#")>
    Partial Public Class SignatureSignedInfoReferenceTransform

        Private xPathField As String

        Private algorithmField As String

        '''<remarks/>
        Public Property XPath() As String
            Get
                Return Me.xPathField
            End Get
            Set
                Me.xPathField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property Algorithm() As String
            Get
                Return Me.algorithmField
            End Get
            Set
                Me.algorithmField = Value
            End Set
        End Property
    End Class

    '''<SignatureSignedInfoReferenceDigestMethod/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.w3.org/2000/09/xmldsig#")>
    Partial Public Class SignatureSignedInfoReferenceDigestMethod

        Private algorithmField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property Algorithm() As String
            Get
                Return Me.algorithmField
            End Get
            Set
                Me.algorithmField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.w3.org/2000/09/xmldsig#")>
    Partial Public Class SignatureKeyInfo

        Private x509DataField As SignatureKeyInfoX509Data

        '''<remarks/>
        Public Property X509Data() As SignatureKeyInfoX509Data
            Get
                Return Me.x509DataField
            End Get
            Set
                Me.x509DataField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.w3.org/2000/09/xmldsig#")>
    Partial Public Class SignatureKeyInfoX509Data

        Private x509CertificateField As String

        '''<remarks/>
        Public Property X509Certificate() As String
            Get
                Return Me.x509CertificateField
            End Get
            Set
                Me.x509CertificateField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.w3.org/2000/09/xmldsig#")>
    Partial Public Class SignatureObject

        Private qualifyingPropertiesField As QualifyingProperties

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="http://uri.etsi.org/01903/v1.3.2#")>
        Public Property QualifyingProperties() As QualifyingProperties
            Get
                Return Me.qualifyingPropertiesField
            End Get
            Set
                Me.qualifyingPropertiesField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://uri.etsi.org/01903/v1.3.2#"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="http://uri.etsi.org/01903/v1.3.2#", IsNullable:=False)>
    Partial Public Class QualifyingProperties

        Private signedPropertiesField As QualifyingPropertiesSignedProperties

        Private targetField As String

        '''<remarks/>
        Public Property SignedProperties() As QualifyingPropertiesSignedProperties
            Get
                Return Me.signedPropertiesField
            End Get
            Set
                Me.signedPropertiesField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property Target() As String
            Get
                Return Me.targetField
            End Get
            Set
                Me.targetField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://uri.etsi.org/01903/v1.3.2#")>
    Partial Public Class QualifyingPropertiesSignedProperties

        Private signedSignaturePropertiesField As QualifyingPropertiesSignedPropertiesSignedSignatureProperties

        Private idField As String

        '''<remarks/>
        Public Property SignedSignatureProperties() As QualifyingPropertiesSignedPropertiesSignedSignatureProperties
            Get
                Return Me.signedSignaturePropertiesField
            End Get
            Set
                Me.signedSignaturePropertiesField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property Id() As String
            Get
                Return Me.idField
            End Get
            Set
                Me.idField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://uri.etsi.org/01903/v1.3.2#")>
    Partial Public Class QualifyingPropertiesSignedPropertiesSignedSignatureProperties

        Private signingTimeField As Date

        Private signingCertificateField As QualifyingPropertiesSignedPropertiesSignedSignaturePropertiesSigningCertificate

        '''<remarks/>
        Public Property SigningTime() As Date
            Get
                Return Me.signingTimeField
            End Get
            Set
                Me.signingTimeField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property SigningCertificate() As QualifyingPropertiesSignedPropertiesSignedSignaturePropertiesSigningCertificate
            Get
                Return Me.signingCertificateField
            End Get
            Set
                Me.signingCertificateField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://uri.etsi.org/01903/v1.3.2#")>
    Partial Public Class QualifyingPropertiesSignedPropertiesSignedSignaturePropertiesSigningCertificate

        Private certField As QualifyingPropertiesSignedPropertiesSignedSignaturePropertiesSigningCertificateCert

        '''<remarks/>
        Public Property Cert() As QualifyingPropertiesSignedPropertiesSignedSignaturePropertiesSigningCertificateCert
            Get
                Return Me.certField
            End Get
            Set
                Me.certField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://uri.etsi.org/01903/v1.3.2#")>
    Partial Public Class QualifyingPropertiesSignedPropertiesSignedSignaturePropertiesSigningCertificateCert

        Private certDigestField As QualifyingPropertiesSignedPropertiesSignedSignaturePropertiesSigningCertificateCertCertDigest

        Private issuerSerialField As QualifyingPropertiesSignedPropertiesSignedSignaturePropertiesSigningCertificateCertIssuerSerial

        '''<remarks/>
        Public Property CertDigest() As QualifyingPropertiesSignedPropertiesSignedSignaturePropertiesSigningCertificateCertCertDigest
            Get
                Return Me.certDigestField
            End Get
            Set
                Me.certDigestField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property IssuerSerial() As QualifyingPropertiesSignedPropertiesSignedSignaturePropertiesSigningCertificateCertIssuerSerial
            Get
                Return Me.issuerSerialField
            End Get
            Set
                Me.issuerSerialField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://uri.etsi.org/01903/v1.3.2#")>
    Partial Public Class QualifyingPropertiesSignedPropertiesSignedSignaturePropertiesSigningCertificateCertCertDigest

        Private digestMethodField As DigestMethod

        Private digestValueField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="http://www.w3.org/2000/09/xmldsig#")>
        Public Property DigestMethod() As DigestMethod
            Get
                Return Me.digestMethodField
            End Get
            Set
                Me.digestMethodField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="http://www.w3.org/2000/09/xmldsig#")>
        Public Property DigestValue() As String
            Get
                Return Me.digestValueField
            End Get
            Set
                Me.digestValueField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.w3.org/2000/09/xmldsig#"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="http://www.w3.org/2000/09/xmldsig#", IsNullable:=False)>
    Partial Public Class DigestMethod

        Private algorithmField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property Algorithm() As String
            Get
                Return Me.algorithmField
            End Get
            Set
                Me.algorithmField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://uri.etsi.org/01903/v1.3.2#")>
    Partial Public Class QualifyingPropertiesSignedPropertiesSignedSignaturePropertiesSigningCertificateCertIssuerSerial

        Private x509IssuerNameField As String

        Private x509SerialNumberField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="http://www.w3.org/2000/09/xmldsig#")>
        Public Property X509IssuerName() As String
            Get
                Return Me.x509IssuerNameField
            End Get
            Set
                Me.x509IssuerNameField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="http://www.w3.org/2000/09/xmldsig#", DataType:="integer")>
        Public Property X509SerialNumber() As String
            Get
                Return Me.x509SerialNumberField
            End Get
            Set
                Me.x509SerialNumberField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2", IsNullable:=False)>
    Partial Public Class InvoiceTypeCode

        Private nameField As UInteger

        Private valueField As UShort

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property name() As UInteger
            Get
                Return Me.nameField
            End Get
            Set
                Me.nameField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlTextAttribute()>
        Public Property Value() As UShort
            Get
                Return Me.valueField
            End Get
            Set
                Me.valueField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2", IsNullable:=False)>
    Partial Public Class Note

        Public languageIDField As String

        Public valueField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property languageID() As String
            Get
                Return Me.languageIDField
            End Get
            Set
                Me.languageIDField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlTextAttribute()>
        Public Property Value() As String
            Get
                Return Me.valueField
            End Get
            Set
                Me.valueField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2", IsNullable:=False)>
    Partial Public Class AdditionalDocumentReference
        'Naresh changed to ID to String
        Public idField As String

        Public attachmentField As AdditionalDocumentReferenceAttachment

        Public uUIDField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property ID() As String
            Get
                Return Me.idField
            End Get
            Set
                Me.idField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property Attachment() As AdditionalDocumentReferenceAttachment
            Get
                Return Me.attachmentField
            End Get
            Set
                Me.attachmentField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property UUID() As String
            Get
                Return Me.uUIDField
            End Get
            Set
                Me.uUIDField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class AdditionalDocumentReferenceAttachment

        Private embeddedDocumentBinaryObjectField As EmbeddedDocumentBinaryObject

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property EmbeddedDocumentBinaryObject() As EmbeddedDocumentBinaryObject
            Get
                Return Me.embeddedDocumentBinaryObjectField
            End Get
            Set
                Me.embeddedDocumentBinaryObjectField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2", IsNullable:=False)>
    Partial Public Class EmbeddedDocumentBinaryObject

        Private mimeCodeField As String

        Private valueField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property mimeCode() As String
            Get
                Return Me.mimeCodeField
            End Get
            Set
                Me.mimeCodeField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlTextAttribute()>
        Public Property Value() As String
            Get
                Return Me.valueField
            End Get
            Set
                Me.valueField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"),
 System.Xml.Serialization.XmlRootAttribute("Signature", [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2", IsNullable:=False)>
    Partial Public Class Signature1

        Private idField As String

        Private signatureMethodField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property ID() As String
            Get
                Return Me.idField
            End Get
            Set
                Me.idField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property SignatureMethod() As String
            Get
                Return Me.signatureMethodField
            End Get
            Set
                Me.signatureMethodField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2", IsNullable:=False)>
    Partial Public Class AccountingSupplierParty

        Private partyField As AccountingSupplierPartyParty

        '''<remarks/>
        Public Property Party() As AccountingSupplierPartyParty
            Get
                Return Me.partyField
            End Get
            Set
                Me.partyField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class AccountingSupplierPartyParty

        Private partyIdentificationField As AccountingSupplierPartyPartyPartyIdentification

        Private postalAddressField As AccountingSupplierPartyPartyPostalAddress

        Private partyTaxSchemeField As AccountingSupplierPartyPartyPartyTaxScheme

        Private partyLegalEntityField As AccountingSupplierPartyPartyPartyLegalEntity

        '''<remarks/>
        Public Property PartyIdentification() As AccountingSupplierPartyPartyPartyIdentification
            Get
                Return Me.partyIdentificationField
            End Get
            Set
                Me.partyIdentificationField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property PostalAddress() As AccountingSupplierPartyPartyPostalAddress
            Get
                Return Me.postalAddressField
            End Get
            Set
                Me.postalAddressField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property PartyTaxScheme() As AccountingSupplierPartyPartyPartyTaxScheme
            Get
                Return Me.partyTaxSchemeField
            End Get
            Set
                Me.partyTaxSchemeField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property PartyLegalEntity() As AccountingSupplierPartyPartyPartyLegalEntity
            Get
                Return Me.partyLegalEntityField
            End Get
            Set
                Me.partyLegalEntityField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class AccountingSupplierPartyPartyPartyIdentification

        Public idField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property ID() As String
            Get
                Return Me.idField
            End Get
            Set
                Me.idField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class AccountingSupplierPartyPartyPostalAddress

        Private streetNameField As String

        Private buildingNumberField As UShort

        Private citySubdivisionNameField As String

        Private cityNameField As String

        Private postalZoneField As UShort

        Private countryField As AccountingSupplierPartyPartyPostalAddressCountry

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property StreetName() As String
            Get
                Return Me.streetNameField
            End Get
            Set
                Me.streetNameField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property BuildingNumber() As UShort
            Get
                Return Me.buildingNumberField
            End Get
            Set
                Me.buildingNumberField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property CitySubdivisionName() As String
            Get
                Return Me.citySubdivisionNameField
            End Get
            Set
                Me.citySubdivisionNameField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property CityName() As String
            Get
                Return Me.cityNameField
            End Get
            Set
                Me.cityNameField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property PostalZone() As UShort
            Get
                Return Me.postalZoneField
            End Get
            Set
                Me.postalZoneField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property Country() As AccountingSupplierPartyPartyPostalAddressCountry
            Get
                Return Me.countryField
            End Get
            Set
                Me.countryField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class AccountingSupplierPartyPartyPostalAddressCountry

        Private identificationCodeField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property IdentificationCode() As String
            Get
                Return Me.identificationCodeField
            End Get
            Set
                Me.identificationCodeField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class AccountingSupplierPartyPartyPartyTaxScheme

        Private companyIDField As ULong

        Private taxSchemeField As AccountingSupplierPartyPartyPartyTaxSchemeTaxScheme

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property CompanyID() As ULong
            Get
                Return Me.companyIDField
            End Get
            Set
                Me.companyIDField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property TaxScheme() As AccountingSupplierPartyPartyPartyTaxSchemeTaxScheme
            Get
                Return Me.taxSchemeField
            End Get
            Set
                Me.taxSchemeField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class AccountingSupplierPartyPartyPartyTaxSchemeTaxScheme

        Private idField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property ID() As String
            Get
                Return Me.idField
            End Get
            Set
                Me.idField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class AccountingSupplierPartyPartyPartyLegalEntity

        Private registrationNameField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property RegistrationName() As String
            Get
                Return Me.registrationNameField
            End Get
            Set
                Me.registrationNameField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2", IsNullable:=False)>
    Partial Public Class AccountingCustomerParty

        Private partyField As AccountingCustomerPartyParty

        '''<remarks/>
        Public Property Party() As AccountingCustomerPartyParty
            Get
                Return Me.partyField
            End Get
            Set
                Me.partyField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class AccountingCustomerPartyParty

        Private postalAddressField As AccountingCustomerPartyPartyPostalAddress

        Private partyTaxSchemeField As AccountingCustomerPartyPartyPartyTaxScheme

        Private partyLegalEntityField As AccountingCustomerPartyPartyPartyLegalEntity

        '''<remarks/>
        Public Property PostalAddress() As AccountingCustomerPartyPartyPostalAddress
            Get
                Return Me.postalAddressField
            End Get
            Set
                Me.postalAddressField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property PartyTaxScheme() As AccountingCustomerPartyPartyPartyTaxScheme
            Get
                Return Me.partyTaxSchemeField
            End Get
            Set
                Me.partyTaxSchemeField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property PartyLegalEntity() As AccountingCustomerPartyPartyPartyLegalEntity
            Get
                Return Me.partyLegalEntityField
            End Get
            Set
                Me.partyLegalEntityField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class AccountingCustomerPartyPartyPostalAddress

        Private streetNameField As String

        Private buildingNumberField As UShort

        Private citySubdivisionNameField As String

        Private cityNameField As String

        Private postalZoneField As UShort

        Private countryField As AccountingCustomerPartyPartyPostalAddressCountry

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property StreetName() As String
            Get
                Return Me.streetNameField
            End Get
            Set
                Me.streetNameField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property BuildingNumber() As UShort
            Get
                Return Me.buildingNumberField
            End Get
            Set
                Me.buildingNumberField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property CitySubdivisionName() As String
            Get
                Return Me.citySubdivisionNameField
            End Get
            Set
                Me.citySubdivisionNameField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property CityName() As String
            Get
                Return Me.cityNameField
            End Get
            Set
                Me.cityNameField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property PostalZone() As UShort
            Get
                Return Me.postalZoneField
            End Get
            Set
                Me.postalZoneField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property Country() As AccountingCustomerPartyPartyPostalAddressCountry
            Get
                Return Me.countryField
            End Get
            Set
                Me.countryField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class AccountingCustomerPartyPartyPostalAddressCountry

        Private identificationCodeField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property IdentificationCode() As String
            Get
                Return Me.identificationCodeField
            End Get
            Set
                Me.identificationCodeField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class AccountingCustomerPartyPartyPartyTaxScheme

        Private companyIDField As ULong

        Private taxSchemeField As AccountingCustomerPartyPartyPartyTaxSchemeTaxScheme

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property CompanyID() As ULong
            Get
                Return Me.companyIDField
            End Get
            Set
                Me.companyIDField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property TaxScheme() As AccountingCustomerPartyPartyPartyTaxSchemeTaxScheme
            Get
                Return Me.taxSchemeField
            End Get
            Set
                Me.taxSchemeField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class AccountingCustomerPartyPartyPartyTaxSchemeTaxScheme

        Private idField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property ID() As String
            Get
                Return Me.idField
            End Get
            Set
                Me.idField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class AccountingCustomerPartyPartyPartyLegalEntity

        Private registrationNameField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property RegistrationName() As String
            Get
                Return Me.registrationNameField
            End Get
            Set
                Me.registrationNameField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2", IsNullable:=False)>
    Partial Public Class PaymentMeans

        Public paymentMeansCodeField As Byte

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property PaymentMeansCode() As Byte
            Get
                Return Me.paymentMeansCodeField
            End Get
            Set
                Me.paymentMeansCodeField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2", IsNullable:=False)>
    Partial Public Class AllowanceCharge

        Private chargeIndicatorField As Boolean

        Private allowanceChargeReasonField As String

        Private amountField As Amount

        Private taxCategoryField() As AllowanceChargeTaxCategory

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property ChargeIndicator() As Boolean
            Get
                Return Me.chargeIndicatorField
            End Get
            Set
                Me.chargeIndicatorField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property AllowanceChargeReason() As String
            Get
                Return Me.allowanceChargeReasonField
            End Get
            Set
                Me.allowanceChargeReasonField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property Amount() As Amount
            Get
                Return Me.amountField
            End Get
            Set
                Me.amountField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute("TaxCategory")>
        Public Property TaxCategory() As AllowanceChargeTaxCategory()
            Get
                Return Me.taxCategoryField
            End Get
            Set
                Me.taxCategoryField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2", IsNullable:=False)>
    Partial Public Class Amount

        Private currencyIDField As String

        Private valueField As Decimal

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property currencyID() As String
            Get
                Return Me.currencyIDField
            End Get
            Set
                Me.currencyIDField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlTextAttribute()>
        Public Property Value() As Decimal
            Get
                Return Me.valueField
            End Get
            Set
                Me.valueField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class AllowanceChargeTaxCategory

        Private idField As String

        Private percentField As Decimal

        Private taxSchemeField As AllowanceChargeTaxCategoryTaxScheme

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property ID() As String
            Get
                Return Me.idField
            End Get
            Set
                Me.idField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property Percent() As Decimal
            Get
                Return Me.percentField
            End Get
            Set
                Me.percentField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property TaxScheme() As AllowanceChargeTaxCategoryTaxScheme
            Get
                Return Me.taxSchemeField
            End Get
            Set
                Me.taxSchemeField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class AllowanceChargeTaxCategoryTaxScheme

        Private idField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property ID() As String
            Get
                Return Me.idField
            End Get
            Set
                Me.idField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2", IsNullable:=False)>
    Partial Public Class TaxTotal

        Private taxAmountField As TaxAmount

        Private taxSubtotalField As TaxTotalTaxSubtotal

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property TaxAmount() As TaxAmount
            Get
                Return Me.taxAmountField
            End Get
            Set
                Me.taxAmountField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property TaxSubtotal() As TaxTotalTaxSubtotal
            Get
                Return Me.taxSubtotalField
            End Get
            Set
                Me.taxSubtotalField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2", IsNullable:=False)>
    Partial Public Class TaxAmount

        Private currencyIDField As String

        Private valueField As Decimal

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property currencyID() As String
            Get
                Return Me.currencyIDField
            End Get
            Set
                Me.currencyIDField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlTextAttribute()>
        Public Property Value() As Decimal
            Get
                Return Me.valueField
            End Get
            Set
                Me.valueField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class TaxTotalTaxSubtotal

        Private taxableAmountField As TaxableAmount

        Private taxAmountField As TaxAmount

        Private taxCategoryField As TaxTotalTaxSubtotalTaxCategory

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property TaxableAmount() As TaxableAmount
            Get
                Return Me.taxableAmountField
            End Get
            Set
                Me.taxableAmountField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property TaxAmount() As TaxAmount
            Get
                Return Me.taxAmountField
            End Get
            Set
                Me.taxAmountField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property TaxCategory() As TaxTotalTaxSubtotalTaxCategory
            Get
                Return Me.taxCategoryField
            End Get
            Set
                Me.taxCategoryField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2", IsNullable:=False)>
    Partial Public Class TaxableAmount

        Private currencyIDField As String

        Private valueField As Decimal

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property currencyID() As String
            Get
                Return Me.currencyIDField
            End Get
            Set
                Me.currencyIDField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlTextAttribute()>
        Public Property Value() As Decimal
            Get
                Return Me.valueField
            End Get
            Set
                Me.valueField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class TaxTotalTaxSubtotalTaxCategory

        Private idField As String

        Private percentField As Decimal

        Private taxSchemeField As TaxTotalTaxSubtotalTaxCategoryTaxScheme

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property ID() As String
            Get
                Return Me.idField
            End Get
            Set
                Me.idField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property Percent() As Decimal
            Get
                Return Me.percentField
            End Get
            Set
                Me.percentField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property TaxScheme() As TaxTotalTaxSubtotalTaxCategoryTaxScheme
            Get
                Return Me.taxSchemeField
            End Get
            Set
                Me.taxSchemeField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class TaxTotalTaxSubtotalTaxCategoryTaxScheme

        Private idField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property ID() As String
            Get
                Return Me.idField
            End Get
            Set
                Me.idField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2", IsNullable:=False)>
    Partial Public Class LegalMonetaryTotal

        Private lineExtensionAmountField As LineExtensionAmount

        Private taxExclusiveAmountField As TaxExclusiveAmount

        Private taxInclusiveAmountField As TaxInclusiveAmount

        Private allowanceTotalAmountField As AllowanceTotalAmount

        Private prepaidAmountField As PrepaidAmount

        Private payableAmountField As PayableAmount

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property LineExtensionAmount() As LineExtensionAmount
            Get
                Return Me.lineExtensionAmountField
            End Get
            Set
                Me.lineExtensionAmountField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property TaxExclusiveAmount() As TaxExclusiveAmount
            Get
                Return Me.taxExclusiveAmountField
            End Get
            Set
                Me.taxExclusiveAmountField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property TaxInclusiveAmount() As TaxInclusiveAmount
            Get
                Return Me.taxInclusiveAmountField
            End Get
            Set
                Me.taxInclusiveAmountField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property AllowanceTotalAmount() As AllowanceTotalAmount
            Get
                Return Me.allowanceTotalAmountField
            End Get
            Set
                Me.allowanceTotalAmountField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property PrepaidAmount() As PrepaidAmount
            Get
                Return Me.prepaidAmountField
            End Get
            Set
                Me.prepaidAmountField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property PayableAmount() As PayableAmount
            Get
                Return Me.payableAmountField
            End Get
            Set
                Me.payableAmountField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2", IsNullable:=False)>
    Partial Public Class LineExtensionAmount

        Private currencyIDField As String

        Private valueField As Decimal

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property currencyID() As String
            Get
                Return Me.currencyIDField
            End Get
            Set
                Me.currencyIDField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlTextAttribute()>
        Public Property Value() As Decimal
            Get
                Return Me.valueField
            End Get
            Set
                Me.valueField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2", IsNullable:=False)>
    Partial Public Class TaxExclusiveAmount

        Private currencyIDField As String

        Private valueField As Decimal

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property currencyID() As String
            Get
                Return Me.currencyIDField
            End Get
            Set
                Me.currencyIDField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlTextAttribute()>
        Public Property Value() As Decimal
            Get
                Return Me.valueField
            End Get
            Set
                Me.valueField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2", IsNullable:=False)>
    Partial Public Class TaxInclusiveAmount

        Private currencyIDField As String

        Private valueField As Decimal

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property currencyID() As String
            Get
                Return Me.currencyIDField
            End Get
            Set
                Me.currencyIDField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlTextAttribute()>
        Public Property Value() As Decimal
            Get
                Return Me.valueField
            End Get
            Set
                Me.valueField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2", IsNullable:=False)>
    Partial Public Class AllowanceTotalAmount

        Private currencyIDField As String

        Private valueField As Decimal

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property currencyID() As String
            Get
                Return Me.currencyIDField
            End Get
            Set
                Me.currencyIDField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlTextAttribute()>
        Public Property Value() As Decimal
            Get
                Return Me.valueField
            End Get
            Set
                Me.valueField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2", IsNullable:=False)>
    Partial Public Class PrepaidAmount

        Private currencyIDField As String

        Private valueField As Decimal

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property currencyID() As String
            Get
                Return Me.currencyIDField
            End Get
            Set
                Me.currencyIDField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlTextAttribute()>
        Public Property Value() As Decimal
            Get
                Return Me.valueField
            End Get
            Set
                Me.valueField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2", IsNullable:=False)>
    Partial Public Class PayableAmount

        Private currencyIDField As String

        Private valueField As Decimal

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property currencyID() As String
            Get
                Return Me.currencyIDField
            End Get
            Set
                Me.currencyIDField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlTextAttribute()>
        Public Property Value() As Decimal
            Get
                Return Me.valueField
            End Get
            Set
                Me.valueField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2", IsNullable:=False)>
    Partial Public Class InvoiceLine

        Private idField As String

        Private invoicedQuantityField As InvoicedQuantity

        Private lineExtensionAmountField As LineExtensionAmount

        Private taxTotalField As InvoiceLineTaxTotal

        Private itemField As InvoiceLineItem

        Private priceField As InvoiceLinePrice

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property ID() As String
            Get
                Return Me.idField
            End Get
            Set
                Me.idField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property InvoicedQuantity() As InvoicedQuantity
            Get
                Return Me.invoicedQuantityField
            End Get
            Set
                Me.invoicedQuantityField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property LineExtensionAmount() As LineExtensionAmount
            Get
                Return Me.lineExtensionAmountField
            End Get
            Set
                Me.lineExtensionAmountField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property TaxTotal() As InvoiceLineTaxTotal
            Get
                Return Me.taxTotalField
            End Get
            Set
                Me.taxTotalField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property Item() As InvoiceLineItem
            Get
                Return Me.itemField
            End Get
            Set
                Me.itemField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property Price() As InvoiceLinePrice
            Get
                Return Me.priceField
            End Get
            Set
                Me.priceField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2", IsNullable:=False)>
    Partial Public Class InvoicedQuantity

        Private unitCodeField As String

        Private valueField As Decimal

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property unitCode() As String
            Get
                Return Me.unitCodeField
            End Get
            Set
                Me.unitCodeField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlTextAttribute()>
        Public Property Value() As Decimal
            Get
                Return Me.valueField
            End Get
            Set
                Me.valueField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class InvoiceLineTaxTotal

        Private taxAmountField As TaxAmount

        Private roundingAmountField As RoundingAmount

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property TaxAmount() As TaxAmount
            Get
                Return Me.taxAmountField
            End Get
            Set
                Me.taxAmountField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property RoundingAmount() As RoundingAmount
            Get
                Return Me.roundingAmountField
            End Get
            Set
                Me.roundingAmountField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2", IsNullable:=False)>
    Partial Public Class RoundingAmount

        Private currencyIDField As String

        Private valueField As Decimal

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property currencyID() As String
            Get
                Return Me.currencyIDField
            End Get
            Set
                Me.currencyIDField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlTextAttribute()>
        Public Property Value() As Decimal
            Get
                Return Me.valueField
            End Get
            Set
                Me.valueField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class InvoiceLineItem

        Private nameField As String

        Private classifiedTaxCategoryField As InvoiceLineItemClassifiedTaxCategory

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property Name() As String
            Get
                Return Me.nameField
            End Get
            Set
                Me.nameField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property ClassifiedTaxCategory() As InvoiceLineItemClassifiedTaxCategory
            Get
                Return Me.classifiedTaxCategoryField
            End Get
            Set
                Me.classifiedTaxCategoryField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class InvoiceLineItemClassifiedTaxCategory

        Private idField As String

        Private percentField As Decimal

        Private taxSchemeField As InvoiceLineItemClassifiedTaxCategoryTaxScheme

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property ID() As String
            Get
                Return Me.idField
            End Get
            Set
                Me.idField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property Percent() As Decimal
            Get
                Return Me.percentField
            End Get
            Set
                Me.percentField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property TaxScheme() As InvoiceLineItemClassifiedTaxCategoryTaxScheme
            Get
                Return Me.taxSchemeField
            End Get
            Set
                Me.taxSchemeField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class InvoiceLineItemClassifiedTaxCategoryTaxScheme

        Private idField As String

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property ID() As String
            Get
                Return Me.idField
            End Get
            Set
                Me.idField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class InvoiceLinePrice

        Private priceAmountField As PriceAmount

        Private allowanceChargeField As InvoiceLinePriceAllowanceCharge

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property PriceAmount() As PriceAmount
            Get
                Return Me.priceAmountField
            End Get
            Set
                Me.priceAmountField = Value
            End Set
        End Property

        '''<remarks/>
        Public Property AllowanceCharge() As InvoiceLinePriceAllowanceCharge
            Get
                Return Me.allowanceChargeField
            End Get
            Set
                Me.allowanceChargeField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2", IsNullable:=False)>
    Partial Public Class PriceAmount

        Private currencyIDField As String

        Private valueField As Decimal

        '''<remarks/>
        <System.Xml.Serialization.XmlAttributeAttribute()>
        Public Property currencyID() As String
            Get
                Return Me.currencyIDField
            End Get
            Set
                Me.currencyIDField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlTextAttribute()>
        Public Property Value() As Decimal
            Get
                Return Me.valueField
            End Get
            Set
                Me.valueField = Value
            End Set
        End Property
    End Class

    '''<remarks/>
    <System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")>
    Partial Public Class InvoiceLinePriceAllowanceCharge

        Private chargeIndicatorField As Boolean

        Private allowanceChargeReasonField As String

        Private amountField As Amount

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property ChargeIndicator() As Boolean
            Get
                Return Me.chargeIndicatorField
            End Get
            Set
                Me.chargeIndicatorField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property AllowanceChargeReason() As String
            Get
                Return Me.allowanceChargeReasonField
            End Get
            Set
                Me.allowanceChargeReasonField = Value
            End Set
        End Property

        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute([Namespace]:="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")>
        Public Property Amount() As Amount
            Get
                Return Me.amountField
            End Get
            Set
                Me.amountField = Value
            End Set
        End Property
    End Class



End Namespace



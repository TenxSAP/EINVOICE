#region Assembly System.Xml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089
// location unknown
// Decompiled with ICSharpCode.Decompiler 8.2.0.7535
#endregion

using System.Collections;
using System.Globalization;
using System.IO;
using System.Security;
using System.Security.Permissions;
using System.Text;
using System.Xml.Schema;
using System.Xml.XPath;

namespace System.Xml;

//
// Summary:
//     Represents an XML document. You can use this class to load, validate, edit, add,
//     and position XML in a document.
public class XmlDocument : XmlNode
{
    private XmlImplementation implementation;

    private DomNameTable domNameTable;

    private XmlLinkedNode lastChild;

    private XmlNamedNodeMap entities;

    private Hashtable htElementIdMap;

    private Hashtable htElementIDAttrDecl;

    private SchemaInfo schemaInfo;

    private XmlSchemaSet schemas;

    private bool reportValidity;

    private bool actualLoadingStatus;

    private XmlNodeChangedEventHandler onNodeInsertingDelegate;

    private XmlNodeChangedEventHandler onNodeInsertedDelegate;

    private XmlNodeChangedEventHandler onNodeRemovingDelegate;

    private XmlNodeChangedEventHandler onNodeRemovedDelegate;

    private XmlNodeChangedEventHandler onNodeChangingDelegate;

    private XmlNodeChangedEventHandler onNodeChangedDelegate;

    internal bool fEntRefNodesPresent;

    internal bool fCDataNodesPresent;

    private bool preserveWhitespace;

    private bool isLoading;

    internal string strDocumentName;

    internal string strDocumentFragmentName;

    internal string strCommentName;

    internal string strTextName;

    internal string strCDataSectionName;

    internal string strEntityName;

    internal string strID;

    internal string strXmlns;

    internal string strXml;

    internal string strSpace;

    internal string strLang;

    internal string strEmpty;

    internal string strNonSignificantWhitespaceName;

    internal string strSignificantWhitespaceName;

    internal string strReservedXmlns;

    internal string strReservedXml;

    internal string baseURI;

    private XmlResolver resolver;

    internal bool bSetResolver;

    internal object objLock;

    private XmlAttribute namespaceXml;

    internal static EmptyEnumerator EmptyEnumerator = new EmptyEnumerator();

    internal static IXmlSchemaInfo NotKnownSchemaInfo = new XmlSchemaInfo(XmlSchemaValidity.NotKnown);

    internal static IXmlSchemaInfo ValidSchemaInfo = new XmlSchemaInfo(XmlSchemaValidity.Valid);

    internal static IXmlSchemaInfo InvalidSchemaInfo = new XmlSchemaInfo(XmlSchemaValidity.Invalid);

    internal SchemaInfo DtdSchemaInfo
    {
        get
        {
            return schemaInfo;
        }
        set
        {
            schemaInfo = value;
        }
    }

    //
    // Summary:
    //     Gets the type of the current node.
    //
    // Returns:
    //     The node type. For XmlDocument nodes, this value is XmlNodeType.Document.
    public override XmlNodeType NodeType => XmlNodeType.Document;

    //
    // Summary:
    //     Gets the parent node of this node (for nodes that can have parents).
    //
    // Returns:
    //     Always returns null.
    public override XmlNode ParentNode => null;

    //
    // Summary:
    //     Gets the node containing the DOCTYPE declaration.
    //
    // Returns:
    //     The System.Xml.XmlNode containing the DocumentType (DOCTYPE declaration).
    public virtual XmlDocumentType DocumentType => (XmlDocumentType)FindChild(XmlNodeType.DocumentType);

    internal virtual XmlDeclaration Declaration
    {
        get
        {
            if (HasChildNodes)
            {
                return FirstChild as XmlDeclaration;
            }

            return null;
        }
    }

    //
    // Summary:
    //     Gets the System.Xml.XmlImplementation object for the current document.
    //
    // Returns:
    //     The XmlImplementation object for the current document.
    public XmlImplementation Implementation => implementation;

    //
    // Summary:
    //     Gets the qualified name of the node.
    //
    // Returns:
    //     For XmlDocument nodes, the name is #document.
    public override string Name => strDocumentName;

    //
    // Summary:
    //     Gets the local name of the node.
    //
    // Returns:
    //     For XmlDocument nodes, the local name is #document.
    public override string LocalName => strDocumentName;

    //
    // Summary:
    //     Gets the root System.Xml.XmlElement for the document.
    //
    // Returns:
    //     The XmlElement that represents the root of the XML document tree. If no root
    //     exists, null is returned.
    public XmlElement DocumentElement => (XmlElement)FindChild(XmlNodeType.Element);

    internal override bool IsContainer => true;

    internal override XmlLinkedNode LastNode
    {
        get
        {
            return lastChild;
        }
        set
        {
            lastChild = value;
        }
    }

    //
    // Summary:
    //     Gets the System.Xml.XmlDocument to which the current node belongs.
    //
    // Returns:
    //     For XmlDocument nodes (System.Xml.XmlDocument.NodeType equals XmlNodeType.Document),
    //     this property always returns null.
    public override XmlDocument OwnerDocument => null;

    //
    // Summary:
    //     Gets or sets the System.Xml.Schema.XmlSchemaSet object associated with this System.Xml.XmlDocument.
    //
    //
    // Returns:
    //     An System.Xml.Schema.XmlSchemaSet object containing the XML Schema Definition
    //     Language (XSD) schemas associated with this System.Xml.XmlDocument; otherwise,
    //     an empty System.Xml.Schema.XmlSchemaSet object.
    public XmlSchemaSet Schemas
    {
        get
        {
            if (schemas == null)
            {
                schemas = new XmlSchemaSet(NameTable);
            }

            return schemas;
        }
        set
        {
            schemas = value;
        }
    }

    internal bool CanReportValidity => reportValidity;

    internal bool HasSetResolver => bSetResolver;

    //
    // Summary:
    //     Sets the System.Xml.XmlResolver to use for resolving external resources.
    //
    // Returns:
    //     The XmlResolver to use.In version 1.1 of the.NET Framework, the caller must be
    //     fully trusted in order to specify an XmlResolver.
    //
    // Exceptions:
    //   T:System.Xml.XmlException:
    //     This property is set to null and an external DTD or entity is encountered.
    public virtual XmlResolver XmlResolver
    {
        set
        {
            if (value != null)
            {
                try
                {
                    new NamedPermissionSet("FullTrust").Demand();
                }
                catch (SecurityException inner)
                {
                    throw new SecurityException(Res.GetString("Xml_UntrustedCodeSettingResolver"), inner);
                }
            }

            resolver = value;
            if (!bSetResolver)
            {
                bSetResolver = true;
            }

            XmlDocumentType documentType = DocumentType;
            if (documentType != null)
            {
                documentType.DtdSchemaInfo = null;
            }
        }
    }

    //
    // Summary:
    //     Gets the System.Xml.XmlNameTable associated with this implementation.
    //
    // Returns:
    //     An XmlNameTable enabling you to get the atomized version of a string within the
    //     document.
    public XmlNameTable NameTable => implementation.NameTable;

    //
    // Summary:
    //     Gets or sets a value indicating whether to preserve white space in element content.
    //
    //
    // Returns:
    //     true to preserve white space; otherwise false. The default is false.
    public bool PreserveWhitespace
    {
        get
        {
            return preserveWhitespace;
        }
        set
        {
            preserveWhitespace = value;
        }
    }

    //
    // Summary:
    //     Gets a value indicating whether the current node is read-only.
    //
    // Returns:
    //     true if the current node is read-only; otherwise false. XmlDocument nodes always
    //     return false.
    public override bool IsReadOnly => false;

    internal XmlNamedNodeMap Entities
    {
        get
        {
            if (entities == null)
            {
                entities = new XmlNamedNodeMap(this);
            }

            return entities;
        }
        set
        {
            entities = value;
        }
    }

    internal bool IsLoading
    {
        get
        {
            return isLoading;
        }
        set
        {
            isLoading = value;
        }
    }

    internal bool ActualLoadingStatus
    {
        get
        {
            return actualLoadingStatus;
        }
        set
        {
            actualLoadingStatus = value;
        }
    }

    internal Encoding TextEncoding
    {
        get
        {
            if (Declaration != null)
            {
                string encoding = Declaration.Encoding;
                if (encoding.Length > 0)
                {
                    return System.Text.Encoding.GetEncoding(encoding);
                }
            }

            return null;
        }
    }

    //
    // Summary:
    //     Throws an System.InvalidOperationException in all cases.
    //
    // Returns:
    //     The values of the node and all its child nodes.
    //
    // Exceptions:
    //   T:System.InvalidOperationException:
    //     In all cases.
    public override string InnerText
    {
        set
        {
            throw new InvalidOperationException(Res.GetString("Xdom_Document_Innertext"));
        }
    }

    //
    // Summary:
    //     Gets or sets the markup representing the children of the current node.
    //
    // Returns:
    //     The markup of the children of the current node.
    //
    // Exceptions:
    //   T:System.Xml.XmlException:
    //     The XML specified when setting this property is not well-formed.
    public override string InnerXml
    {
        get
        {
            return base.InnerXml;
        }
        set
        {
            LoadXml(value);
        }
    }

    internal string Version => Declaration?.Version;

    internal string Encoding => Declaration?.Encoding;

    internal string Standalone => Declaration?.Standalone;

    //
    // Summary:
    //     Returns the Post-Schema-Validation-Infoset (PSVI) of the node.
    //
    // Returns:
    //     The System.Xml.Schema.IXmlSchemaInfo object representing the PSVI of the node.
    public override IXmlSchemaInfo SchemaInfo
    {
        get
        {
            if (reportValidity)
            {
                XmlElement documentElement = DocumentElement;
                if (documentElement != null)
                {
                    switch (documentElement.SchemaInfo.Validity)
                    {
                        case XmlSchemaValidity.Valid:
                            return ValidSchemaInfo;
                        case XmlSchemaValidity.Invalid:
                            return InvalidSchemaInfo;
                    }
                }
            }

            return NotKnownSchemaInfo;
        }
    }

    //
    // Summary:
    //     Gets the base URI of the current node.
    //
    // Returns:
    //     The location from which the node was loaded.
    public override string BaseURI => baseURI;

    internal override XPathNodeType XPNodeType => XPathNodeType.Root;

    internal bool HasEntityReferences => fEntRefNodesPresent;

    internal XmlAttribute NamespaceXml
    {
        get
        {
            if (namespaceXml == null)
            {
                namespaceXml = new XmlAttribute(AddAttrXmlName(strXmlns, strXml, strReservedXmlns, null), this);
                namespaceXml.Value = strReservedXml;
            }

            return namespaceXml;
        }
    }

    //
    // Summary:
    //     Occurs when a node belonging to this document is about to be inserted into another
    //     node.
    public event XmlNodeChangedEventHandler NodeInserting
    {
        add
        {
            onNodeInsertingDelegate = (XmlNodeChangedEventHandler)Delegate.Combine(onNodeInsertingDelegate, value);
        }
        remove
        {
            onNodeInsertingDelegate = (XmlNodeChangedEventHandler)Delegate.Remove(onNodeInsertingDelegate, value);
        }
    }

    //
    // Summary:
    //     Occurs when a node belonging to this document has been inserted into another
    //     node.
    public event XmlNodeChangedEventHandler NodeInserted
    {
        add
        {
            onNodeInsertedDelegate = (XmlNodeChangedEventHandler)Delegate.Combine(onNodeInsertedDelegate, value);
        }
        remove
        {
            onNodeInsertedDelegate = (XmlNodeChangedEventHandler)Delegate.Remove(onNodeInsertedDelegate, value);
        }
    }

    //
    // Summary:
    //     Occurs when a node belonging to this document is about to be removed from the
    //     document.
    public event XmlNodeChangedEventHandler NodeRemoving
    {
        add
        {
            onNodeRemovingDelegate = (XmlNodeChangedEventHandler)Delegate.Combine(onNodeRemovingDelegate, value);
        }
        remove
        {
            onNodeRemovingDelegate = (XmlNodeChangedEventHandler)Delegate.Remove(onNodeRemovingDelegate, value);
        }
    }

    //
    // Summary:
    //     Occurs when a node belonging to this document has been removed from its parent.
    public event XmlNodeChangedEventHandler NodeRemoved
    {
        add
        {
            onNodeRemovedDelegate = (XmlNodeChangedEventHandler)Delegate.Combine(onNodeRemovedDelegate, value);
        }
        remove
        {
            onNodeRemovedDelegate = (XmlNodeChangedEventHandler)Delegate.Remove(onNodeRemovedDelegate, value);
        }
    }

    //
    // Summary:
    //     Occurs when the System.Xml.XmlNode.Value of a node belonging to this document
    //     is about to be changed.
    public event XmlNodeChangedEventHandler NodeChanging
    {
        add
        {
            onNodeChangingDelegate = (XmlNodeChangedEventHandler)Delegate.Combine(onNodeChangingDelegate, value);
        }
        remove
        {
            onNodeChangingDelegate = (XmlNodeChangedEventHandler)Delegate.Remove(onNodeChangingDelegate, value);
        }
    }

    //
    // Summary:
    //     Occurs when the System.Xml.XmlNode.Value of a node belonging to this document
    //     has been changed.
    public event XmlNodeChangedEventHandler NodeChanged
    {
        add
        {
            onNodeChangedDelegate = (XmlNodeChangedEventHandler)Delegate.Combine(onNodeChangedDelegate, value);
        }
        remove
        {
            onNodeChangedDelegate = (XmlNodeChangedEventHandler)Delegate.Remove(onNodeChangedDelegate, value);
        }
    }

    //
    // Summary:
    //     Initializes a new instance of the System.Xml.XmlDocument class.
    public XmlDocument()
        : this(new XmlImplementation())
    {
    }

    //
    // Summary:
    //     Initializes a new instance of the XmlDocument class with the specified System.Xml.XmlNameTable.
    //
    //
    // Parameters:
    //   nt:
    //     The XmlNameTable to use.
    public XmlDocument(XmlNameTable nt)
        : this(new XmlImplementation(nt))
    {
    }

    //
    // Summary:
    //     Initializes a new instance of the XmlDocument class with the specified System.Xml.XmlImplementation.
    //
    //
    // Parameters:
    //   imp:
    //     The XmlImplementation to use.
    protected internal XmlDocument(XmlImplementation imp)
    {
        implementation = imp;
        domNameTable = new DomNameTable(this);
        XmlNameTable nameTable = NameTable;
        nameTable.Add(string.Empty);
        strDocumentName = nameTable.Add("#document");
        strDocumentFragmentName = nameTable.Add("#document-fragment");
        strCommentName = nameTable.Add("#comment");
        strTextName = nameTable.Add("#text");
        strCDataSectionName = nameTable.Add("#cdata-section");
        strEntityName = nameTable.Add("#entity");
        strID = nameTable.Add("id");
        strNonSignificantWhitespaceName = nameTable.Add("#whitespace");
        strSignificantWhitespaceName = nameTable.Add("#significant-whitespace");
        strXmlns = nameTable.Add("xmlns");
        strXml = nameTable.Add("xml");
        strSpace = nameTable.Add("space");
        strLang = nameTable.Add("lang");
        strReservedXmlns = nameTable.Add("http://www.w3.org/2000/xmlns/");
        strReservedXml = nameTable.Add("http://www.w3.org/XML/1998/namespace");
        strEmpty = nameTable.Add(string.Empty);
        baseURI = string.Empty;
        objLock = new object();
    }

    internal static void CheckName(string name)
    {
        int num = ValidateNames.ParseNmtoken(name, 0);
        if (num < name.Length)
        {
            throw new XmlException("Xml_BadNameChar", XmlException.BuildCharExceptionArgs(name, num));
        }
    }

    internal XmlName AddXmlName(string prefix, string localName, string namespaceURI, IXmlSchemaInfo schemaInfo)
    {
        return domNameTable.AddName(prefix, localName, namespaceURI, schemaInfo);
    }

    internal XmlName GetXmlName(string prefix, string localName, string namespaceURI, IXmlSchemaInfo schemaInfo)
    {
        return domNameTable.GetName(prefix, localName, namespaceURI, schemaInfo);
    }

    internal XmlName AddAttrXmlName(string prefix, string localName, string namespaceURI, IXmlSchemaInfo schemaInfo)
    {
        XmlName xmlName = AddXmlName(prefix, localName, namespaceURI, schemaInfo);
        if (!IsLoading)
        {
            object prefix2 = xmlName.Prefix;
            object namespaceURI2 = xmlName.NamespaceURI;
            object localName2 = xmlName.LocalName;
            if ((prefix2 == strXmlns || (prefix2 == strEmpty && localName2 == strXmlns)) ^ (namespaceURI2 == strReservedXmlns))
            {
                throw new ArgumentException(Res.GetString("Xdom_Attr_Reserved_XmlNS", namespaceURI));
            }
        }

        return xmlName;
    }

    internal bool AddIdInfo(XmlName eleName, XmlName attrName)
    {
        if (htElementIDAttrDecl == null || htElementIDAttrDecl[eleName] == null)
        {
            if (htElementIDAttrDecl == null)
            {
                htElementIDAttrDecl = new Hashtable();
            }

            htElementIDAttrDecl.Add(eleName, attrName);
            return true;
        }

        return false;
    }

    private XmlName GetIDInfoByElement_(XmlName eleName)
    {
        XmlName xmlName = GetXmlName(eleName.Prefix, eleName.LocalName, string.Empty, null);
        if (xmlName != null)
        {
            return (XmlName)htElementIDAttrDecl[xmlName];
        }

        return null;
    }

    internal XmlName GetIDInfoByElement(XmlName eleName)
    {
        if (htElementIDAttrDecl == null)
        {
            return null;
        }

        return GetIDInfoByElement_(eleName);
    }

    private WeakReference GetElement(ArrayList elementList, XmlElement elem)
    {
        ArrayList arrayList = new ArrayList();
        foreach (WeakReference element in elementList)
        {
            if (!element.IsAlive)
            {
                arrayList.Add(element);
            }
            else if ((XmlElement)element.Target == elem)
            {
                return element;
            }
        }

        foreach (WeakReference item in arrayList)
        {
            elementList.Remove(item);
        }

        return null;
    }

    internal void AddElementWithId(string id, XmlElement elem)
    {
        if (htElementIdMap == null || !htElementIdMap.Contains(id))
        {
            if (htElementIdMap == null)
            {
                htElementIdMap = new Hashtable();
            }

            ArrayList arrayList = new ArrayList();
            arrayList.Add(new WeakReference(elem));
            htElementIdMap.Add(id, arrayList);
        }
        else
        {
            ArrayList arrayList2 = (ArrayList)htElementIdMap[id];
            if (GetElement(arrayList2, elem) == null)
            {
                arrayList2.Add(new WeakReference(elem));
            }
        }
    }

    internal void RemoveElementWithId(string id, XmlElement elem)
    {
        if (htElementIdMap == null || !htElementIdMap.Contains(id))
        {
            return;
        }

        ArrayList arrayList = (ArrayList)htElementIdMap[id];
        WeakReference element = GetElement(arrayList, elem);
        if (element != null)
        {
            arrayList.Remove(element);
            if (arrayList.Count == 0)
            {
                htElementIdMap.Remove(id);
            }
        }
    }

    //
    // Summary:
    //     Creates a duplicate of this node.
    //
    // Parameters:
    //   deep:
    //     true to recursively clone the subtree under the specified node; false to clone
    //     only the node itself.
    //
    // Returns:
    //     The cloned XmlDocument node.
    public override XmlNode CloneNode(bool deep)
    {
        XmlDocument xmlDocument = Implementation.CreateDocument();
        xmlDocument.SetBaseURI(baseURI);
        if (deep)
        {
            xmlDocument.ImportChildren(this, xmlDocument, deep);
        }

        return xmlDocument;
    }

    internal XmlResolver GetResolver()
    {
        return resolver;
    }

    internal override bool IsValidChildType(XmlNodeType type)
    {
        switch (type)
        {
            case XmlNodeType.ProcessingInstruction:
            case XmlNodeType.Comment:
            case XmlNodeType.Whitespace:
            case XmlNodeType.SignificantWhitespace:
                return true;
            case XmlNodeType.DocumentType:
                if (DocumentType != null)
                {
                    throw new InvalidOperationException(Res.GetString("Xdom_DualDocumentTypeNode"));
                }

                return true;
            case XmlNodeType.Element:
                if (DocumentElement != null)
                {
                    throw new InvalidOperationException(Res.GetString("Xdom_DualDocumentElementNode"));
                }

                return true;
            case XmlNodeType.XmlDeclaration:
                if (Declaration != null)
                {
                    throw new InvalidOperationException(Res.GetString("Xdom_DualDeclarationNode"));
                }

                return true;
            default:
                return false;
        }
    }

    private bool HasNodeTypeInPrevSiblings(XmlNodeType nt, XmlNode refNode)
    {
        if (refNode == null)
        {
            return false;
        }

        XmlNode xmlNode = null;
        if (refNode.ParentNode != null)
        {
            xmlNode = refNode.ParentNode.FirstChild;
        }

        while (xmlNode != null)
        {
            if (xmlNode.NodeType == nt)
            {
                return true;
            }

            if (xmlNode == refNode)
            {
                break;
            }

            xmlNode = xmlNode.NextSibling;
        }

        return false;
    }

    private bool HasNodeTypeInNextSiblings(XmlNodeType nt, XmlNode refNode)
    {
        for (XmlNode xmlNode = refNode; xmlNode != null; xmlNode = xmlNode.NextSibling)
        {
            if (xmlNode.NodeType == nt)
            {
                return true;
            }
        }

        return false;
    }

    internal override bool CanInsertBefore(XmlNode newChild, XmlNode refChild)
    {
        if (refChild == null)
        {
            refChild = FirstChild;
        }

        if (refChild == null)
        {
            return true;
        }

        switch (newChild.NodeType)
        {
            case XmlNodeType.XmlDeclaration:
                return refChild == FirstChild;
            case XmlNodeType.ProcessingInstruction:
            case XmlNodeType.Comment:
                return refChild.NodeType != XmlNodeType.XmlDeclaration;
            case XmlNodeType.DocumentType:
                if (refChild.NodeType != XmlNodeType.XmlDeclaration)
                {
                    return !HasNodeTypeInPrevSiblings(XmlNodeType.Element, refChild.PreviousSibling);
                }

                break;
            case XmlNodeType.Element:
                if (refChild.NodeType != XmlNodeType.XmlDeclaration)
                {
                    return !HasNodeTypeInNextSiblings(XmlNodeType.DocumentType, refChild);
                }

                break;
        }

        return false;
    }

    internal override bool CanInsertAfter(XmlNode newChild, XmlNode refChild)
    {
        if (refChild == null)
        {
            refChild = LastChild;
        }

        if (refChild == null)
        {
            return true;
        }

        switch (newChild.NodeType)
        {
            case XmlNodeType.ProcessingInstruction:
            case XmlNodeType.Comment:
            case XmlNodeType.Whitespace:
            case XmlNodeType.SignificantWhitespace:
                return true;
            case XmlNodeType.DocumentType:
                return !HasNodeTypeInPrevSiblings(XmlNodeType.Element, refChild);
            case XmlNodeType.Element:
                return !HasNodeTypeInNextSiblings(XmlNodeType.DocumentType, refChild.NextSibling);
            default:
                return false;
        }
    }

    //
    // Summary:
    //     Creates an System.Xml.XmlAttribute with the specified System.Xml.XmlDocument.Name.
    //
    //
    // Parameters:
    //   name:
    //     The qualified name of the attribute. If the name contains a colon, the System.Xml.XmlNode.Prefix
    //     property reflects the part of the name preceding the first colon and the System.Xml.XmlDocument.LocalName
    //     property reflects the part of the name following the first colon. The System.Xml.XmlNode.NamespaceURI
    //     remains empty unless the prefix is a recognized built-in prefix such as xmlns.
    //     In this case NamespaceURI has a value of http://www.w3.org/2000/xmlns/.
    //
    // Returns:
    //     The new XmlAttribute.
    public XmlAttribute CreateAttribute(string name)
    {
        string prefix = string.Empty;
        string localName = string.Empty;
        string namespaceURI = string.Empty;
        XmlNode.SplitName(name, out prefix, out localName);
        SetDefaultNamespace(prefix, localName, ref namespaceURI);
        return CreateAttribute(prefix, localName, namespaceURI);
    }

    internal void SetDefaultNamespace(string prefix, string localName, ref string namespaceURI)
    {
        if (prefix == strXmlns || (prefix.Length == 0 && localName == strXmlns))
        {
            namespaceURI = strReservedXmlns;
        }
        else if (prefix == strXml)
        {
            namespaceURI = strReservedXml;
        }
    }

    //
    // Summary:
    //     Creates an System.Xml.XmlCDataSection containing the specified data.
    //
    // Parameters:
    //   data:
    //     The content of the new XmlCDataSection.
    //
    // Returns:
    //     The new XmlCDataSection.
    public virtual XmlCDataSection CreateCDataSection(string data)
    {
        fCDataNodesPresent = true;
        return new XmlCDataSection(data, this);
    }

    //
    // Summary:
    //     Creates an System.Xml.XmlComment containing the specified data.
    //
    // Parameters:
    //   data:
    //     The content of the new XmlComment.
    //
    // Returns:
    //     The new XmlComment.
    public virtual XmlComment CreateComment(string data)
    {
        return new XmlComment(data, this);
    }

    //
    // Summary:
    //     Returns a new System.Xml.XmlDocumentType object.
    //
    // Parameters:
    //   name:
    //     Name of the document type.
    //
    //   publicId:
    //     The public identifier of the document type or null. You can specify a public
    //     URI and also a system identifier to identify the location of the external DTD
    //     subset.
    //
    //   systemId:
    //     The system identifier of the document type or null. Specifies the URL of the
    //     file location for the external DTD subset.
    //
    //   internalSubset:
    //     The DTD internal subset of the document type or null.
    //
    // Returns:
    //     The new XmlDocumentType.
    [PermissionSet(SecurityAction.InheritanceDemand, Name = "FullTrust")]
    public virtual XmlDocumentType CreateDocumentType(string name, string publicId, string systemId, string internalSubset)
    {
        return new XmlDocumentType(name, publicId, systemId, internalSubset, this);
    }

    //
    // Summary:
    //     Creates an System.Xml.XmlDocumentFragment.
    //
    // Returns:
    //     The new XmlDocumentFragment.
    public virtual XmlDocumentFragment CreateDocumentFragment()
    {
        return new XmlDocumentFragment(this);
    }

    //
    // Summary:
    //     Creates an element with the specified name.
    //
    // Parameters:
    //   name:
    //     The qualified name of the element. If the name contains a colon then the System.Xml.XmlNode.Prefix
    //     property reflects the part of the name preceding the colon and the System.Xml.XmlDocument.LocalName
    //     property reflects the part of the name after the colon. The qualified name cannot
    //     include a prefix of'xmlns'.
    //
    // Returns:
    //     The new XmlElement.
    public XmlElement CreateElement(string name)
    {
        string prefix = string.Empty;
        string localName = string.Empty;
        XmlNode.SplitName(name, out prefix, out localName);
        return CreateElement(prefix, localName, string.Empty);
    }

    internal void AddDefaultAttributes(XmlElement elem)
    {
        SchemaInfo dtdSchemaInfo = DtdSchemaInfo;
        SchemaElementDecl schemaElementDecl = GetSchemaElementDecl(elem);
        if (schemaElementDecl == null || schemaElementDecl.AttDefs == null)
        {
            return;
        }

        IDictionaryEnumerator dictionaryEnumerator = schemaElementDecl.AttDefs.GetEnumerator();
        while (dictionaryEnumerator.MoveNext())
        {
            SchemaAttDef schemaAttDef = (SchemaAttDef)dictionaryEnumerator.Value;
            if (schemaAttDef.Presence == SchemaDeclBase.Use.Default || schemaAttDef.Presence == SchemaDeclBase.Use.Fixed)
            {
                string empty = string.Empty;
                string name = schemaAttDef.Name.Name;
                string attrNamespaceURI = string.Empty;
                if (dtdSchemaInfo.SchemaType == SchemaType.DTD)
                {
                    empty = schemaAttDef.Name.Namespace;
                }
                else
                {
                    empty = schemaAttDef.Prefix;
                    attrNamespaceURI = schemaAttDef.Name.Namespace;
                }

                XmlAttribute attributeNode = PrepareDefaultAttribute(schemaAttDef, empty, name, attrNamespaceURI);
                elem.SetAttributeNode(attributeNode);
            }
        }
    }

    private SchemaElementDecl GetSchemaElementDecl(XmlElement elem)
    {
        SchemaInfo dtdSchemaInfo = DtdSchemaInfo;
        if (dtdSchemaInfo != null)
        {
            XmlQualifiedName key = new XmlQualifiedName(elem.LocalName, (dtdSchemaInfo.SchemaType == SchemaType.DTD) ? elem.Prefix : elem.NamespaceURI);
            if (dtdSchemaInfo.ElementDecls.TryGetValue(key, out var value))
            {
                return value;
            }
        }

        return null;
    }

    private XmlAttribute PrepareDefaultAttribute(SchemaAttDef attdef, string attrPrefix, string attrLocalname, string attrNamespaceURI)
    {
        SetDefaultNamespace(attrPrefix, attrLocalname, ref attrNamespaceURI);
        XmlAttribute xmlAttribute = CreateDefaultAttribute(attrPrefix, attrLocalname, attrNamespaceURI);
        xmlAttribute.InnerXml = attdef.DefaultValueRaw;
        if (xmlAttribute is XmlUnspecifiedAttribute xmlUnspecifiedAttribute)
        {
            xmlUnspecifiedAttribute.SetSpecified(f: false);
        }

        return xmlAttribute;
    }

    //
    // Summary:
    //     Creates an System.Xml.XmlEntityReference with the specified name.
    //
    // Parameters:
    //   name:
    //     The name of the entity reference.
    //
    // Returns:
    //     The new XmlEntityReference.
    //
    // Exceptions:
    //   T:System.ArgumentException:
    //     The name is invalid (for example, names starting with'#' are invalid.)
    public virtual XmlEntityReference CreateEntityReference(string name)
    {
        return new XmlEntityReference(name, this);
    }

    //
    // Summary:
    //     Creates an System.Xml.XmlProcessingInstruction with the specified name and data.
    //
    //
    // Parameters:
    //   target:
    //     The name of the processing instruction.
    //
    //   data:
    //     The data for the processing instruction.
    //
    // Returns:
    //     The new XmlProcessingInstruction.
    public virtual XmlProcessingInstruction CreateProcessingInstruction(string target, string data)
    {
        return new XmlProcessingInstruction(target, data, this);
    }

    //
    // Summary:
    //     Creates an System.Xml.XmlDeclaration node with the specified values.
    //
    // Parameters:
    //   version:
    //     The version must be "1.0".
    //
    //   encoding:
    //     The value of the encoding attribute. This is the encoding that is used when you
    //     save the System.Xml.XmlDocument to a file or a stream; therefore, it must be
    //     set to a string supported by the System.Text.Encoding class, otherwise System.Xml.XmlDocument.Save(System.String)
    //     fails. If this is null or String.Empty, the Save method does not write an encoding
    //     attribute on the XML declaration and therefore the default encoding, UTF-8, is
    //     used.Note: If the XmlDocument is saved to either a System.IO.TextWriter or an
    //     System.Xml.XmlTextWriter, this encoding value is discarded. Instead, the encoding
    //     of the TextWriter or the XmlTextWriter is used. This ensures that the XML written
    //     out can be read back using the correct encoding.
    //
    //   standalone:
    //     The value must be either "yes" or "no". If this is null or String.Empty, the
    //     Save method does not write a standalone attribute on the XML declaration.
    //
    // Returns:
    //     The new XmlDeclaration node.
    //
    // Exceptions:
    //   T:System.ArgumentException:
    //     The values of version or standalone are something other than the ones specified
    //     above.
    public virtual XmlDeclaration CreateXmlDeclaration(string version, string encoding, string standalone)
    {
        return new XmlDeclaration(version, encoding, standalone, this);
    }

    //
    // Summary:
    //     Creates an System.Xml.XmlText with the specified text.
    //
    // Parameters:
    //   text:
    //     The text for the Text node.
    //
    // Returns:
    //     The new XmlText node.
    public virtual XmlText CreateTextNode(string text)
    {
        return new XmlText(text, this);
    }

    //
    // Summary:
    //     Creates an System.Xml.XmlSignificantWhitespace node.
    //
    // Parameters:
    //   text:
    //     The string must contain only the following characters &#20; &#10; &#13; and &#9;
    //
    //
    // Returns:
    //     A new XmlSignificantWhitespace node.
    public virtual XmlSignificantWhitespace CreateSignificantWhitespace(string text)
    {
        return new XmlSignificantWhitespace(text, this);
    }

    //
    // Summary:
    //     Creates a new System.Xml.XPath.XPathNavigator object for navigating this document.
    //
    //
    // Returns:
    //     An System.Xml.XPath.XPathNavigator object.
    public override XPathNavigator CreateNavigator()
    {
        return CreateNavigator(this);
    }

    //
    // Summary:
    //     Creates an System.Xml.XPath.XPathNavigator object for navigating this document
    //     positioned on the System.Xml.XmlNode specified.
    //
    // Parameters:
    //   node:
    //     The System.Xml.XmlNode you want the navigator initially positioned on.
    //
    // Returns:
    //     An System.Xml.XPath.XPathNavigator object.
    protected internal virtual XPathNavigator CreateNavigator(XmlNode node)
    {
        switch (node.NodeType)
        {
            case XmlNodeType.EntityReference:
            case XmlNodeType.Entity:
            case XmlNodeType.DocumentType:
            case XmlNodeType.Notation:
            case XmlNodeType.XmlDeclaration:
                return null;
            case XmlNodeType.Text:
            case XmlNodeType.CDATA:
            case XmlNodeType.SignificantWhitespace:
                {
                    XmlNode xmlNode = node.ParentNode;
                    if (xmlNode != null)
                    {
                        do
                        {
                            switch (xmlNode.NodeType)
                            {
                                case XmlNodeType.Attribute:
                                    return null;
                                case XmlNodeType.EntityReference:
                                    goto IL_006a;
                            }

                            break;
                        IL_006a:
                            xmlNode = xmlNode.ParentNode;
                        }
                        while (xmlNode != null);
                    }

                    node = NormalizeText(node);
                    break;
                }
            case XmlNodeType.Whitespace:
                {
                    XmlNode xmlNode = node.ParentNode;
                    if (xmlNode != null)
                    {
                        do
                        {
                            switch (xmlNode.NodeType)
                            {
                                case XmlNodeType.Attribute:
                                case XmlNodeType.Document:
                                    return null;
                                case XmlNodeType.EntityReference:
                                    goto IL_009f;
                            }

                            break;
                        IL_009f:
                            xmlNode = xmlNode.ParentNode;
                        }
                        while (xmlNode != null);
                    }

                    node = NormalizeText(node);
                    break;
                }
        }

        return new DocumentXPathNavigator(this, node);
    }

    internal static bool IsTextNode(XmlNodeType nt)
    {
        if ((uint)(nt - 3) <= 1u || (uint)(nt - 13) <= 1u)
        {
            return true;
        }

        return false;
    }

    private XmlNode NormalizeText(XmlNode n)
    {
        XmlNode xmlNode = null;
        while (IsTextNode(n.NodeType))
        {
            xmlNode = n;
            n = n.PreviousSibling;
            if (n == null)
            {
                XmlNode xmlNode2 = xmlNode;
                while (xmlNode2.ParentNode != null && xmlNode2.ParentNode.NodeType == XmlNodeType.EntityReference)
                {
                    if (xmlNode2.ParentNode.PreviousSibling != null)
                    {
                        n = xmlNode2.ParentNode.PreviousSibling;
                        break;
                    }

                    xmlNode2 = xmlNode2.ParentNode;
                    if (xmlNode2 == null)
                    {
                        break;
                    }
                }
            }

            if (n == null)
            {
                break;
            }

            while (n.NodeType == XmlNodeType.EntityReference)
            {
                n = n.LastChild;
            }
        }

        return xmlNode;
    }

    //
    // Summary:
    //     Creates an System.Xml.XmlWhitespace node.
    //
    // Parameters:
    //   text:
    //     The string must contain only the following characters &#20; &#10; &#13; and &#9;
    //
    //
    // Returns:
    //     A new XmlWhitespace node.
    public virtual XmlWhitespace CreateWhitespace(string text)
    {
        return new XmlWhitespace(text, this);
    }

    //
    // Summary:
    //     Returns an System.Xml.XmlNodeList containing a list of all descendant elements
    //     that match the specified System.Xml.XmlDocument.Name.
    //
    // Parameters:
    //   name:
    //     The qualified name to match. It is matched against the Name property of the matching
    //     node. The special value "*" matches all tags.
    //
    // Returns:
    //     An System.Xml.XmlNodeList containing a list of all matching nodes. If no nodes
    //     match name, the returned collection will be empty.
    public virtual XmlNodeList GetElementsByTagName(string name)
    {
        return new XmlElementList(this, name);
    }

    //
    // Summary:
    //     Creates an System.Xml.XmlAttribute with the specified qualified name and System.Xml.XmlNode.NamespaceURI.
    //
    //
    // Parameters:
    //   qualifiedName:
    //     The qualified name of the attribute. If the name contains a colon then the System.Xml.XmlNode.Prefix
    //     property will reflect the part of the name preceding the colon and the System.Xml.XmlDocument.LocalName
    //     property will reflect the part of the name after the colon.
    //
    //   namespaceURI:
    //     The namespaceURI of the attribute. If the qualified name includes a prefix of
    //     xmlns, then this parameter must be http://www.w3.org/2000/xmlns/.
    //
    // Returns:
    //     The new XmlAttribute.
    public XmlAttribute CreateAttribute(string qualifiedName, string namespaceURI)
    {
        string prefix = string.Empty;
        string localName = string.Empty;
        XmlNode.SplitName(qualifiedName, out prefix, out localName);
        return CreateAttribute(prefix, localName, namespaceURI);
    }

    //
    // Summary:
    //     Creates an System.Xml.XmlElement with the qualified name and System.Xml.XmlNode.NamespaceURI.
    //
    //
    // Parameters:
    //   qualifiedName:
    //     The qualified name of the element. If the name contains a colon then the System.Xml.XmlNode.Prefix
    //     property will reflect the part of the name preceding the colon and the System.Xml.XmlDocument.LocalName
    //     property will reflect the part of the name after the colon. The qualified name
    //     cannot include a prefix of'xmlns'.
    //
    //   namespaceURI:
    //     The namespace URI of the element.
    //
    // Returns:
    //     The new XmlElement.
    public XmlElement CreateElement(string qualifiedName, string namespaceURI)
    {
        string prefix = string.Empty;
        string localName = string.Empty;
        XmlNode.SplitName(qualifiedName, out prefix, out localName);
        return CreateElement(prefix, localName, namespaceURI);
    }

    //
    // Summary:
    //     Returns an System.Xml.XmlNodeList containing a list of all descendant elements
    //     that match the specified System.Xml.XmlDocument.LocalName and System.Xml.XmlNode.NamespaceURI.
    //
    //
    // Parameters:
    //   localName:
    //     The LocalName to match. The special value "*" matches all tags.
    //
    //   namespaceURI:
    //     NamespaceURI to match.
    //
    // Returns:
    //     An System.Xml.XmlNodeList containing a list of all matching nodes. If no nodes
    //     match the specified localName and namespaceURI, the returned collection will
    //     be empty.
    public virtual XmlNodeList GetElementsByTagName(string localName, string namespaceURI)
    {
        return new XmlElementList(this, localName, namespaceURI);
    }

    //
    // Summary:
    //     Gets the System.Xml.XmlElement with the specified ID.
    //
    // Parameters:
    //   elementId:
    //     The attribute ID to match.
    //
    // Returns:
    //     The XmlElement with the matching ID or null if no matching element is found.
    public virtual XmlElement GetElementById(string elementId)
    {
        if (htElementIdMap != null)
        {
            ArrayList arrayList = (ArrayList)htElementIdMap[elementId];
            if (arrayList != null)
            {
                foreach (WeakReference item in arrayList)
                {
                    XmlElement xmlElement = (XmlElement)item.Target;
                    if (xmlElement != null && xmlElement.IsConnected())
                    {
                        return xmlElement;
                    }
                }
            }
        }

        return null;
    }

    //
    // Summary:
    //     Imports a node from another document to the current document.
    //
    // Parameters:
    //   node:
    //     The node being imported.
    //
    //   deep:
    //     true to perform a deep clone; otherwise, false.
    //
    // Returns:
    //     The imported System.Xml.XmlNode.
    //
    // Exceptions:
    //   T:System.InvalidOperationException:
    //     Calling this method on a node type which cannot be imported.
    public virtual XmlNode ImportNode(XmlNode node, bool deep)
    {
        return ImportNodeInternal(node, deep);
    }

    private XmlNode ImportNodeInternal(XmlNode node, bool deep)
    {
        XmlNode xmlNode = null;
        if (node == null)
        {
            throw new InvalidOperationException(Res.GetString("Xdom_Import_NullNode"));
        }

        switch (node.NodeType)
        {
            case XmlNodeType.Element:
                xmlNode = CreateElement(node.Prefix, node.LocalName, node.NamespaceURI);
                ImportAttributes(node, xmlNode);
                if (deep)
                {
                    ImportChildren(node, xmlNode, deep);
                }

                break;
            case XmlNodeType.Attribute:
                xmlNode = CreateAttribute(node.Prefix, node.LocalName, node.NamespaceURI);
                ImportChildren(node, xmlNode, deep: true);
                break;
            case XmlNodeType.Text:
                xmlNode = CreateTextNode(node.Value);
                break;
            case XmlNodeType.Comment:
                xmlNode = CreateComment(node.Value);
                break;
            case XmlNodeType.ProcessingInstruction:
                xmlNode = CreateProcessingInstruction(node.Name, node.Value);
                break;
            case XmlNodeType.XmlDeclaration:
                {
                    XmlDeclaration xmlDeclaration = (XmlDeclaration)node;
                    xmlNode = CreateXmlDeclaration(xmlDeclaration.Version, xmlDeclaration.Encoding, xmlDeclaration.Standalone);
                    break;
                }
            case XmlNodeType.CDATA:
                xmlNode = CreateCDataSection(node.Value);
                break;
            case XmlNodeType.DocumentType:
                {
                    XmlDocumentType xmlDocumentType = (XmlDocumentType)node;
                    xmlNode = CreateDocumentType(xmlDocumentType.Name, xmlDocumentType.PublicId, xmlDocumentType.SystemId, xmlDocumentType.InternalSubset);
                    break;
                }
            case XmlNodeType.DocumentFragment:
                xmlNode = CreateDocumentFragment();
                if (deep)
                {
                    ImportChildren(node, xmlNode, deep);
                }

                break;
            case XmlNodeType.EntityReference:
                xmlNode = CreateEntityReference(node.Name);
                break;
            case XmlNodeType.Whitespace:
                xmlNode = CreateWhitespace(node.Value);
                break;
            case XmlNodeType.SignificantWhitespace:
                xmlNode = CreateSignificantWhitespace(node.Value);
                break;
            default:
                throw new InvalidOperationException(string.Format(CultureInfo.InvariantCulture, Res.GetString("Xdom_Import"), node.NodeType.ToString()));
        }

        return xmlNode;
    }

    private void ImportAttributes(XmlNode fromElem, XmlNode toElem)
    {
        int count = fromElem.Attributes.Count;
        for (int i = 0; i < count; i++)
        {
            if (fromElem.Attributes[i].Specified)
            {
                toElem.Attributes.SetNamedItem(ImportNodeInternal(fromElem.Attributes[i], deep: true));
            }
        }
    }

    private void ImportChildren(XmlNode fromNode, XmlNode toNode, bool deep)
    {
        for (XmlNode xmlNode = fromNode.FirstChild; xmlNode != null; xmlNode = xmlNode.NextSibling)
        {
            toNode.AppendChild(ImportNodeInternal(xmlNode, deep));
        }
    }

    //
    // Summary:
    //     Creates an System.Xml.XmlAttribute with the specified System.Xml.XmlNode.Prefix,
    //     System.Xml.XmlDocument.LocalName, and System.Xml.XmlNode.NamespaceURI.
    //
    // Parameters:
    //   prefix:
    //     The prefix of the attribute (if any). String.Empty and null are equivalent.
    //
    //   localName:
    //     The local name of the attribute.
    //
    //   namespaceURI:
    //     The namespace URI of the attribute (if any). String.Empty and null are equivalent.
    //     If prefix is xmlns, then this parameter must be http://www.w3.org/2000/xmlns/;
    //     otherwise an exception is thrown.
    //
    // Returns:
    //     The new XmlAttribute.
    public virtual XmlAttribute CreateAttribute(string prefix, string localName, string namespaceURI)
    {
        return new XmlAttribute(AddAttrXmlName(prefix, localName, namespaceURI, null), this);
    }

    //
    // Summary:
    //     Creates a default attribute with the specified prefix, local name and namespace
    //     URI.
    //
    // Parameters:
    //   prefix:
    //     The prefix of the attribute (if any).
    //
    //   localName:
    //     The local name of the attribute.
    //
    //   namespaceURI:
    //     The namespace URI of the attribute (if any).
    //
    // Returns:
    //     The new System.Xml.XmlAttribute.
    protected internal virtual XmlAttribute CreateDefaultAttribute(string prefix, string localName, string namespaceURI)
    {
        return new XmlUnspecifiedAttribute(prefix, localName, namespaceURI, this);
    }

    //
    // Summary:
    //     Creates an element with the specified System.Xml.XmlNode.Prefix, System.Xml.XmlDocument.LocalName,
    //     and System.Xml.XmlNode.NamespaceURI.
    //
    // Parameters:
    //   prefix:
    //     The prefix of the new element (if any). String.Empty and null are equivalent.
    //
    //
    //   localName:
    //     The local name of the new element.
    //
    //   namespaceURI:
    //     The namespace URI of the new element (if any). String.Empty and null are equivalent.
    //
    //
    // Returns:
    //     The new System.Xml.XmlElement.
    public virtual XmlElement CreateElement(string prefix, string localName, string namespaceURI)
    {
        XmlElement xmlElement = new XmlElement(AddXmlName(prefix, localName, namespaceURI, null), empty: true, this);
        if (!IsLoading)
        {
            AddDefaultAttributes(xmlElement);
        }

        return xmlElement;
    }

    //
    // Summary:
    //     Creates a System.Xml.XmlNode with the specified System.Xml.XmlNodeType, System.Xml.XmlNode.Prefix,
    //     System.Xml.XmlDocument.Name, and System.Xml.XmlNode.NamespaceURI.
    //
    // Parameters:
    //   type:
    //     The XmlNodeType of the new node.
    //
    //   prefix:
    //     The prefix of the new node.
    //
    //   name:
    //     The local name of the new node.
    //
    //   namespaceURI:
    //     The namespace URI of the new node.
    //
    // Returns:
    //     The new XmlNode.
    //
    // Exceptions:
    //   T:System.ArgumentException:
    //     The name was not provided and the XmlNodeType requires a name.
    public virtual XmlNode CreateNode(XmlNodeType type, string prefix, string name, string namespaceURI)
    {
        switch (type)
        {
            case XmlNodeType.Element:
                if (prefix != null)
                {
                    return CreateElement(prefix, name, namespaceURI);
                }

                return CreateElement(name, namespaceURI);
            case XmlNodeType.Attribute:
                if (prefix != null)
                {
                    return CreateAttribute(prefix, name, namespaceURI);
                }

                return CreateAttribute(name, namespaceURI);
            case XmlNodeType.Text:
                return CreateTextNode(string.Empty);
            case XmlNodeType.CDATA:
                return CreateCDataSection(string.Empty);
            case XmlNodeType.EntityReference:
                return CreateEntityReference(name);
            case XmlNodeType.ProcessingInstruction:
                return CreateProcessingInstruction(name, string.Empty);
            case XmlNodeType.XmlDeclaration:
                return CreateXmlDeclaration("1.0", null, null);
            case XmlNodeType.Comment:
                return CreateComment(string.Empty);
            case XmlNodeType.DocumentFragment:
                return CreateDocumentFragment();
            case XmlNodeType.DocumentType:
                return CreateDocumentType(name, string.Empty, string.Empty, string.Empty);
            case XmlNodeType.Document:
                return new XmlDocument();
            case XmlNodeType.SignificantWhitespace:
                return CreateSignificantWhitespace(string.Empty);
            case XmlNodeType.Whitespace:
                return CreateWhitespace(string.Empty);
            default:
                throw new ArgumentException(Res.GetString("Arg_CannotCreateNode", type));
        }
    }

    //
    // Summary:
    //     Creates an System.Xml.XmlNode with the specified node type, System.Xml.XmlDocument.Name,
    //     and System.Xml.XmlNode.NamespaceURI.
    //
    // Parameters:
    //   nodeTypeString:
    //     String version of the System.Xml.XmlNodeType of the new node. This parameter
    //     must be one of the values listed in the table below.
    //
    //   name:
    //     The qualified name of the new node. If the name contains a colon, it is parsed
    //     into System.Xml.XmlNode.Prefix and System.Xml.XmlDocument.LocalName components.
    //
    //
    //   namespaceURI:
    //     The namespace URI of the new node.
    //
    // Returns:
    //     The new XmlNode.
    //
    // Exceptions:
    //   T:System.ArgumentException:
    //     The name was not provided and the XmlNodeType requires a name; or nodeTypeString
    //     is not one of the strings listed below.
    public virtual XmlNode CreateNode(string nodeTypeString, string name, string namespaceURI)
    {
        return CreateNode(ConvertToNodeType(nodeTypeString), name, namespaceURI);
    }

    //
    // Summary:
    //     Creates an System.Xml.XmlNode with the specified System.Xml.XmlNodeType, System.Xml.XmlDocument.Name,
    //     and System.Xml.XmlNode.NamespaceURI.
    //
    // Parameters:
    //   type:
    //     The XmlNodeType of the new node.
    //
    //   name:
    //     The qualified name of the new node. If the name contains a colon then it is parsed
    //     into System.Xml.XmlNode.Prefix and System.Xml.XmlDocument.LocalName components.
    //
    //
    //   namespaceURI:
    //     The namespace URI of the new node.
    //
    // Returns:
    //     The new XmlNode.
    //
    // Exceptions:
    //   T:System.ArgumentException:
    //     The name was not provided and the XmlNodeType requires a name.
    public virtual XmlNode CreateNode(XmlNodeType type, string name, string namespaceURI)
    {
        return CreateNode(type, null, name, namespaceURI);
    }

    //
    // Summary:
    //     Creates an System.Xml.XmlNode object based on the information in the System.Xml.XmlReader.
    //     The reader must be positioned on a node or attribute.
    //
    // Parameters:
    //   reader:
    //     The XML source
    //
    // Returns:
    //     The new XmlNode or null if no more nodes exist.
    //
    // Exceptions:
    //   T:System.NullReferenceException:
    //     The reader is positioned on a node type that does not translate to a valid DOM
    //     node (for example, EndElement or EndEntity).
    [PermissionSet(SecurityAction.InheritanceDemand, Name = "FullTrust")]
    public virtual XmlNode ReadNode(XmlReader reader)
    {
        XmlNode xmlNode = null;
        try
        {
            IsLoading = true;
            XmlLoader xmlLoader = new XmlLoader();
            return xmlLoader.ReadCurrentNode(this, reader);
        }
        finally
        {
            IsLoading = false;
        }
    }

    internal XmlNodeType ConvertToNodeType(string nodeTypeString)
    {
        return nodeTypeString switch
        {
            "element" => XmlNodeType.Element,
            "attribute" => XmlNodeType.Attribute,
            "text" => XmlNodeType.Text,
            "cdatasection" => XmlNodeType.CDATA,
            "entityreference" => XmlNodeType.EntityReference,
            "entity" => XmlNodeType.Entity,
            "processinginstruction" => XmlNodeType.ProcessingInstruction,
            "comment" => XmlNodeType.Comment,
            "document" => XmlNodeType.Document,
            "documenttype" => XmlNodeType.DocumentType,
            "documentfragment" => XmlNodeType.DocumentFragment,
            "notation" => XmlNodeType.Notation,
            "significantwhitespace" => XmlNodeType.SignificantWhitespace,
            "whitespace" => XmlNodeType.Whitespace,
            _ => throw new ArgumentException(Res.GetString("Xdom_Invalid_NT_String", nodeTypeString)),
        };
    }

    private XmlTextReader SetupReader(XmlTextReader tr)
    {
        tr.XmlValidatingReaderCompatibilityMode = true;
        tr.EntityHandling = EntityHandling.ExpandCharEntities;
        if (HasSetResolver)
        {
            tr.XmlResolver = GetResolver();
        }

        return tr;
    }

    //
    // Summary:
    //     Loads the XML document from the specified URL.
    //
    // Parameters:
    //   filename:
    //     URL for the file containing the XML document to load. The URL can be either a
    //     local file or an HTTP URL (a Web address).
    //
    // Exceptions:
    //   T:System.Xml.XmlException:
    //     There is a load or parse error in the XML. In this case, a System.IO.FileNotFoundException
    //     is raised.
    //
    //   T:System.ArgumentException:
    //     filename is a zero-length string, contains only white space, or contains one
    //     or more invalid characters as defined by System.IO.Path.InvalidPathChars.
    //
    //   T:System.ArgumentNullException:
    //     filename is null.
    //
    //   T:System.IO.PathTooLongException:
    //     The specified path, file name, or both exceed the system-defined maximum length.
    //     For example, on Windows-based platforms, paths must be less than 248 characters,
    //     and file names must be less than 260 characters.
    //
    //   T:System.IO.DirectoryNotFoundException:
    //     The specified path is invalid (for example, it is on an unmapped drive).
    //
    //   T:System.IO.IOException:
    //     An I/O error occurred while opening the file.
    //
    //   T:System.UnauthorizedAccessException:
    //     filename specified a file that is read-only.-or- This operation is not supported
    //     on the current platform.-or- filename specified a directory.-or- The caller does
    //     not have the required permission.
    //
    //   T:System.IO.FileNotFoundException:
    //     The file specified in filename was not found.
    //
    //   T:System.NotSupportedException:
    //     filename is in an invalid format.
    //
    //   T:System.Security.SecurityException:
    //     The caller does not have the required permission.
    public virtual void Load(string filename)
    {
        XmlTextReader xmlTextReader = SetupReader(new XmlTextReader(filename, NameTable));
        try
        {
            Load(xmlTextReader);
        }
        finally
        {
            xmlTextReader.Close();
        }
    }

    //
    // Summary:
    //     Loads the XML document from the specified stream.
    //
    // Parameters:
    //   inStream:
    //     The stream containing the XML document to load.
    //
    // Exceptions:
    //   T:System.Xml.XmlException:
    //     There is a load or parse error in the XML. In this case, a System.IO.FileNotFoundException
    //     is raised.
    public virtual void Load(Stream inStream)
    {
        XmlTextReader xmlTextReader = SetupReader(new XmlTextReader(inStream, NameTable));
        try
        {
            Load(xmlTextReader);
        }
        finally
        {
            xmlTextReader.Impl.Close(closeInput: false);
        }
    }

    //
    // Summary:
    //     Loads the XML document from the specified System.IO.TextReader.
    //
    // Parameters:
    //   txtReader:
    //     The TextReader used to feed the XML data into the document.
    //
    // Exceptions:
    //   T:System.Xml.XmlException:
    //     There is a load or parse error in the XML. In this case, the document remains
    //     empty.
    public virtual void Load(TextReader txtReader)
    {
        XmlTextReader xmlTextReader = SetupReader(new XmlTextReader(txtReader, NameTable));
        try
        {
            Load(xmlTextReader);
        }
        finally
        {
            xmlTextReader.Impl.Close(closeInput: false);
        }
    }

    //
    // Summary:
    //     Loads the XML document from the specified System.Xml.XmlReader.
    //
    // Parameters:
    //   reader:
    //     The XmlReader used to feed the XML data into the document.
    //
    // Exceptions:
    //   T:System.Xml.XmlException:
    //     There is a load or parse error in the XML. In this case, the document remains
    //     empty.
    public virtual void Load(XmlReader reader)
    {
        try
        {
            IsLoading = true;
            actualLoadingStatus = true;
            RemoveAll();
            fEntRefNodesPresent = false;
            fCDataNodesPresent = false;
            reportValidity = true;
            XmlLoader xmlLoader = new XmlLoader();
            xmlLoader.Load(this, reader, preserveWhitespace);
        }
        finally
        {
            IsLoading = false;
            actualLoadingStatus = false;
            reportValidity = true;
        }
    }

    //
    // Summary:
    //     Loads the XML document from the specified string.
    //
    // Parameters:
    //   xml:
    //     String containing the XML document to load.
    //
    // Exceptions:
    //   T:System.Xml.XmlException:
    //     There is a load or parse error in the XML. In this case, the document remains
    //     empty.
    public virtual void LoadXml(string xml)
    {
        XmlTextReader xmlTextReader = SetupReader(new XmlTextReader(new StringReader(xml), NameTable));
        try
        {
            Load(xmlTextReader);
        }
        finally
        {
            xmlTextReader.Close();
        }
    }

    //
    // Summary:
    //     Saves the XML document to the specified file. If the specified file exists, this
    //     method overwrites it.
    //
    // Parameters:
    //   filename:
    //     The location of the file where you want to save the document.
    //
    // Exceptions:
    //   T:System.Xml.XmlException:
    //     The operation would not result in a well formed XML document (for example, no
    //     document element or duplicate XML declarations).
    public virtual void Save(string filename)
    {
        if (DocumentElement == null)
        {
            throw new XmlException("Xml_InvalidXmlDocument", Res.GetString("Xdom_NoRootEle"));
        }

        XmlDOMTextWriter xmlDOMTextWriter = new XmlDOMTextWriter(filename, TextEncoding);
        try
        {
            if (!preserveWhitespace)
            {
                xmlDOMTextWriter.Formatting = Formatting.Indented;
            }

            WriteTo(xmlDOMTextWriter);
            xmlDOMTextWriter.Flush();
        }
        finally
        {
            xmlDOMTextWriter.Close();
        }
    }

    //
    // Summary:
    //     Saves the XML document to the specified stream.
    //
    // Parameters:
    //   outStream:
    //     The stream to which you want to save.
    //
    // Exceptions:
    //   T:System.Xml.XmlException:
    //     The operation would not result in a well formed XML document (for example, no
    //     document element or duplicate XML declarations).
    public virtual void Save(Stream outStream)
    {
        XmlDOMTextWriter xmlDOMTextWriter = new XmlDOMTextWriter(outStream, TextEncoding);
        if (!preserveWhitespace)
        {
            xmlDOMTextWriter.Formatting = Formatting.Indented;
        }

        WriteTo(xmlDOMTextWriter);
        xmlDOMTextWriter.Flush();
    }

    //
    // Summary:
    //     Saves the XML document to the specified System.IO.TextWriter.
    //
    // Parameters:
    //   writer:
    //     The TextWriter to which you want to save.
    //
    // Exceptions:
    //   T:System.Xml.XmlException:
    //     The operation would not result in a well formed XML document (for example, no
    //     document element or duplicate XML declarations).
    public virtual void Save(TextWriter writer)
    {
        XmlDOMTextWriter xmlDOMTextWriter = new XmlDOMTextWriter(writer);
        if (!preserveWhitespace)
        {
            xmlDOMTextWriter.Formatting = Formatting.Indented;
        }

        Save(xmlDOMTextWriter);
    }

    //
    // Summary:
    //     Saves the XML document to the specified System.Xml.XmlWriter.
    //
    // Parameters:
    //   w:
    //     The XmlWriter to which you want to save.
    //
    // Exceptions:
    //   T:System.Xml.XmlException:
    //     The operation would not result in a well formed XML document (for example, no
    //     document element or duplicate XML declarations).
    public virtual void Save(XmlWriter w)
    {
        XmlNode xmlNode = FirstChild;
        if (xmlNode == null)
        {
            return;
        }

        if (w.WriteState == WriteState.Start)
        {
            if (xmlNode is XmlDeclaration)
            {
                if (Standalone.Length == 0)
                {
                    w.WriteStartDocument();
                }
                else if (Standalone == "yes")
                {
                    w.WriteStartDocument(standalone: true);
                }
                else if (Standalone == "no")
                {
                    w.WriteStartDocument(standalone: false);
                }

                xmlNode = xmlNode.NextSibling;
            }
            else
            {
                w.WriteStartDocument();
            }
        }

        while (xmlNode != null)
        {
            xmlNode.WriteTo(w);
            xmlNode = xmlNode.NextSibling;
        }

        w.Flush();
    }

    //
    // Summary:
    //     Saves the XmlDocument node to the specified System.Xml.XmlWriter.
    //
    // Parameters:
    //   w:
    //     The XmlWriter to which you want to save.
    public override void WriteTo(XmlWriter w)
    {
        WriteContentTo(w);
    }

    //
    // Summary:
    //     Saves all the children of the XmlDocument node to the specified System.Xml.XmlWriter.
    //
    //
    // Parameters:
    //   xw:
    //     The XmlWriter to which you want to save.
    public override void WriteContentTo(XmlWriter xw)
    {
        IEnumerator enumerator = GetEnumerator();
        try
        {
            while (enumerator.MoveNext())
            {
                XmlNode xmlNode = (XmlNode)enumerator.Current;
                xmlNode.WriteTo(xw);
            }
        }
        finally
        {
            IDisposable disposable = enumerator as IDisposable;
            if (disposable != null)
            {
                disposable.Dispose();
            }
        }
    }

    //
    // Summary:
    //     Validates the System.Xml.XmlDocument against the XML Schema Definition Language
    //     (XSD) schemas contained in the System.Xml.XmlDocument.Schemas property.
    //
    // Parameters:
    //   validationEventHandler:
    //     The System.Xml.Schema.ValidationEventHandler object that receives information
    //     about schema validation warnings and errors.
    //
    // Exceptions:
    //   T:System.Xml.Schema.XmlSchemaValidationException:
    //     A schema validation event occurred and no System.Xml.Schema.ValidationEventHandler
    //     object was specified.
    public void Validate(ValidationEventHandler validationEventHandler)
    {
        Validate(validationEventHandler, this);
    }

    //
    // Summary:
    //     Validates the System.Xml.XmlNode object specified against the XML Schema Definition
    //     Language (XSD) schemas in the System.Xml.XmlDocument.Schemas property.
    //
    // Parameters:
    //   validationEventHandler:
    //     The System.Xml.Schema.ValidationEventHandler object that receives information
    //     about schema validation warnings and errors.
    //
    //   nodeToValidate:
    //     The System.Xml.XmlNode object created from an System.Xml.XmlDocument to validate.
    //
    //
    // Exceptions:
    //   T:System.ArgumentException:
    //     The System.Xml.XmlNode object parameter was not created from an System.Xml.XmlDocument.
    //
    //
    //   T:System.InvalidOperationException:
    //     The System.Xml.XmlNode object parameter is not an element, attribute, document
    //     fragment, or the root node.
    //
    //   T:System.Xml.Schema.XmlSchemaValidationException:
    //     A schema validation event occurred and no System.Xml.Schema.ValidationEventHandler
    //     object was specified.
    public void Validate(ValidationEventHandler validationEventHandler, XmlNode nodeToValidate)
    {
        if (schemas == null || schemas.Count == 0)
        {
            throw new InvalidOperationException(Res.GetString("XmlDocument_NoSchemaInfo"));
        }

        XmlDocument document = nodeToValidate.Document;
        if (document != this)
        {
            throw new ArgumentException(Res.GetString("XmlDocument_NodeNotFromDocument", "nodeToValidate"));
        }

        if (nodeToValidate == this)
        {
            reportValidity = false;
        }

        DocumentSchemaValidator documentSchemaValidator = new DocumentSchemaValidator(this, schemas, validationEventHandler);
        documentSchemaValidator.Validate(nodeToValidate);
        if (nodeToValidate == this)
        {
            reportValidity = true;
        }
    }

    internal override XmlNodeChangedEventArgs GetEventArgs(XmlNode node, XmlNode oldParent, XmlNode newParent, string oldValue, string newValue, XmlNodeChangedAction action)
    {
        reportValidity = false;
        switch (action)
        {
            case XmlNodeChangedAction.Insert:
                if (onNodeInsertingDelegate == null && onNodeInsertedDelegate == null)
                {
                    return null;
                }

                break;
            case XmlNodeChangedAction.Remove:
                if (onNodeRemovingDelegate == null && onNodeRemovedDelegate == null)
                {
                    return null;
                }

                break;
            case XmlNodeChangedAction.Change:
                if (onNodeChangingDelegate == null && onNodeChangedDelegate == null)
                {
                    return null;
                }

                break;
        }

        return new XmlNodeChangedEventArgs(node, oldParent, newParent, oldValue, newValue, action);
    }

    internal XmlNodeChangedEventArgs GetInsertEventArgsForLoad(XmlNode node, XmlNode newParent)
    {
        if (onNodeInsertingDelegate == null && onNodeInsertedDelegate == null)
        {
            return null;
        }

        string value = node.Value;
        return new XmlNodeChangedEventArgs(node, null, newParent, value, value, XmlNodeChangedAction.Insert);
    }

    internal override void BeforeEvent(XmlNodeChangedEventArgs args)
    {
        if (args == null)
        {
            return;
        }

        switch (args.Action)
        {
            case XmlNodeChangedAction.Insert:
                if (onNodeInsertingDelegate != null)
                {
                    onNodeInsertingDelegate(this, args);
                }

                break;
            case XmlNodeChangedAction.Remove:
                if (onNodeRemovingDelegate != null)
                {
                    onNodeRemovingDelegate(this, args);
                }

                break;
            case XmlNodeChangedAction.Change:
                if (onNodeChangingDelegate != null)
                {
                    onNodeChangingDelegate(this, args);
                }

                break;
        }
    }

    internal override void AfterEvent(XmlNodeChangedEventArgs args)
    {
        if (args == null)
        {
            return;
        }

        switch (args.Action)
        {
            case XmlNodeChangedAction.Insert:
                if (onNodeInsertedDelegate != null)
                {
                    onNodeInsertedDelegate(this, args);
                }

                break;
            case XmlNodeChangedAction.Remove:
                if (onNodeRemovedDelegate != null)
                {
                    onNodeRemovedDelegate(this, args);
                }

                break;
            case XmlNodeChangedAction.Change:
                if (onNodeChangedDelegate != null)
                {
                    onNodeChangedDelegate(this, args);
                }

                break;
        }
    }

    internal XmlAttribute GetDefaultAttribute(XmlElement elem, string attrPrefix, string attrLocalname, string attrNamespaceURI)
    {
        SchemaInfo dtdSchemaInfo = DtdSchemaInfo;
        SchemaElementDecl schemaElementDecl = GetSchemaElementDecl(elem);
        if (schemaElementDecl != null && schemaElementDecl.AttDefs != null)
        {
            IDictionaryEnumerator dictionaryEnumerator = schemaElementDecl.AttDefs.GetEnumerator();
            while (dictionaryEnumerator.MoveNext())
            {
                SchemaAttDef schemaAttDef = (SchemaAttDef)dictionaryEnumerator.Value;
                if ((schemaAttDef.Presence == SchemaDeclBase.Use.Default || schemaAttDef.Presence == SchemaDeclBase.Use.Fixed) && schemaAttDef.Name.Name == attrLocalname && ((dtdSchemaInfo.SchemaType == SchemaType.DTD && schemaAttDef.Name.Namespace == attrPrefix) || (dtdSchemaInfo.SchemaType != SchemaType.DTD && schemaAttDef.Name.Namespace == attrNamespaceURI)))
                {
                    return PrepareDefaultAttribute(schemaAttDef, attrPrefix, attrLocalname, attrNamespaceURI);
                }
            }
        }

        return null;
    }

    internal XmlEntity GetEntityNode(string name)
    {
        if (DocumentType != null)
        {
            XmlNamedNodeMap xmlNamedNodeMap = DocumentType.Entities;
            if (xmlNamedNodeMap != null)
            {
                return (XmlEntity)xmlNamedNodeMap.GetNamedItem(name);
            }
        }

        return null;
    }

    internal void SetBaseURI(string inBaseURI)
    {
        baseURI = inBaseURI;
    }

    internal override XmlNode AppendChildForLoad(XmlNode newChild, XmlDocument doc)
    {
        if (!IsValidChildType(newChild.NodeType))
        {
            throw new InvalidOperationException(Res.GetString("Xdom_Node_Insert_TypeConflict"));
        }

        if (!CanInsertAfter(newChild, LastChild))
        {
            throw new InvalidOperationException(Res.GetString("Xdom_Node_Insert_Location"));
        }

        XmlNodeChangedEventArgs insertEventArgsForLoad = GetInsertEventArgsForLoad(newChild, this);
        if (insertEventArgsForLoad != null)
        {
            BeforeEvent(insertEventArgsForLoad);
        }

        XmlLinkedNode xmlLinkedNode = (XmlLinkedNode)newChild;
        if (lastChild == null)
        {
            xmlLinkedNode.next = xmlLinkedNode;
        }
        else
        {
            xmlLinkedNode.next = lastChild.next;
            lastChild.next = xmlLinkedNode;
        }

        lastChild = xmlLinkedNode;
        xmlLinkedNode.SetParentForLoad(this);
        if (insertEventArgsForLoad != null)
        {
            AfterEvent(insertEventArgsForLoad);
        }

        return xmlLinkedNode;
    }
}
#if false // Decompilation log
'25' items in cache
------------------
Resolve: 'mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
Found single assembly: 'mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
Load from: 'C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.5\mscorlib.dll'
------------------
Resolve: 'System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
Found single assembly: 'System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
Load from: 'C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.5\System.dll'
------------------
Resolve: 'System.Data.SqlXml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
Could not find by name: 'System.Data.SqlXml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
------------------
Resolve: 'System.Configuration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a'
Found single assembly: 'System.Configuration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a'
Load from: 'C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.5\System.Configuration.dll'
#endif

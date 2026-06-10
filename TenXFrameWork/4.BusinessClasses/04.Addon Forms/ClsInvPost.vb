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

Public Class ClsInvPost
    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix

    Dim objutilities As Utilities
    Public rs, RsNum As SAPbobsCOM.Recordset

    'Form Creation
    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("Invoice_Posting.xml", "INV_P", ResourceType.Embeded) 'Loading Form.xml file
            objForm = objMain.objApplication.Forms.GetForm("INV_P", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_IP")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_IP_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific
            Me.SetDefault(objForm.UniqueID)


            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_IP")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_IP_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific
            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNX_IPUDO"))

            Me.SetNewLine(objForm.UniqueID)


            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "2", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objMatrix.AutoResizeColumns()
            'Dim DocType As SAPbouiCOM.ComboBox = objForm.Items.Item("DocType").Specific
            'DocType.Select(388, SAPbouiCOM.BoSearchKey.psk_ByValue)
            'Dim Status As SAPbouiCOM.ComboBox = objForm.Items.Item("Status").Specific
            'Status.Select(1, SAPbouiCOM.BoSearchKey.psk_ByValue)

            objForm.Freeze(False)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "Inv_Posting" And pVal.BeforeAction = False Then
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
                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_IP")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_IP_C0")
                    Dim objMtx As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific
                    If pVal.ColUID = "DocEntry" And pVal.BeforeAction = True Then
                        Try

                            Dim DocTypeStr As String = oDBs_Head.GetValue("U_DocType", 0)
                            If DocTypeStr = "" Then
                                DocTypeStr = "388"
                            End If
                            If DocTypeStr = "381" Then
                                BubbleEvent = False
                                Try
                                    objMain.objApplication.ActivateMenuItem("2055")
                                    Dim objTargetForm As SAPbouiCOM.Form = objMain.objApplication.Forms.ActiveForm
                                    objTargetForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                    objTargetForm.Items.Item("8").Specific.Value = objMtx.Columns.Item("DNum").Cells.Item(pVal.Row).Specific.Value.ToString()
                                    objTargetForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Catch ex As Exception

                                End Try
                            ElseIf DocTypeStr = "383" Then
                                BubbleEvent = False
                                Try
                                    objMain.objApplication.ActivateMenuItem("2071")
                                    Dim objTargetForm As SAPbouiCOM.Form = objMain.objApplication.Forms.ActiveForm
                                    objTargetForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                    objTargetForm.Items.Item("8").Specific.Value = objMtx.Columns.Item("DNum").Cells.Item(pVal.Row).Specific.Value.ToString()
                                    objTargetForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Catch ex As Exception

                                End Try
                            End If
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_IP")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_IP_C0")
                    Dim objMtx As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific
                    If pVal.ItemUID = "3" And pVal.BeforeAction = False Then
                        Try
                            Dim DocType As String = oDBs_Head.GetValue("U_DocType", 0)
                            For i As Integer = 1 To objMtx.VisualRowCount

                                Dim DocEntry As String = objMtx.Columns.Item("DNum").Cells.Item(i).Specific.Value.ToString()
                                Dim checkbox As SAPbouiCOM.CheckBox = objMtx.Columns.Item("Select").Cells.Item(i).Specific
                                If checkbox.Checked = True Then
                                    If DocType = "381" Then
                                        Me.FlickCreditNotePosting(DocEntry, DocType)
                                    Else
                                        Me.FlickINVPosting(DocEntry, DocType)
                                    End If

                                End If
                            Next
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If
                    If pVal.ItemUID = "4" And pVal.BeforeAction = True Then
                        Dim FromDate As String = objForm.Items.Item("Fdate").Specific.Value
                        Dim ToDate As String = objForm.Items.Item("Tdate").Specific.Value
                        If objForm.Items.Item("Fdate").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        ElseIf objForm.Items.Item("Tdate").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            Dim DetQuery As String = ""
                            Dim DocTypeStr As String = oDBs_Head.GetValue("U_DocType", 0)
                            If DocTypeStr = "" Then
                                DocTypeStr = "388"
                            End If
                            If DocTypeStr = "381" Then
                                DetQuery = "Select ""DocNum"",""DocEntry"",""CardCode"",""CardName""," &
                                    " ""DocTotal""-""VatSum"" As ""TotalBeforeDiscount"", ""DiscSum"" As ""TotalDiscount"", ""VatSum"" As ""TotalTax"", ""DocTotal"" As ""Total"", " &
                                    " ""U_STATUS"" As ""Status"",""U_MSG"" As ""Message"" " &
                                        "from ""ORIN"" Where ""DocStatus"" = 'O' and ""DocDate"" between '" & FromDate & "' and '" & ToDate & "'"
                            ElseIf DocTypeStr = "383" Then
                                DetQuery = "Select ""DocNum"",""DocEntry"",""CardCode"",""CardName""," &
                                    " ""DocTotal""-""VatSum"" As ""TotalBeforeDiscount"", ""DiscSum"" As ""TotalDiscount"", ""VatSum"" As ""TotalTax"", ""DocTotal"" As ""Total"", " &
                                    " ""U_STATUS"" As ""Status"",""U_MSG"" As ""Message"" " &
                                        "from ""ODPI"" Where ""DocStatus"" = 'O' and ""DocDate"" between '" & FromDate & "' and '" & ToDate & "'"
                            Else
                                DetQuery = "Select ""DocNum"",""DocEntry"",""CardCode"",""CardName""," &
                                    " ""DocTotal""-""VatSum"" As ""TotalBeforeDiscount"", ""DiscSum"" As ""TotalDiscount"", ""VatSum"" As ""TotalTax"", ""DocTotal"" As ""Total"", " &
                                    " ""U_STATUS"" As ""Status"",""U_MSG"" As ""Message"" " &
                                        "from ""OINV"" Where ""DocStatus"" = 'O' and ""DocDate"" between '" & FromDate & "' and '" & ToDate & "'"
                            End If

                            Dim DocStatusStr As String = oDBs_Head.GetValue("U_STATUS", 0)
                            If DocStatusStr = "1" Then
                                DetQuery = DetQuery & " AND IFNULL(""U_STATUS"",'')='success'"
                            End If
                            If DocStatusStr = "0" Then
                                DetQuery = DetQuery & " AND IFNULL(""U_STATUS"",'')<>'success'"
                            End If

                            Dim oRsDetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsDetQuery.DoQuery(DetQuery)
                            objMatrix.Clear()
                            objMatrix.FlushToDataSource()
                            objMatrix.AutoResizeColumns()
                            For i As Integer = 1 To oRsDetQuery.RecordCount
                                objMatrix.AddRow()

                                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, i)
                                oDBs_Details.SetValue("U_Select", oDBs_Details.Offset, "N")
                                oDBs_Details.SetValue("U_DNum", oDBs_Details.Offset, oRsDetQuery.Fields.Item("DocNum").Value.ToString())
                                oDBs_Details.SetValue("U_DocEntry", oDBs_Details.Offset, oRsDetQuery.Fields.Item("DocEntry").Value.ToString())
                                oDBs_Details.SetValue("U_CCode", oDBs_Details.Offset, oRsDetQuery.Fields.Item("CardCode").Value.ToString())
                                oDBs_Details.SetValue("U_CName", oDBs_Details.Offset, oRsDetQuery.Fields.Item("CardName").Value.ToString())
                                oDBs_Details.SetValue("U_Tbdisc", oDBs_Details.Offset, oRsDetQuery.Fields.Item("TotalBeforeDiscount").Value.ToString())
                                oDBs_Details.SetValue("U_Disc", oDBs_Details.Offset, oRsDetQuery.Fields.Item("TotalDiscount").Value.ToString())
                                oDBs_Details.SetValue("U_Tax", oDBs_Details.Offset, oRsDetQuery.Fields.Item("TotalTax").Value.ToString())
                                oDBs_Details.SetValue("U_Total", oDBs_Details.Offset, oRsDetQuery.Fields.Item("Total").Value.ToString())
                                oDBs_Details.SetValue("U_STATUS", oDBs_Details.Offset, oRsDetQuery.Fields.Item("Status").Value.ToString())
                                oDBs_Details.SetValue("U_MSG", oDBs_Details.Offset, oRsDetQuery.Fields.Item("Message").Value.ToString())
                                objMatrix.SetLineData(i)

                                oRsDetQuery.MoveNext()
                            Next
                            objMain.objApplication.StatusBar.SetText("Data loaded successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            objMatrix.AutoResizeColumns()
                        End If
                    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_IP")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_IP_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific

            objMatrix.AddRow()
            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_Select", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_DNum", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_DocEntry", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_CCode", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_CName", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Tbdisc", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Disc", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Tax", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Total", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_STATUS", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_MSG", oDBs_Details.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FlickINVPosting(ByVal INVEntry As String, ByVal DocumentType As String)
        Try
            Dim Qry As String = ""
            If DocumentType = "388" Then
                Qry = "Call ""TNX_ZATCA_OINV""('" & INVEntry & "')"
            ElseIf DocumentType = "381" Then
                Qry = "Call ""TNX_ZATCA_ORIN""('" & INVEntry & "')"
            ElseIf DocumentType = "386" Then
                Qry = "Call ""TNX_ZATCA_ODPI""('" & INVEntry & "')"
            End If

            Dim rsInvoice As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            rsInvoice.DoQuery(Qry)

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

                Dim oInvoice As SAPbobsCOM.Documents = DirectCast(objMain.objCompany.GetBusinessObject(13), SAPbobsCOM.Documents)
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

    Sub FlickCreditNotePosting(ByVal INVEntry As String, ByVal DocumentType As String)
        Try
            Dim Qry As String = ""
            If DocumentType = "388" Then
                Qry = "Call ""TNX_ZATCA_OINV""('" & INVEntry & "')"
            ElseIf DocumentType = "381" Then
                Qry = "Call ""TNX_ZATCA_ORIN""('" & INVEntry & "')"
            ElseIf DocumentType = "386" Then
                Qry = "Call ""TNX_ZATCA_ODPI""('" & INVEntry & "')"
            End If

            Dim rsInvoice As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            rsInvoice.DoQuery(Qry)

            Dim dtInvoiceDt As Date = rsInvoice.Fields.Item("DocDate").Value
            Dim IssueDate As String = dtInvoiceDt.ToString("yyyy-MM-dd")
            Dim IssueTime As String = rsInvoice.Fields.Item("DocTime").Value
            Dim InvDocNum As String = rsInvoice.Fields.Item("DocNum").Value
            Dim QR As String = rsInvoice.Fields.Item("QRCodeSrc").Value
            Dim InvRef = rsInvoice.Fields.Item("NumAtCard").Value
            Dim CusCode As String = rsInvoice.Fields.Item("CusCode").Value
            Dim CustNameAr As String = rsInvoice.Fields.Item("CardFName").Value
            Dim BaseDocNum As String = rsInvoice.Fields.Item("BaseRef").Value
            Dim Comments As String = rsInvoice.Fields.Item("Comments").Value

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
            oNotesDetails.related_doc_reference = BaseDocNum
            oNotesDetails.reason_text = Comments
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

                Dim oInvoice As SAPbobsCOM.Documents = DirectCast(objMain.objCompany.GetBusinessObject(13), SAPbobsCOM.Documents)
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

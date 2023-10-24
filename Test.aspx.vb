Imports System.IO
Imports System.Net
Imports System.Web.Script.Serialization
Imports System.Windows.Forms
Imports Newtonsoft.Json

Public Class ViewReport
    Inherits System.Web.UI.Page
    Dim udf As New GUDF
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not Page.IsPostBack Then
            lbl_check.Text = "False"
            lblTotal.Text = 0.0
            lblTo.Text = 0.0
        End If
        Me.MaintainScrollPositionOnPostBack = True
    End Sub

    Protected Sub cmdAdd_Item_Click(sender As Object, e As ImageClickEventArgs)
        '=============if adding empty text=========
        If Not String.IsNullOrWhiteSpace(txt_Item.Text) Or String.IsNullOrWhiteSpace(txt_Price.Text) Then

            '=========if price is not a number (double)======
            Dim tempBool = False
            Try
                Dim p = CDbl(txt_Price.Text)
                tempBool = True
            Catch ex As Exception
                tempBool = False
            End Try
            If tempBool = False Then

            Else
                '=============add item to grid=========
                Dim col As String() = {"ID", "DATE_OF_SUPPLY", "DESCRIPTION", "DUE_AMOUNT"}
                Dim dt As New DataTable
                For Each c In col
                    dt.Columns.Add(New System.Data.DataColumn(c))
                Next
                Dim maxId = 1
                Dim item = txt_Item.Text
                item = item.Replace(Chr(34), "")
                item = item.Replace("'", "")
                item = item.Replace("%", "")
                item = item.Replace("&", "")


                If gvItems.Rows.Count <> 0 Then
                    '========if grid has rows========
                    For Each r As GridViewRow In gvItems.Rows
                        Dim lblid = r.Cells(1).Text
                        Dim lDate = r.Cells(2).Text
                        Dim lblItem = r.Cells(3).Text
                        Dim lblPrice = r.Cells(4).Text
                        If maxId < lblid Then
                            maxId = lblid
                        End If
                        dt.Rows.Add({lblid, lDate, lblItem, lblPrice})
                    Next
                    maxId = maxId + 1
                    dt.Rows.Add({maxId, txt_Date.Text, item, txt_Price.Text})
                Else
                    '======if grid has no rows=======
                    dt.Rows.Add({maxId, txt_Date.Text, item, txt_Price.Text})
                End If

                gvItems.DataSource = dt
                gvItems.DataBind()
                CalcTotal()
                txt_Item.Text = ""
                txt_Date.Text = ""
                txt_Price.Text = ""

            End If

        End If
    End Sub
    Protected Sub cmdDelete_Click(sender As Object, e As GridViewDeleteEventArgs)
        Dim index As Integer = Convert.ToInt32(e.RowIndex) + 1
        Dim col As String() = {"ID", "DATE_OF_SUPPLY", "DESCRIPTION", "DUE_AMOUNT"}
        Dim dt As New DataTable
        For Each c In col
            dt.Columns.Add(New System.Data.DataColumn(c))
        Next
        Dim rIndex As Integer = 1
        For Each r As GridViewRow In gvItems.Rows
            Dim lblid = r.Cells(1).Text
            Dim lblItem = r.Cells(2).Text
            Dim lblDate = r.Cells(3).Text
            Dim lblPrice = r.Cells(4).Text
            If lblid <> (index) Then
                dt.Rows.Add({rIndex, lblDate, lblItem, lblPrice})
                rIndex += 1
            End If

        Next
        gvItems.DataSource = dt
        gvItems.DataBind()
        CalcTotal()
    End Sub
    Protected Sub CalcTotal()
        Dim total As Double = 0.0
        Dim fees = 0.0
        Dim vat = 0.0
        Dim dis = 0.0
        For Each r As GridViewRow In gvItems.Rows
            Dim lblPrice = r.Cells(4).Text
            total += CDbl(lblPrice)
        Next
        If Not String.IsNullOrWhiteSpace(txt_Vat.Text) Then
            vat = txt_Vat.Text
        End If
        If Not String.IsNullOrWhiteSpace(txt_Fees.Text) Then
            fees = txt_Fees.Text
        End If
        If Not String.IsNullOrWhiteSpace(txt_Discount.Text) Then
            dis = txt_Discount.Text
        End If
        lblTotal.Text = total

        lblTo.Text = total + fees + vat - dis
    End Sub

    Protected Function Get_gvDt() As DataTable
        Dim col As String() = {"INVOICE_NUM", "DATE_OF_SUPPLY", "DESCRIPTION", "DUE_AMOUNT"}
        Dim dt As New DataTable
        For Each c In col
            dt.Columns.Add(New System.Data.DataColumn(c))
        Next
        For Each r As GridViewRow In gvItems.Rows
            Dim lblid = r.Cells(1).Text
            Dim ldate = r.Cells(2).Text
            Dim lblItem = r.Cells(3).Text
            Dim lblPrice = r.Cells(4).Text
            dt.Rows.Add({lblid, ldate, lblItem, lblPrice})
        Next
        Return dt
    End Function

    Protected Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
        Dim parameters As New Generic.Dictionary(Of String, String)

        parameters.Add("Client_No", txt_CustNo.Text)
        Dim tt = lblTotal.Text
        Dim vat = "0.0"
        Dim fees = "0.0"
        Dim dis = "0.0"
        If Not String.IsNullOrWhiteSpace(txt_Vat.Text) Then
            vat = txt_Vat.Text
        End If
        If Not String.IsNullOrWhiteSpace(txt_Fees.Text) Then
            fees = txt_Fees.Text
        End If
        If Not String.IsNullOrWhiteSpace(txt_Discount.Text) Then
            dis = txt_Discount.Text
        End If

        Dim total As Double = (CDbl(tt) + CDbl(fees) + CDbl(vat)) - CDbl(dis)
        'Dim total1 = Format(CStr(total), "#,###,##0.000")
        total = Format(total, "#,###,##0.000")
        '=============Report Param=================
        parameters.Add("FULL_NAME", "S.  Jalal Hasan Ebrahim")
        parameters.Add("ADDRESS", "any address")
        parameters.Add("Total_due", tt)
        parameters.Add("Trans_Amount", fees)
        parameters.Add("Discount_Amount", dis)
        parameters.Add("Vat", vat)
        parameters.Add("Total_Amount", total)
        parameters.Add("Vat_Percentage", "5")
        parameters.Add("User_Id", "882")
        parameters.Add("Ext_Ref#", "00001234000")
        'parameters.Add("Inv_Num", "00101")
        Dim js = JsonConvert.SerializeObject(parameters)
        js = js.Replace("{", "")
        js = js.Replace("}", "")
        Dim js2 = JsonConvert.SerializeObject(Get_gvDt())
        Dim JsonString = "{" + js + "," + Chr(34) + "Items" + Chr(34) + ":" + js2 + "}"
        Dim response = ShowPDF(JsonString)
    End Sub

    Function ShowPDF(jsonstring As String) As Threading.Tasks.Task

        Dim url = New Uri(HttpContext.Current.Request.Url.AbsoluteUri)
        Dim port = url.Port.ToString
        Dim host = url.Host
        Dim baseHost = "https://" + host + ":" + port + "/"

        Dim jsonObj = Calling_API("POST", baseHost + "post-invoice/", jsonstring)
        Dim jsondata = JsonConvert.DeserializeObject(jsonObj)
        If jsondata("Status") = "OK" Then
            Dim jsonObj2 As Object = Calling_API("GET", baseHost + "get-invoice/" + jsondata("ResponseValue"), "")
            Dim jsondata2 = JsonConvert.DeserializeObject(jsonObj2)
            Dim Buffer As Byte() = Convert.FromBase64String(jsondata2("ResponseBytes"))
            Response.ContentType = "application/pdf"
            Response.OutputStream.Write(Buffer, 0, Buffer.Length)
            Response.Flush()
        Else
            'udf.ShowMessage(Page, jsondata("Message"), True)
            MessageBox.Show(String.Format("Error: {0}", jsondata("Message")))
        End If

    End Function

    Function Calling_API(ByVal Method As String, ByVal url As String, ByVal contenets As String) As Object
        Dim request As WebRequest = WebRequest.Create(url)
        request.Credentials = CredentialCache.DefaultCredentials
        request.Method = Method
        If contenets <> "" Then
            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(contenets)
            request.ContentLength = byteArray.Length
            request.ContentType = "application/json"
            Dim dataStream As Stream = request.GetRequestStream()
            dataStream.Write(byteArray, 0, byteArray.Length)
            dataStream.Close()
        End If
        Dim response As WebResponse = request.GetResponse()
        Dim webStream As Stream
        Dim webResponse = ""
        webStream = response.GetResponseStream() ' Get Response
        Dim webStreamReader As New StreamReader(webStream)
        While webStreamReader.Peek >= 0
            webResponse = webStreamReader.ReadToEnd()
        End While
        Dim jsonSerializer As JavaScriptSerializer = New JavaScriptSerializer()
        Return (jsonSerializer.DeserializeObject(webResponse))
    End Function

    Private Sub txt_Fees_TextChanged(sender As Object, e As EventArgs) Handles txt_Fees.TextChanged
        CalcTotal()
    End Sub

    Private Sub txt_Vat_TextChanged(sender As Object, e As EventArgs) Handles txt_Vat.TextChanged
        CalcTotal()
    End Sub

    Private Sub txt_Discount_TextChanged(sender As Object, e As EventArgs) Handles txt_Discount.TextChanged
        CalcTotal()
    End Sub

    Protected Sub gvItems_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub
End Class
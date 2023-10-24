Imports Microsoft.Reporting.WebForms
Imports System.Net
Imports System.Web.Http
Imports Newtonsoft.Json
Imports System.Data.OleDb
Imports System.IO
Imports System.Net.Http
'ID       DATE          BY           CHANGE
'=======================================
'001     25/10/2022    Marwa         Transaction number instead of invoice number and save file name as name given in parameter.
Namespace Controllers
    Public Class InvoiceController
        Inherits ApiController
        Dim UDF As New GUDF

        'Get :    /EBULTS_APIS/get-invoice/00001
        <HttpGet>
        <Route("get-invoice/{inv_num}")>
        Public Function GetInvoice(inv_num As String) As String

            Dim bytes As Byte()
            Dim parameters As New Dictionary(Of String, String)
            Dim filePath = HttpContext.Current.Server.MapPath(UDF.GetAppPath() + "/Generated/Invoices/" + inv_num + ".pdf")

            If Not System.IO.File.Exists(filePath) Then
                Try
                    Dim sql = "SELECT * FROM INVOICING_INVOICES WHERE INVOICE_NUM='" + inv_num + "' "
                    Dim EBDB = UDF.EBDB_CS()
                    Dim ICBS = UDF.ICBS_CS()
                    Dim dr = UDF.getDataRow(UDF.EBDB_CS(), sql)
                    sql = "SELECT * FROM INVOICING_INVOICE_ITEMS WHERE INVOICE_NUM='" + inv_num + "' "
                    Dim dt = UDF.GetDataTable(UDF.EBDB_CS(), sql)
                    Dim CustTin = ""
                    Dim Custname = ""
                    Dim CustAddr = ""
                    If dr Is Nothing Then
                        Dim ApiResponse As New Dictionary(Of String, Object)
                        ApiResponse.Add("Status", "ERROR")
                        ApiResponse.Add("Message", "Invalid Invoice Number")
                        ApiResponse.Add("ResponseBytes", "")
                        Dim js = JsonConvert.SerializeObject(ApiResponse)
                        Return js
                    Else
                        Dim custDr = GetCustomerInfo(dr("CUST_NO"), CustTin, Custname, CustAddr)
                        parameters.Add("CustName", Custname)
                        parameters.Add("CustAddress", CustAddr)
                        parameters.Add("Client_No", dr("CUST_NO").ToString)
                        parameters.Add("Client_TIN", CustTin)
                        parameters.Add("Total_due", dr("Total_due").ToString)
                        parameters.Add("Trans_Amount", dr("Trans_Amount").ToString)
                        parameters.Add("Discount_Amount", dr("Disc_Amount").ToString)
                        parameters.Add("Vat", dr("Vat_Amount").ToString)
                        parameters.Add("Total_Amount", dr("Total_Amount").ToString)
                        '=============== CHANGE LOG : 001 : Print Ext Ref as invoice num============= 
                        If String.IsNullOrEmpty(dr("Ext_Ref#").ToString) Then
                            parameters.Add("Invoice_Num", inv_num)
                        Else
                            parameters.Add("Invoice_Num", dr("Ext_Ref#").ToString)
                        End If
                        '=============== END CHANGE============= 
                        parameters.Add("Date_Issued", Date.Now())
                            parameters.Add("Vat_Percentage", dr("Vat_Percentage"))

                            bytes = Generate(parameters, dt)
                            Dim responseBytes = Convert.ToBase64String(bytes)

                            Dim ApiResponse As New Dictionary(Of String, Object)
                            ApiResponse.Add("Status", "OK")
                            ApiResponse.Add("Message", "Successful")
                            ApiResponse.Add("ResponseBytes", responseBytes)
                            Dim js = JsonConvert.SerializeObject(ApiResponse)
                            Return js
                        End If

                Catch ex As Exception
                    Dim ApiResponse As New Dictionary(Of String, Object)
                    ApiResponse.Add("Status", "ERROR")
                    ApiResponse.Add("Message", ex.Message)
                    ApiResponse.Add("ResponseBytes", "")
                    Dim js = JsonConvert.SerializeObject(ApiResponse)
                    Return js
                End Try
            Else
                bytes = File.ReadAllBytes(filePath)
                Dim ApiResponse As New Dictionary(Of String, Object)
                ApiResponse.Add("Status", "OK")
                ApiResponse.Add("Message", "Successful")
                ApiResponse.Add("ResponseBytes", bytes)
                Dim js = JsonConvert.SerializeObject(ApiResponse)
                Return js
            End If
        End Function
        'POST: /api/invoice
        <HttpPost>
        <Route("post-invoice/")>
        Public Function PostInvoice(<FromBody()> ByVal JsonString) As String
            '===========Initiate variables================

            Dim parameters As New Dictionary(Of String, String)
            Dim col As String() = {"INVOICE_NUM", "DATE_OF_SUPPLY", "DESCRIPTION", "DUE_AMOUNT"}
            Dim dt As New System.Data.DataTable
            For Each c In col
                dt.Columns.Add(New System.Data.DataColumn(c))
            Next

            Dim user_Id As String = ""
            Dim Ext_Ref_num As String = ""

            Dim msg As String = ""
            '============================================
            FetchData(JsonString, user_Id, Ext_Ref_num, parameters, dt) '=====fill parameters and dt=======
            Dim Inv_Num = parameters.Item("Invoice_Num")
            Try
                Dim isSaved = SaveToDB(parameters, Ext_Ref_num, user_Id, dt, msg)

                If isSaved = True Then

                    '=============== CHANGE LOG : 001 : Print Ext Ref as invoice num============= 
                    If Not String.IsNullOrEmpty(Ext_Ref_num) Then
                        parameters("Invoice_Num") = Ext_Ref_num
                    End If
                    '=============== END CHANGE 001 ============= 
                    Dim renderedBytes = Generate(parameters, dt)
                    Dim TargetPath As String = HttpContext.Current.Server.MapPath("../Generated/Invoices/" + Inv_Num + ".pdf")
                    Dim oFileStream As System.IO.FileStream
                    oFileStream = New System.IO.FileStream(TargetPath, System.IO.FileMode.Create)
                    oFileStream.Write(renderedBytes, 0, renderedBytes.Length)
                    oFileStream.Close()
                    Dim ApiResponse As New Dictionary(Of String, Object)
                    ApiResponse.Add("Status", "OK")
                    ApiResponse.Add("Message", "Successful")
                    ApiResponse.Add("ResponseValue", Inv_Num)
                    Dim js = JsonConvert.SerializeObject(ApiResponse)
                    Return js
                Else
                    Dim ApiResponse As New Dictionary(Of String, Object)
                    ApiResponse.Add("Status", "ERROR")
                    ApiResponse.Add("Message", msg)
                    ApiResponse.Add("ResponseValue", "")
                    Dim js = JsonConvert.SerializeObject(ApiResponse)
                    Return js
                End If

            Catch ex As Exception
                Dim ApiResponse As New Dictionary(Of String, String)
                ApiResponse.Add("Status", "ERROR")
                ApiResponse.Add("Message", ex.Message)
                ApiResponse.Add("ResponseValue", "")
                Dim js = JsonConvert.SerializeObject(ApiResponse)
                Return js
            End Try

        End Function
        Private Function SaveToDB(p As Dictionary(Of String, String), Ext_Ref_Num As String, user_Id As String, dt As DataTable, ByRef lcMessage As String) As Boolean
            Dim tempBool = False
            Try

                Dim sql = "SELECT INVOICE_NUM FROM INVOICING_INVOICES WHERE INVOICE_NUM = '" + p("Invoice_Num") + "'"
                Dim dr = UDF.getDataRow(UDF.EBDB_CS, sql)
                If dr Is Nothing Then
                    '==================Insert in EBDB if invoice number is not dup=================
                    '=============== CHANGE LOG : 001 : Insert Ext Ref num============= 
                    sql = "INSERT INTO INVOICING_INVOICES VALUES('" + p("Invoice_Num") + "','" + p("Client_No") + "','" + p("Total_Amount") + "','"
                    sql = sql + p("Date_Issued") + "','" + p("Total_due") + "','" + p("Trans_Amount") + "','" + p("Discount_Amount") + "','"
                    sql = sql + p("Vat") + "','" + user_Id + "','" + p("Vat_Percentage") + "','" + Ext_Ref_Num + "') "
                    Dim runBool = UDF.RunSql(UDF.EBDB_CS, sql)
                    If runBool = True Then

                        For Each row In dt.Rows()
                            sql = "INSERT INTO INVOICING_INVOICE_ITEMS(INVOICE_NUM,DATE_OF_SUPPLY,DESCRIPTION,DUE_AMOUNT) VALUES('"
                            sql = sql + row("INVOICE_NUM") + "','" + row("DATE_OF_SUPPLY") + "','" + row("DESCRIPTION") + "','" + row("DUE_AMOUNT") + "')"
                            Dim runbool2 = UDF.RunSql(UDF.EBDB_CS, sql)
                        Next
                        lcMessage = "Invoice Added Successfully."
                        tempBool = True
                    End If
                Else
                    lcMessage = "Invoice Number Exists Already."
                    tempBool = False
                End If

            Catch ex As Exception
                lcMessage = ex.Message
                tempBool = False
            End Try
            Return tempBool
        End Function

        Private Sub FetchData(JsonObj As Object, ByRef User_Id As String, ByRef Ext_Ref_Num As String, ByRef parameters As Dictionary(Of String, String), ByRef dt As DataTable)
            '===================Fetching Customer details========================

            Dim INV_NUM = ""
            Ext_Ref_Num = ""
            User_Id = JsonObj("User_Id").ToString
            'Dim dr = GetCustomerInfo(JsonObj("Client_No"), CustTin, CustAddr, CustName)

            '=============== CHANGE LOG : 001 ============= 
            If JsonObj.ContainsKey("Ext_Ref#") Then
                Ext_Ref_Num = JsonObj("Ext_Ref#")
            End If
            'END CHANGE
            If JsonObj.ContainsKey("Inv_Num") Then
                INV_NUM = JsonObj("Inv_Num")
            Else
                INV_NUM = UDF.GetNextID_New(UDF.EBDB_CS(), "INVOICING_INVOICES", "INVOICE_NUM", 5, 1)
            End If

            '================Fetching customer details end=======================

            parameters.Add("CustName", JsonObj("FULL_NAME"))
            parameters.Add("CustAddress", JsonObj("ADDRESS"))
            parameters.Add("Client_No", JsonObj("Client_No"))
            parameters.Add("Client_TIN", JsonObj("TIN"))
            parameters.Add("Total_due", JsonObj("Total_due"))
            parameters.Add("Trans_Amount", JsonObj("Trans_Amount"))
            parameters.Add("Discount_Amount", JsonObj("Discount_Amount"))
            parameters.Add("Vat", JsonObj("Vat"))
            parameters.Add("Total_Amount", JsonObj("Total_Amount"))
            parameters.Add("Date_Issued", Date.Now())
            parameters.Add("Vat_Percentage", JsonObj("Vat_Percentage"))
            parameters.Add("Invoice_Num", INV_NUM)

            Dim items As Object = JsonObj("Items")
            Dim INum = INV_NUM
            For Each x In items
                Dim dateSupply = x("DATE_OF_SUPPLY")
                Dim Description = x("DESCRIPTION")
                Dim Due_Amount = x("DUE_AMOUNT")
                dt.Rows.Add(INum, dateSupply, Description, Due_Amount)
            Next
        End Sub
        Private Function GetCustomerInfo(Cust_ID As String, ByRef CustTin As String, ByRef CustAddr As String, ByRef CustName As String)
            'Adjustment 06/09/2022 by marwa to check if cust is moral or physical as requested by Aqeel.
            Dim custType = Chech_Moral_Physical(Cust_ID)
            Dim dr As DataRow
            If custType("CUSM_TYPE") = "1" Then
                dr = Get_Physical(Cust_ID)
            Else custType("CUSM_TYPE") = "2"
                dr = Get_Moral(Cust_ID)
            End If

            If dr IsNot Nothing Then
                CustTin = dr("TIN").ToString()
                CustName = dr("FULL_NAME").ToString()

                If Not String.IsNullOrEmpty(dr("BLOCK").ToString()) Then
                    CustAddr = CustAddr + " Blk " + dr("BLOCK").ToString()
                End If
                If Not String.IsNullOrEmpty(dr("ROAD").ToString()) Then
                    CustAddr = CustAddr + " Rd " + dr("ROAD").ToString()
                End If
                If Not String.IsNullOrEmpty(dr("BUILDING").ToString()) Then
                    CustAddr = CustAddr + " Bld " + dr("BUILDING").ToString()
                End If
                If Not String.IsNullOrEmpty(dr("FLAT").ToString()) Then
                    CustAddr = CustAddr + " Flt " + dr("FLAT").ToString()
                End If
            End If
            Return dr
        End Function
        '===============================Supporting GUDF Fuctions===========================

        Private Function Chech_Moral_Physical(ByVal Cust_ID As String) As Data.DataRow
            Dim lcSql As String = ""
            lcSql = lcSql + " Select  CUSM_TYPE FROM BBSD_CUST_MEMBERS WHERE CUST_ID='" + Cust_ID + "' "
            Dim dr = UDF.getDataRow(UDF.ICBS_CS(), lcSql)
            Return (dr)
        End Function

        Private Function Get_Physical(ByVal Cust_ID As String) As Data.DataRow
            Dim lcSql As String = ""
            lcSql = lcSql + " Select  "
            lcSql = lcSql + "       T1.PHPR_ID AS ID, T3.CUST_ID AS Customer_Number, "
            lcSql = lcSql + "       T1.PHPR_FULL_NAME FULL_NAME, T1.PHPR_TAX_ID TIN, "
            lcSql = lcSql + "       T4.LCTY_CODE AS BLOCK, T4.PADR_B_LINE_2 AS BUILDING, T4.PADR_B_LINE_3 AS FLAT, T4.PADR_B_LINE_4 AS ROAD"
            lcSql = lcSql + " from bbsd_physical_persons T1 "
            lcSql = lcSql + " left join bbsd_cust_members T2 on T1.PHPR_ID=T2.CUSM_ID "
            lcSql = lcSql + " Left Join BBSD_CUSTOMERS T3 on T2.CUST_ID=T3.CUST_ID "
            lcSql = lcSql + " Left Join BBSD_PHPR_ADDRESSES T4 on T1.PHPR_ID=T4.PHPR_ID"
            lcSql = lcSql + " WHERE T3.CUST_ID='" + Cust_ID + "'"
            Dim dr = UDF.getDataRow(UDF.ICBS_CS(), lcSql)
            Return (dr)
        End Function

        Private Function Get_Moral(Cust_ID As String) As Data.DataRow
            Dim lcSql As String = ""

            lcSql = lcSql + " Select distinct "
            lcSql = lcSql + " T1.MRPR_ID AS ID, T1.MRPR_ID As Customer_Number, "
            lcSql = lcSql + " T1.MRPR_B_NAME FULL_NAME, T1.MRPR_TAX_ID_NBR TIN, "
            lcSql = lcSql + " T4.LCTY_CODE AS BLOCK, "
            lcSql = lcSql + " T4.MADR_B_LINE_2 As BUILDING, "
            lcSql = lcSql + " T4.MADR_B_LINE_3 AS FLAT, "
            lcSql = lcSql + " T4.MADR_B_LINE_4 AS ROAD "
            lcSql = lcSql + " From bbsd_moral_persons T1 "
            lcSql = lcSql + " Left Join bbsd_cust_members T2 on T1.MRPR_ID=T2.CUSM_ID "
            lcSql = lcSql + " Left Join BBSD_CUSTOMERS T3 on T2.CUST_ID=T3.CUST_ID "
            lcSql = lcSql + " Left Join BBSD_MRPR_ADDRESSES T4 on T1.MRPR_ID=T4.MRPR_ID "
            lcSql = lcSql + " WHERE T1.MRPR_ID ='" + Cust_ID + "'"

            Dim dr = UDF.getDataRow(UDF.ICBS_CS(), lcSql)
            Return dr
        End Function






        Private Function Generate(params As Dictionary(Of String, String), Data As DataTable) As Byte()
            Dim encoding As String = Nothing
            Dim streams As String() = Nothing
            Dim mimeType As String = "application/PDF"
            Dim extension As String = "pdf"
            Dim warnings As Warning() = Nothing

            Dim deviceInfo As String = GetDeviceInfo(False)
            Dim report As New LocalReport()
            report.ReportPath = HttpContext.Current.Server.MapPath("../Generated/Template/Invoice.rdlc")

            For Each param In params
                Try
                    Dim reportParam As New ReportParameter(param.Key, param.Value)
                    report.SetParameters(reportParam)
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

            Next
            report.DataSources.Add(New ReportDataSource("dsInvoice", Data))
            Return report.Render("PDF", deviceInfo, mimeType, encoding, extension, streams, warnings)
        End Function
        Private Function GetDeviceInfo(Optional isLandScape As Boolean = False) As String

            Dim sbDeviceInfo As New StringBuilder()
            sbDeviceInfo.Append("<DeviceInfo> ")
            sbDeviceInfo.AppendFormat("<OutputFormat>{0}</OutputFormat>", "PDF")
            If (isLandScape) Then
                sbDeviceInfo.Append("<Orientation>Landscape</Orientation>")
                sbDeviceInfo.Append("<PageWidth>11in</PageWidth>")
                sbDeviceInfo.Append("<PageHeight>8.5in</PageHeight></DeviceInfo>")
            Else
                sbDeviceInfo.Append("<PageWidth>8.5in</PageWidth>")
                sbDeviceInfo.Append("<PageHeight>11in</PageHeight></DeviceInfo>")
            End If
            Return sbDeviceInfo.ToString()
        End Function

        'Private Function Get_Moral(Cust_ID As String) As Data.DataRow
        '    Dim lcSql As String = ""
        '    lcSql = lcSql + " Select distinct "
        '    lcSql = lcSql + " T1.MRPR_ID AS ID, T1.MRPR_ID As Customer_Number, "
        '    lcSql = lcSql + " T1.MRPR_B_NAME FULL_NAME, T1.MRPR_TAX_ID_NBR TIN, "
        '    lcSql = lcSql + " T4.LCTY_CODE AS BLOCK, "
        '    lcSql = lcSql + " T4.MADR_B_LINE_2 As BUILDING, "
        '    lcSql = lcSql + " T4.MADR_B_LINE_3 AS FLAT, "
        '    lcSql = lcSql + " T4.MADR_B_LINE_4 AS ROAD "
        '    lcSql = lcSql + " From bbsd_moral_persons T1 "
        '    lcSql = lcSql + " Left Join bbsd_cust_members T2 on T1.MRPR_ID=T2.CUSM_ID "
        '    lcSql = lcSql + " Left Join BBSD_CUSTOMERS T3 on T2.CUST_ID=T3.CUST_ID "
        '    lcSql = lcSql + " Left Join BBSD_MRPR_ADDRESSES T4 on T1.MRPR_ID=T4.MRPR_ID "
        '    lcSql = lcSql + " WHERE T1.MRPR_ID ='" + Cust_ID + "'"

        '    Dim dr = UDF.getDataRow(UDF.ICBS_CS(), lcSql)
        '    Return (dr)
        'End Function
        'Private Function Get_Physical(ByVal Cust_ID As String) As Data.DataRow
        '    Dim lcSql As String = ""
        '    lcSql = lcSql + " Select  "
        '    lcSql = lcSql + "       T1.PHPR_ID AS ID, T3.CUST_ID AS Customer_Number, "
        '    lcSql = lcSql + "       T1.PHPR_FULL_NAME FULL_NAME, T1.PHPR_TAX_ID TIN, "
        '    lcSql = lcSql + "       T4.LCTY_CODE AS BLOCK, T4.PADR_B_LINE_2 AS BUILDING, T4.PADR_B_LINE_3 AS FLAT, T4.PADR_B_LINE_4 AS ROAD"
        '    lcSql = lcSql + " from bbsd_physical_persons T1 "
        '    lcSql = lcSql + " left join bbsd_cust_members T2 on T1.PHPR_ID=T2.CUSM_ID "
        '    lcSql = lcSql + " Left Join BBSD_CUSTOMERS T3 on T2.CUST_ID=T3.CUST_ID "
        '    lcSql = lcSql + " Left Join BBSD_PHPR_ADDRESSES T4 on T1.PHPR_ID=T4.PHPR_ID"
        '    lcSql = lcSql + " WHERE T3.CUST_ID='" + Cust_ID + "'"
        '    Dim dr = UDF.getDataRow(UDF.ICBS_CS(), lcSql)
        '    Return (dr)
        'End Function
        'Private Function Chech_Moral_Physical(ByVal Cust_ID As String) As Data.DataRow
        '    Dim lcSql As String = ""
        '    lcSql = lcSql + " Select  CUSM_TYPE FROM BBSD_CUST_MEMBERS WHERE CUST_ID='" + Cust_ID + "' "
        '    Dim dr = UDF.getDataRow(UDF.ICBS_CS(), lcSql)
        '    Return (dr)
        'End Function
    End Class



End Namespace
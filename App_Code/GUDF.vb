Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Net.Mail
Imports System
Imports System.IO
Imports System.Text
Imports System.Globalization
Imports System.Web.UI.WebControls
Imports System.Web
Imports System.Web.UI.HtmlControls
Imports System.Web.UI
Imports System.Data.OleDb
Imports System.Xml

'Imports System.Data
'Imports System.Data.OleDb
'Imports Microsoft.VisualBasic
'Imports System.Globalization
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
'Imports System.IOm 
Imports System.Web.HttpRequest

Public Class GUDF

    Inherits System.Web.UI.Page
    'Dim aaaa As System.Web.HttpContex = System.Web.HttpContext.Current
    'Dim context1 As System.Web.HttpContext = System.Web.HttpContext.Current
    'Dim Server_IP As String = context1.Request.ServerVariables("HTTP_X_FORWARDED_FOR")
    ' Dim Server_Name As String = context1.Request.ServerVariables("SERVER_NAME")

    Dim dt As Data.DataTable
    Dim cmd As OleDbCommand
    Dim odr As OleDbDataReader
    Dim oda As OleDbDataAdapter

    Public Lookups As New Collection

    Public Function Get_Server_Url() As String
        Dim context As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim xx As String = Replace(context.Request.Url.AbsoluteUri, context.Request.RawUrl, "")
        If InStr(LCase(context.Request.Url.AbsoluteUri), "eskanbank.com") > 0 Then
            Dim laServerUrl As String() = Split(context.Request.Url.AbsoluteUri, "/")
            Dim lcServerUrl As String = laServerUrl(0) + "//" + Split(laServerUrl(2), ":")(0) + GetAppPath()
            xx = lcServerUrl
            If InStr(LCase(xx), "https") = 0 Then
                xx = Replace(xx, "http", "https")
            End If
        Else
            Dim laServerUrl As String() = Split(context.Request.Url.AbsoluteUri, "/")
            xx = laServerUrl(0) + "//" + laServerUrl(2) + GetAppPath()
        End If
        Return (xx)
    End Function


    Function EFTS_CS(Optional lcDatabase As String = "SBMDB", Optional lcUserID As String = "inap", Optional lcUserPW As String = "inap") As String
        EFTS_CS = "Provider=sqloledb;Data Source=s0312;Initial Catalog=EFTS_TEST;User Id=efts;Password=Nest1237;"           ' test
        EFTS_CS = "Provider=sqloledb;Data Source=s0313;Initial Catalog=EFTS_PROD;User Id=efts;Password=Nest1235;"           ' Producton
    End Function


    Function Create_Dynamic_Detaild_View(ByVal aColumns As String())
        Dim DT As New System.Data.DataTable
        For Each c In aColumns
            DT.Columns.Add(New System.Data.DataColumn(c))
        Next
        Return DT
    End Function

    Function Add_Dynamic_Row(ByRef dt As Data.DataTable, ByVal aValues As String()) As Boolean
        Dim rw As System.Data.DataRow
        rw = dt.NewRow()
        Dim i As Integer = 0
        For Each v In aValues
            rw(dt.Columns(i).ColumnName) = v
            i = i + 1
        Next
        dt.Rows.Add(rw)
        Return (True)
    End Function

    Function Get_Moral(ByVal lcCR As String) As Data.DataRow
        Dim lcSql As String = ""
        lcSql = lcSql + " Select  "
        lcSql = lcSql + "       T1.MRPR_ID AS ID, "
        lcSql = lcSql + "       T1.MRPR_COMM_REG_NBR AS CR, "
        lcSql = lcSql + "       T3.CUST_ID AS Customer_Number, "
        lcSql = lcSql + "       T1.MRPR_B_NAME FULL_NAME,"
        lcSql = lcSql + "       T4.MADR_PHONE_1 as Land_Line,"
        lcSql = lcSql + "       T4.MADR_PHONE_2 as Mobile,"
        lcSql = lcSql + "       T4.MADR_B_LINE_2 as Building, "
        lcSql = lcSql + "       T4.MADR_B_LINE_3 as Flat, "
        lcSql = lcSql + "       T4.MADR_B_LINE_4 as Road, "
        lcSql = lcSql + "       T4.LCTY_CODE as Block, "
        lcSql = lcSql + "       T4.MADR_E_MAIL_1 as Email "
        lcSql = lcSql + " from BBSD_MORAL_PERSONS T1 "
        lcSql = lcSql + " left join bbsd_cust_members T2 on T1.MRPR_ID||'2'=T2.CUSM_ID||CUSM_TYPE "             ' companies
        lcSql = lcSql + " Left Join BBSD_CUSTOMERS T3 on T2.CUST_ID=T3.CUST_ID "
        lcSql = lcSql + " Left Join BBSD_MRPR_ADDRESSES T4 on T1.MRPR_ID=T4.MRPR_ID"
        lcSql = lcSql + " WHERE MRPR_COMM_REG_NBR='" + lcCR + "'"
        Dim dr As DataRow = getDataRow(ICBS_CS, lcSql)
        Return (dr)
    End Function

    Function Get_Customer_Using_CPR(ByVal lcCPR As String) As Data.DataRow
        Dim lcSql As String = ""
        lcSql = lcSql + " Select  "
        lcSql = lcSql + "       T1.PHPR_ID AS ID, "
        lcSql = lcSql + "       T1.PHPR_NATIONAL_NBR AS CPR, "
        lcSql = lcSql + "       T3.CUST_ID AS Customer_Number, "
        lcSql = lcSql + "       T1.PHPR_FULL_NAME FULL_NAME,"
        lcSql = lcSql + "       T4.PADR_PHONE_1 as Land_Line,"
        lcSql = lcSql + "       T4.PADR_PHONE_2 as Mobile"
        lcSql = lcSql + " from bbsd_physical_persons T1 "
        lcSql = lcSql + " left join bbsd_cust_members T2 on T1.PHPR_ID=T2.CUSM_ID "
        lcSql = lcSql + " Left Join BBSD_CUSTOMERS T3 on T2.CUST_ID=T3.CUST_ID "
        lcSql = lcSql + " Left Join BBSD_PHPR_ADDRESSES T4 on T1.PHPR_ID=T4.PHPR_ID"
        lcSql = lcSql + " WHERE PHPR_NATIONAL_NBR='" + lcCPR + "'"
        Dim dr As Data.DataRow = getDataRow(ICBS_CS, lcSql)
        Return (dr)
    End Function


    Function isValidCPR(ByVal lcCPR As String) As Boolean
        isValidCPR = False
        Dim Sum As Double = 0
        Dim ValidateCPR As Boolean = False
        If Len(lcCPR) <> 9 Then Exit Function
        For i = 1 To Len(lcCPR) - 1
            Dim c As String = Mid(lcCPR, i, 1)
            Sum = Sum + Val(c) * (i + 1)
        Next
        Dim lc9th As String = Right("000" + CStr(Sum Mod 11), 1)
        If Right(lcCPR, 1) <> lc9th Then Exit Function
        isValidCPR = True
    End Function

    Function getValidCPR(ByVal lcCPR As String) As String
        getValidCPR = ""
        Dim Sum As Double = 0
        Dim ValidateCPR As Boolean = False
        Dim lcCPR1 As String = Left(lcCPR, 8)
        For i = 1 To Len(lcCPR1)
            Dim c As String = Mid(lcCPR, i, 1)
            Sum = Sum + Val(c) * (i + 1)
        Next
        Dim lc9th As String = Right("000" + CStr(Sum Mod 11), 1)
        getValidCPR = lcCPR1 + lc9th
    End Function

    Function Format_Old_System_Date(ByVal lcDate As String) As String
        If lcDate <> "" Then
            If Val(lcDate) <> 0 Then
                Return (Mid(lcDate, 1, 4) + "/" + Mid(lcDate, 5, 2) + "/" + Mid(lcDate, 7, 2))
            End If
        End If
        Return ("")
    End Function

    Public Function Get_Server_Name() As String
        'Try
        '    Dim context As System.Web.HttpContext = System.Web.HttpContext.Current
        '    Dim sIPAddress As String = context.Request.ServerVariables("HTTP_X_FORWARDED_FOR")
        '    If String.IsNullOrEmpty(sIPAddress) Then
        '        Return context.Request.ServerVariables("SERVER_NAME")
        '    Else
        '        Dim ipArray As String() = sIPAddress.Split(New [Char]() {","c})
        '        Return ipArray(0)
        '    End If
        'Catch ex As Exception
        '    'If String.IsNullOrEmpty(Server_IP) Then
        '    '    Return Server_Name
        '    'Else
        '    '    Dim ipArray As String() = Server_IP.Split(New [Char]() {","c})
        '    '    Return ipArray(0)
        '    'End If
        'End Try
        Return "localhost"
        'sIPAddress
    End Function

    Sub CreateTabsMenu(ByRef mnu As System.Web.UI.WebControls.Menu, ByVal lcItems As String, Optional ByVal lnSelectedIndex As Integer = 0)
        Dim i As Integer
        mnu.Items.Clear()
        Dim laItems = Split(lcItems, "^")
        For i = LBound(laItems) To UBound(laItems)
            Dim x As New System.Web.UI.WebControls.MenuItem
            x.Text = Split(laItems(i) + ";;;;", ";", 2)(0)
            'x.Target = Split(laItems(i) + ";;;;", ";", 2)(1)
            'x.Value = CStr(i + 1)
            x.Value = Split(laItems(i) + ";;;;", ";", 2)(1)
            If x.Text <> "" Then
                If (lnSelectedIndex = -1 And i = 0) Or lnSelectedIndex = i Then
                    x.Selected = True
                Else
                    x.Selected = False
                End If
                mnu.Items.Add(x)
            End If
        Next
    End Sub

    Function Get_Business_Date() As String
        Dim lcBusiness_Date As String = getDataColumn(ICBS_CS, " SELECT MAX(to_char(BRCH_CURRENT_DATE,'yyyy/mm/dd')) FROM BBSC_BRANCHES")
        Get_Business_Date = lcBusiness_Date
    End Function

    Function GetAreaName(lcBlock As String) As String
        GetAreaName = ""
        Dim lcSql As String = "Select LCTY_B_DESC from BBSC_LOCALITIES where LCTY_CODE='" + lcBlock + "' "
        GetAreaName = Trim(GetSqlColumn(ICBS_CS, lcSql))
    End Function

    Function FixNullDouble(v) As Double
        FixNullDouble = 0
        If Not IsDBNull(v) Then
            FixNullDouble = Val(CStr(Replace(v, ",", "")))
        End If
    End Function

    Public Shared Function GetIPAddress() As String
        Dim context As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim sIPAddress As String = context.Request.ServerVariables("HTTP_X_FORWARDED_FOR")
        If String.IsNullOrEmpty(sIPAddress) Then
            Return context.Request.ServerVariables("REMOTE_ADDR")
        Else
            Dim ipArray As String() = sIPAddress.Split(New [Char]() {","c})
            Return ipArray(0)
        End If
    End Function

    Function GetSessionVariable(lcVAR As String) As String
        'Dim lcSID As String = GetCurrentUser()
        Dim lcSID As String = GetIPAddress()
        'Dim lcSID As String = Session.SessionID
        'If lcSID = "" Then lcSID = Request.Url.Authority
        GetSessionVariable = getDataColumn(EBDB_CS, "Select SVAL From EB_SessionsVariables where nvl(SID,'')='" + lcSID + "' AND SVAR='" + lcVAR + "'", "SVAL")
    End Function
    Function getCurrentCompany() As String
        getCurrentCompany = GetSessionVariable("OID")
        getCurrentCompany = "100"
    End Function

    Function UnformatAmount(ByVal lcAmount As String) As Double
        UnformatAmount = Val(Replace(lcAmount, ",", ""))
    End Function


    Function PadL(ByVal S As String, ByVal l As Integer, Optional ByVal c As String = " ") As String
        PadL = Right(New String(c, l) + CStr(S), l)
    End Function

    Function PadR(ByVal S As String, ByVal l As Integer, Optional ByVal c As String = " ") As String
        PadR = Left(CStr(S) + New String(c, l), l)
    End Function

    Function Get_Login_Page(ByRef p As Web.UI.Page) As String
        Dim laServerUrl As String() = Split(p.Request.Url.ToString, "/")
        Dim lcServerUrl As String = laServerUrl(0) + "//" + laServerUrl(2) + GetAppPath()
        Get_Login_Page = lcServerUrl + "/LoginPage.aspx"
    End Function

    Function FixDate(lcDate As String, Optional lcFormat As String = "yyyy-MMM-dd") As String
        lcDate = Replace(lcDate, "/", "-")
        If lcDate = "" Then
            FixDate = ""
            Exit Function
        End If
        Dim lcYear As String = Split(lcDate, "-")(2)
        Dim lcMonth As String = Split(lcDate, "-")(1)
        Dim lcDay As String = Split(lcDate, "-")(0)
        If lcDay > 31 Then
            lcYear = Split(lcDate, "-")(0)
            lcMonth = Split(lcDate, "-")(1)
            lcDay = Split(lcDate, "-")(2)
        End If
        FixDate = Format(CDate(lcYear + "/" + lcMonth + "/" + lcDay), lcFormat)
    End Function

    Function GetFilesList(ByVal lcFolder As String, ByVal lcPrefFix As String) As Data.DataTable
        Dim DT As New Data.DataTable()
        Dim Row1 As Data.DataRow
        DT.Columns.Add(New Data.DataColumn("NAME", System.Type.GetType("System.String")))
        DT.Columns.Add(New Data.DataColumn("Date", System.Type.GetType("System.String")))
        DT.Columns.Add(New Data.DataColumn("Size", System.Type.GetType("System.String")))
        Dim lcPath As String = Server.MapPath(GetFilesFolder(lcFolder))
        Dim lcFile As String = Dir(lcPath + "\" + lcPrefFix + "*")
        Do While lcFile <> ""
            Dim lcFileName As String = lcPath + "\" + lcFile
            Dim lnLen As Integer = FileLen(lcFileName)
            Row1 = DT.NewRow()
            Row1("NAME") = lcFile
            Row1("DATE") = FileDateTime(lcFileName)
            Row1("SIZE") = CStr(lnLen / 1000) + " KB"
            DT.Rows.Add(Row1)
            lcFile = Dir()
        Loop
        DT.DefaultView.Sort = "NAME DESC"
        GetFilesList = DT
    End Function

    Function Array2DataTable(ByVal laArray As String(), Optional ByVal lcFields As String = "") As Data.DataTable
        Dim y As Integer
        Dim DT As New Data.DataTable()
        Dim Row1 As Data.DataRow
        Dim I As Integer
        Dim lcFldName As String
        Dim lcKeys As String = ""
        Dim llFirstLine As Boolean = True
        Dim laColumns As String() = Split(lcFields, ",")
        Dim lnColumns As Integer = 0
        For y = LBound(laArray) To UBound(laArray)
            If laArray(y) <> "" Then
                Dim laLine As String() = Split(Replace(laArray(y), "^", vbTab), vbTab)
                If llFirstLine Then
                    lnColumns = laLine.Length
                    For I = 0 To lnColumns - 1
                        lcFldName = "F" + CStr(I)
                        If lcFields <> "" Then
                            If I < laColumns.Length Then
                                lcFldName = laColumns(I)
                            End If
                        End If
                        DT.Columns.Add(New Data.DataColumn(lcFldName, System.Type.GetType("System.String")))
                    Next
                    llFirstLine = False
                End If
                Row1 = DT.NewRow()
                For I = 0 To lnColumns - 1
                    lcFldName = "F" + CStr(I)
                    If I < laColumns.Length Then
                        lcFldName = laColumns(I)
                    End If
                    If lcFldName <> "" Then
                        Row1(lcFldName) = laLine(I)
                    End If
                Next
                DT.Rows.Add(Row1)
            End If
        Next
        Return DT
    End Function

    Sub Get_Sql_Collection(ByVal lcSql As String, ByRef cl As Collection, ByVal laKeys As String(), ByVal lcPreFix As String, Optional ByVal IgnoreDuplicates As Boolean = False)
        Dim Dt1 As Data.DataTable
        Dt1 = GetDataTable(EBDB_CS, lcSql)
        Dim rs As DataRow
        Dim I As Integer
        cl.Clear()
        Dim lcKeys As String = ""
        For Each rs In Dt1.Rows
            Dim lcLine As String = ""
            For I = 0 To Dt1.Columns.Count - 1
                lcLine = lcLine + IIf(lcLine = "", "", vbTab) + FixNull(rs(I))
            Next
            Dim lcKey As String = ""
            For I = 0 To UBound(laKeys)
                Dim lcKeyItem As String = ""
                lcKeyItem = CStr(rs(laKeys(I)))
                lcKey = lcKey + IIf(I = 0, "", "-") + lcKeyItem
            Next
            If Not cl.Contains(lcPreFix + "-" + lcKey) Then
                cl.Add(lcLine, lcPreFix + "-" + lcKey)
            Else
                Dim lcOldValues As String = cl(lcPreFix + "-" + lcKey)
                cl.Remove(lcPreFix + "-" + lcKey)
                cl.Add(lcOldValues + "<ENTRY>" + lcLine, lcPreFix + "-" + lcKey)
            End If
        Next
    End Sub

    Public Function GetWindow(ByVal lcUrl As String, Optional ByVal lcHeight As String = "600px", Optional ByVal lcWidth As String = "800px") As String
        Dim lcScript As String = ""
        GetWindow = lcScript
        lcScript = ""
        lcScript = lcScript + "window.open('" + lcUrl + " ')"
        'lcScript = lcScript + AddAppPath(lcUrl + "',null,'dialogHeight:" + lcHeight + ";dialogLeft:100px;dialogTop:50px;dialogWidth:" + lcWidth + ";center:yes;dialogHide:yes;edge:raised;help:no;resizable:yes;scroll:no;status:yes;unadorned:yes' );")
        GetWindow = lcScript
    End Function

    Public Function Get_IP_Address() As String
        Dim context As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim sIPAddress As String = context.Request.ServerVariables("HTTP_X_FORWARDED_FOR")
        If String.IsNullOrEmpty(sIPAddress) Then
            Return context.Request.ServerVariables("REMOTE_ADDR")
        Else
            Dim ipArray As String() = sIPAddress.Split(New [Char]() {","c})
            Return ipArray(0)
        End If
    End Function

    Function GetCurrentUserID() As String
        GetCurrentUserID = UCase(GetSessionVariable(Session.SessionID, "UserID"))
    End Function

    Function GetUserCompany(ByVal lcUserId As String) As String
        GetUserCompany = ""
    End Function

    Public Function setCloseDialogJQ(ByVal lcDialogID As String) As String
        Return ("window.parent.$('" + "." + lcDialogID + "').dialog('close')")
    End Function

    Sub AddFormsStyles(ByRef p As System.Web.UI.Page)
        Dim laServerUrl As String() = Split(p.Request.Url.ToString, "/")
        Dim lcServerUrl As String = laServerUrl(0) + "//" + laServerUrl(2) + GetAppPath()
        AddClientCss(p, lcServerUrl + "~/css/Forms.css")
    End Sub

    Sub AlighGridColumns(ByRef r As GridViewRow, Optional startWith As Integer = 0)
        For i = startWith To r.Cells.Count - 1
            Dim lcColumnText As String = r.Cells(i).Text
            r.Cells(i).Text = Replace(lcColumnText, "|", "<BR>")
            r.Cells(i).HorizontalAlign = HorizontalAlign.Left
        Next
    End Sub

    Function Get_User_Name(ByVal lcUser As String) As String
        Get_User_Name = GetSqlColumn(EBDB_CS(), "Select USer_Name from EBLINKS_USERS where USER_ID='" + lcUser + "'")
    End Function

    Function Get_User_Department(ByVal lcUser As String) As String
        'Get_User_Department = GetSqlColumn(EBDB_CS(), "Select DEPARTMENT from WF_Users where ID='" + lcUser + "'")
        Get_User_Department = GetSqlColumn(EBDB_CS(), "Select DEPARTMENT_ID from EBLINKS_USERS where USER_ID='" + lcUser + "'")
    End Function

    Function GetUserDepartments(ByVal lcUser As String) As String
        Dim laDepartments = GetSqlArray(EBDB_CS(), "Select DepartmentID from L24_UserDepartments where UserID='" + lcUser + "'")
        Dim lcDepartments As String = "~" + Join(laDepartments, "~") + "~"
        GetUserDepartments = lcDepartments
    End Function

    Function LoadRsIntoCollection(ByVal lcDSN As String, ByVal lcSQL As String, ByVal lcPrefix As String, ByVal cl As Collection) As Boolean
        Dim i As Integer = 0
        LoadRsIntoCollection = False
        Dim dt As Data.DataTable = GetDataTable(lcDSN, lcSQL)
        If dt.Rows.Count > 0 Then
            Dim rs As Data.DataRow = dt.Rows(0)
            For Each c As Data.DataColumn In dt.Columns
                Try
                    If IsDBNull(rs(c.ColumnName)) Then
                        cl.Add("", lcPrefix + "." + UCase(c.ColumnName))
                    Else
                        cl.Add(rs(c.ColumnName), lcPrefix + "." + UCase(c.ColumnName))
                    End If
                Catch ex As Exception

                End Try
                LoadRsIntoCollection = True
            Next
        End If
        LoadRsIntoCollection = True
    End Function

    Sub DT2COLLECTION(ByRef DT As Data.DataTable, _
                  ByRef cl As Collection, _
                  ByVal lcPrefix As String, _
                  ByVal laKeys As String())
        Dim row As Data.DataRow
        'cl.Clear()
        For Each row In DT.Rows
            Dim lcKey As String = ""
            For ii = LBound(laKeys) To UBound(laKeys)
                lcKey = lcKey + IIf(lcKey <> "", "-", "") + FixNull(row.Item(laKeys(ii)))
            Next
            If lcPrefix <> "" Then lcKey = lcPrefix + "-" + lcKey + ""
            'Dim lcLine As String = lcKey + vbTab
            Dim lcLine As String = ""
            For ii = LBound(row.ItemArray) To UBound(row.ItemArray)
                lcLine = lcLine + CStr(FixNull(row.Item(ii))) + vbTab
            Next
            If cl.Contains(lcKey) Then
                Dim lcOldData As String = cl(lcKey)
                cl.Remove(lcKey)
                cl.Add(lcOldData + "^" + lcLine, lcKey)
            Else
                cl.Add(lcLine, UCase(lcKey))
            End If
        Next
    End Sub

    ' **********************************
    ' Function to check file Existance
    ' **********************************
    Function IsFileExists(ByVal FileName As String) As Boolean
        If FileName = "" Then
            IsFileExists = False
            Exit Function
        End If

        Dim objFSO As Object
        objFSO = CreateObject("Scripting.FileSystemObject")
        If (objFSO.FileExists(FileName)) = True Then
            IsFileExists = True
        Else
            IsFileExists = False
        End If
        objFSO = Nothing
    End Function

    Sub DataTable2CSV(ByVal table As Data.DataTable, ByVal filename As String, ByVal sepChar As String)
        Dim writer As System.IO.StreamWriter
        Try
            writer = New System.IO.StreamWriter(filename)
            ' first write a line with the columns name
            Dim sep As String = ""
            Dim builder As New System.Text.StringBuilder
            For Each col As Data.DataColumn In table.Columns
                builder.Append(sep).Append(col.ColumnName)
                sep = sepChar
            Next
            writer.WriteLine(builder.ToString())

            ' then write all the rows
            For Each row As Data.DataRow In table.Rows
                sep = ""
                builder = New System.Text.StringBuilder
                For Each col As Data.DataColumn In table.Columns
                    builder.Append(sep).Append(row(col.ColumnName))
                    sep = sepChar
                Next
                writer.WriteLine(builder.ToString())
            Next
        Finally
            If Not writer Is Nothing Then writer.Close()
        End Try
    End Sub

 

    Function DigitizeString(ByVal txt As String) As String
        Dim S As String = ""
        For i = 1 To Len(txt)
            S = S + Right("000" + CStr(Asc(Mid(txt, i, 1))), 3)
        Next
        DigitizeString = S
    End Function

    Function UnDigitizeString(ByVal txt As String) As String
        Dim S As String = ""
        For i = 1 To Len(txt) Step 3
            S = S + Chr(Val(Mid(txt, i, 3)))
        Next
        UnDigitizeString = S
    End Function
    Function GetFileName(ByVal lcPath As String) As String
        GetFileName = ""
        If lcPath <> "" Then
            lcPath = Replace(lcPath, "/", "\")
            Dim a = Split(lcPath, "\")
            GetFileName = a(UBound(a))
        End If
    End Function

    Function GetFilesFolder(ByVal lcFolder As String) As String
        GetFilesFolder = GetAppPath() + "/files/" + lcFolder + "/"
    End Function

    Function GetSystemVariable(lcVariable As String)
        Dim value As String
        value = Environment.GetEnvironmentVariable("windir")
        ' If necessary, create it.
        If value Is Nothing Then
            'Environment.SetEnvironmentVariable("Test1", "Value1")
            value = Environment.GetEnvironmentVariable("Test1")
        End If
    End Function

    Sub SetConfirmBox(ByVal ob As Object, ByVal lcMsg As String)
        Dim lcScript As String = ""
        lcScript = lcScript + "return confirm('" & lcMsg & "');"
        Replace(lcScript, "[", Chr(34))
        Replace(lcScript, "]", Chr(34))
        ob.Attributes.Add("onclick", lcScript)
    End Sub

    Sub Add_Char(ByRef c As Collection, ByVal lcKey As String, ByVal v As Object, Optional lnLen As Integer = 0)
        Dim laKeys As String() = Split(lcKey, ";")
        For Each lcKey1 In laKeys
            If lcKey1 <> "" Then
                Dim cp As New CollectionPair
                cp.TheKey = lcKey1
                v = Replace(v, "&", "")
                cp.TheValue = "'" + v + "'"
                If lnLen <> 0 Then cp.TheValue = Mid(v, 1, lnLen)
                If c.Contains(lcKey1) Then
                    c.Remove(lcKey1)
                End If
                c.Add(cp, lcKey1)
            End If
        Next
    End Sub

    Sub Add_Numb(ByRef c As Collection, ByVal lcKey As String, ByVal v As Object, Optional lnLen As Integer = 0)
        Dim laKeys As String() = Split(lcKey, ";")
        For Each lcKey1 In laKeys
            If lcKey1 <> "" Then
                Dim cp As New CollectionPair
                cp.TheKey = lcKey1
                v = Replace(v, "&", "")
                cp.TheValue = "0" + CStr(v)
                If lnLen <> 0 Then cp.TheValue = Mid(v, 1, lnLen)
                c.Add(cp, lcKey1)
            End If
        Next
    End Sub

    Function GetExcelSheet(lcXlFile As String, ByVal lcSheet As String, Optional HeaderIsAvailable As Boolean = True) As System.Data.DataTable
        Dim MyConnection As System.Data.OleDb.OleDbConnection
        Dim DtSet As System.Data.DataSet
        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
        ';Password=password
        ' Excel 12.0;Imex=2
        'Dim lcConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;" & " Data Source='" + lcXlFile + "'; " & "Extended Properties=[Excel 12.0 XML" + If(HeaderIsAvailable, "", ";HDR=NO") + "];"
        Dim lcConnectionString = ""
        If InStr(LCase(lcXlFile), ".xlsx") > 0 Then
            lcConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;" & " Data Source='" + lcXlFile + "'; " & "Extended Properties=[Excel 12.0 XML" + If(HeaderIsAvailable, "", ";HDR=NO") + ";IMEX=1];"
        Else
            lcConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;" & " Data Source='" + lcXlFile + "'; " & "Extended Properties=[Excel 12.0" + If(HeaderIsAvailable, "", ";HDR=NO") + ";IMEX=1];"
        End If
        lcConnectionString = Replace(lcConnectionString, "[", Chr(34))
        lcConnectionString = Replace(lcConnectionString, "]", Chr(34))
        MyConnection = New System.Data.OleDb.OleDbConnection(lcConnectionString)
        If lcSheet = "" Then
            MyConnection.Open()
            For Each rw1 In MyConnection.GetSchema("Tables").Rows
                If InStr(UCase(rw1("TABLE_NAME")), "FILTERDATABASE") = 0 Then
                    lcSheet = MyConnection.GetSchema("Tables").Rows(0)("TABLE_NAME")
                    lcSheet = Replace(lcSheet, "$", "")
                End If
            Next
            'lcSheet = MyConnection.GetSchema("Tables").Rows(0)("TABLE_NAME")
            'lcSheet = Replace(lcSheet, "$", "")
            MyConnection.Close()
        End If
        If lcSheet = "" Then

        End If
        Dim cmdExcel As OleDbCommand = New OleDbCommand()
        cmdExcel.Connection = MyConnection
        cmdExcel.CommandText = "SELECT * From [" + lcSheet + "$]"
        Dim da As New System.Data.OleDb.OleDbDataAdapter
        da.SelectCommand = cmdExcel
        Dim ds As DataSet = New DataSet()
        da.Fill(ds)
        GetExcelSheet = ds.Tables(0)
        MyConnection.Close()



        'MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [" + lcSheet + "$]", MyConnection)
        'MyCommand.TableMappings.Add("Table", "TestTable")
        'DtSet = New System.Data.DataSet
        'MyCommand.Fill(DtSet)
        'GetExcelSheet = DtSet.Tables(0)
        'MyCommand.Dispose()
        'MyConnection.Close()

        'Try
        '    MyConnection.Dispose()
        'Catch ex As Exception

        'End Try
        'Try
        '    DtSet.Dispose()
        'Catch ex As Exception

        'End Try

    End Function


    'Function GetExcelSheet(lcXlFile As String, ByVal lcSheet As String, Optional HeaderIsAvailable As Boolean = True) As System.Data.DataTable
    '    Dim MyConnection As System.Data.OleDb.OleDbConnection
    '    Dim DtSet As System.Data.DataSet

    '    ';Password=password
    '    ' Excel 12.0;Imex=2
    '    'Dim lcConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;" & " Data Source='" + lcXlFile + "'; " & "Extended Properties=[Excel 12.0 XML" + If(HeaderIsAvailable, "", ";HDR=NO") + "];"
    '    Dim lcConnectionString = ""
    '    If InStr(LCase(lcXlFile), ".xlsx") > 0 Then
    '        lcConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;" & " Data Source='" + lcXlFile + "'; " & "Extended Properties=[Excel 12.0 XML" + If(HeaderIsAvailable, "", ";HDR=NO") + ";IMEX=1];"
    '    Else
    '        lcConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;" & " Data Source='" + lcXlFile + "'; " & "Extended Properties=[Excel 12.0" + If(HeaderIsAvailable, "", ";HDR=NO") + ";IMEX=1];"
    '    End If
    '    lcConnectionString = Replace(lcConnectionString, "[", Chr(34))
    '    lcConnectionString = Replace(lcConnectionString, "]", Chr(34))

    '    MyConnection = New System.Data.OleDb.OleDbConnection(lcConnectionString)
    '    If lcSheet = "" Then
    '        MyConnection.Open()
    '        lcSheet = MyConnection.GetSchema("Tables").Rows(0)("TABLE_NAME")
    '        lcSheet = Replace(lcSheet, "$", "")
    '    End If
    '    Dim MyCommand As New System.Data.OleDb.OleDbDataAdapter
    '    MyCommand.SelectCommand = New OleDbCommand("select * from [" + lcSheet + "$]", )
    '    'MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [" + lcSheet + "$]", MyConnection)
    '    MyCommand.TableMappings.Add("Table", "TestTable")
    '    Dim  data as DataTable= new DataTable("MyTable");
    '    DtSet = New System.Data.DataSet
    '    MyCommand.Fill(DtSet)
    '    GetExcelSheet = DtSet.Tables(0)
    '    MyCommand.Dispose()
    '    MyConnection.Close()
    '    Try
    '        MyConnection.Dispose()
    '    Catch ex As Exception

    '    End Try
    '    Try
    '        DtSet.Dispose()
    '    Catch ex As Exception

    '    End Try

    'End Function
    Function Humanize(ByVal lcText As String)
        lcText = ReplaceAll(lcText, "~~", "~")
        Dim laText As String() = Split(lcText, "~")
        Dim lcText1 As String = ""
        For i = LBound(laText) To UBound(laText)
            If laText(i) <> "" Then
                lcText1 = lcText1 + IIf(lcText1 = "", "", "~") + laText(i)
            End If
        Next
        laText = Split(lcText1, "~")
        lcText1 = ""
        If laText.Length > 1 Then
            For i = LBound(laText) To UBound(laText) - 1
                lcText1 = lcText1 + IIf(lcText1 = "", "", ", ") + laText(i)
            Next
            lcText1 = lcText1 + " and " + laText(UBound(laText))
            Humanize = lcText1
        Else
            Humanize = Replace(lcText, "~", " ")
        End If
    End Function

    Sub FillRadio(ByVal a As Array, ByRef CB As RadioButtonList, Optional ByVal lcSelected As String = "")
        Dim i As Integer
        Dim lcItemText As String
        Dim lcItemValue As String
        CB.Items.Clear()
        For i = LBound(a) To UBound(a)
            If a(i) <> "" Then
                a(i) = Replace(a(i), "^", vbTab)
                lcItemValue = Split(a(i), vbTab)(0)
                Try
                    lcItemText = Split(a(i), vbTab)(1)
                Catch
                    lcItemText = lcItemValue
                End Try
                CB.Items.Add(lcItemText)
                CB.Items(i).Value = lcItemValue
                CB.Items(i).Selected = False
                If UCase(lcSelected) = UCase(lcItemValue) Then
                    CB.Items(i).Selected = True
                End If
            End If
        Next
    End Sub


    Function GetSelectedValues(ByVal cbs As CheckBoxList) As String
        Dim lcValues As String = ""
        For i = 0 To cbs.Items.Count - 1
            If cbs.Items(i).Selected Then
                lcValues = lcValues + IIf(lcValues = "", "", "~") + cbs.Items(i).Value
            End If
        Next
        GetSelectedValues = "~" + lcValues + "~"
    End Function

    Function SetSelectedValues(ByVal cbs As CheckBoxList, ByVal lcValues As String)
        If lcValues = "~~" Or lcValues = "" Then Exit Function
        For i = 0 To cbs.Items.Count - 1
            If InStr(lcValues, "~" + cbs.Items(i).Value + "~") > 0 Then
                cbs.Items(i).Selected = True
            End If
        Next
    End Function

    Function IsAddMode(lcMode As String) As Boolean
        IsAddMode = False
        If lcMode = "NEW" Or _
            lcMode = "ADD" Then
            IsAddMode = True
        End If
    End Function

    Function SaveRecord(ByVal CS As String, mr As Collection, lcMode As String, ByVal lcTable As String, ByVal lcWhere As String, ByRef lcMessage As String) As Boolean
        SaveRecord = False
        Dim lcSearchSql As String = "Select * from " + lcTable + IIf(lcWhere = "", "", " WHERE " + lcWhere)
        If InStr("ADD;NEW,COPY", UCase(lcMode)) Then
            If CheckIfFound(CS, lcSearchSql) Then
                lcMessage = "Record Already Exist"
                Exit Function
            End If
            InsertTableRecord(CS, mr, lcTable)
        Else                            '   Edit mode
            If Not CheckIfFound(CS, lcSearchSql) Then
                lcMessage = "Record dose not Exist"
                Exit Function
            End If
            UpdateTableRecord(CS, mr, lcTable, lcWhere)
        End If
        SaveRecord = True
    End Function

    Function GetFirstName(ByVal lcName As String) As String
        Dim laName As String() = Split(lcName, " ")
        GetFirstName = laName(0)
    End Function

    Function GetNextID_New(ByVal CS As String, ByVal lcTable As String, ByVal lcField As String, ByVal lnLength As Integer, Optional lnIncrement As Integer = 1) As String
        Dim lcSql As String = "Select Max(" + lcField + ") as MAXID from " + lcTable
        Dim dt As Data.DataTable = GetDataTable(CS, lcSql)
        Dim lnNextID As Integer = lnIncrement
        Try
            lnNextID = Val(CStr(dt.Rows(0)(0)))
            lnNextID = lnNextID + lnIncrement
        Catch ex As Exception
        End Try
        GetNextID_New = Right("00000000000000000000" + CStr(lnNextID), lnLength)
    End Function

    Function GetNextID(ByVal CS As String, ByVal lcSql As String, ByVal lnLength As Integer) As String
        Dim dt As System.Data.DataTable = GetDataTable(CS, lcSql)
        Dim lnNextID As Integer = 1
        Try
            lnNextID = Val(CStr(dt.Rows(0)(0)))
            lnNextID = lnNextID + 1
        Catch ex As Exception

        End Try
        GetNextID = Right("00000000000000000000" + CStr(lnNextID), lnLength)
    End Function

    Public Function getDataRow(ByVal ConnectionString As String, sSql As String) As Data.DataRow
        If ConnectionString = "" Then ConnectionString = EBDB_CS()
        Dim oc As New System.Data.OleDb.OleDbConnection(ConnectionString)
        Dim dt As System.Data.DataTable = New System.Data.DataTable
        Dim oda As OleDbDataAdapter = New OleDbDataAdapter(sSql, oc)
        oda.Fill(dt)
        If dt.Rows.Count > 0 Then
            Return (dt.Rows(0))
        Else
            Return (Nothing)
        End If
    End Function


    Function CheckIfFound(ByVal CS As String, ByVal lcSql As String) As Boolean
        CheckIfFound = False
        Dim dt As Data.DataTable = GetDataTable(CS, lcSql)
        If dt.Rows.Count > 0 Then
            CheckIfFound = True
        End If
    End Function

    Function GetLookupArray(ByVal lcLookupGroup As String) As String()
        Dim dt As Data.DataTable = GetDataTable(EBDB_CS, "Select Lkval,lkText from eblib_lrmLkp00pf where lkGID='" + lcLookupGroup + "'")
        Dim lcLines As String = ""
        For Each r In dt.Rows
            lcLines = lcLines + IIf(lcLines <> "", vbCrLf, "") + r(0) + vbTab + r(1)
        Next
        GetLookupArray = Split(lcLines, vbCrLf)
    End Function

    Function GetExtractFilesFolder(ByVal lcSubFolder As String) As String
        GetExtractFilesFolder = GetAppPath() + "/files/" + lcSubFolder + "/"
    End Function

    Function GetFilesList(ByVal lcFolder As String, ByVal lcType As String, Optional lcPrefix As String = "") As Data.DataTable

        'Dim x As Data.DataView
        Dim DT As New Data.DataTable()
        Dim Row1 As Data.DataRow
        DT.Columns.Add(New Data.DataColumn("NAME", System.Type.GetType("System.String")))
        DT.Columns.Add(New Data.DataColumn("Date", System.Type.GetType("System.String")))
        DT.Columns.Add(New Data.DataColumn("Size", System.Type.GetType("System.String")))
        Dim lcPath As String = Server.MapPath(GetExtractFilesFolder(lcFolder))
        Dim lcFile As String = Dir(lcPath + "\" + lcPrefix + "*" + lcType)
        Do While lcFile <> ""
            Dim lcFileName As String = lcPath + "\" + lcFile
            Dim lnLen As Integer = FileLen(lcFileName)
            Row1 = DT.NewRow()
            Row1("NAME") = lcFile
            Row1("DATE") = FileDateTime(lcFileName)
            Row1("SIZE") = CStr(lnLen / 1000) + " KB"
            DT.Rows.Add(Row1)
            lcFile = Dir()
        Loop
        'DT.DefaultView.Sort("DESC NAME")
        DT.DefaultView.Sort = "NAME DESC"
        GetFilesList = DT

    End Function

    'Function ShowMessages(ByVal p As System.Web.UI.Page, lcmessage As String, Optional isError As Boolean = False)
    '    For Each msg In Split(lcmessage, vbCrLf)
    '        If msg <> "" Then ShowMessage(p, msg, isError)
    '    Next
    'End Function

    'Function Show_Messages(ByVal p As System.Web.UI.Page, clMsg As Collection, Optional isError As Boolean = False)
    '    For Each msg In clMsg
    '        If msg <> "" Then ShowMessage(p, msg, isError)
    '    Next
    'End Function

    'Public Function ShowMessage(ByVal p As System.Web.UI.Page, ByVal msg As String, Optional isError As Boolean = False) As String
    '    msg = Replace(msg, vbCrLf, " ")
    '    msg = Replace(msg, vbCr, " ")
    '    msg = Replace(msg, vbLf, " ")
    '    Dim myScript As String = "ShowMessage('" + msg + "');"
    '    If isError Then
    '        myScript = "ShowError('" + msg + "');"
    '    End If
    '    Dim lcGUID As String = Guid.NewGuid().ToString
    '    ScriptManager.RegisterClientScriptBlock(p, p.GetType(), lcGUID, myScript, True)
    'End Function

    Function GetLookup(ByVal lcKey As String) As String
        GetLookup = ""
        If Lookups.Contains(lcKey) Then GetLookup = Lookups(UCase(lcKey))
    End Function

    Public psUser, psLevel, psDept, psSCN, psDLR, psDLP, psBRNM As String
    Class CollectionPair
        Public TheValue As Object
        Public TheKey As String
    End Class

    Function GetLookups(ByVal lcGroup As String) As String()
        GetLookups = GetSqlArray(EBDB_CS, "Select LKVAL,LKTEXT from EB_Lookups Where upper(LTRIM(RTRIM(LKGID)))='" + UCase(lcGroup) + "'")
    End Function

    Sub DisableButtonOnClick(ByVal bt As System.Web.UI.WebControls.Button)
        Dim lcScript As String = ""
        lcScript = lcScript + "this.value=this.value+', Please Wait..';"
        lcScript = lcScript + "this.disabled=true;"
        Replace(lcScript, "[", Chr(34))
        Replace(lcScript, "]", Chr(34))
        lcScript = lcScript + ClientScript.GetPostBackEventReference(bt, "").ToString()
        bt.Attributes.Add("onclick", lcScript)
    End Sub

    Sub DisableLinkOnClick(ByVal bt As System.Web.UI.WebControls.LinkButton)
        Dim lcScript As String = ""
        lcScript = lcScript + "this.text=this.text+', Please Wait..';"
        lcScript = lcScript + "this.disabled=true;"
        Replace(lcScript, "[", Chr(34))
        Replace(lcScript, "]", Chr(34))
        lcScript = lcScript + ClientScript.GetPostBackEventReference(bt, "").ToString()
        bt.Attributes.Add("onclick", lcScript)
    End Sub

    Function Get_Application_Id(ByVal lcShort_Name As String) As String
        Get_Application_Id = getDataColumn(EBDB_CS, "Select APP_ID from EBLINKS_Apps where upper(Short_Name)='" + UCase(lcShort_Name) + "' ")
    End Function
    Function Get_User_Level(ByVal lcUser_Id As String, ByVal lcApplication_id As String) As String
        Get_User_Level = ""
        Dim lcLevel As String = ""
        'lcLevel = getDataColumn(EBDB_CS, "Select Access_Level from EBLINKS_Apps_Access where APP_ID='" + lcApplication_id + "' and User_Id='" + lcUser_Id + "'")
        'ACCESS_RULES
        Dim rs As Data.DataRow = getDataRow(EBDB_CS, "Select Access_Level,ACCESS_RULES from EBLINKS_Apps_Access where APP_ID='" + lcApplication_id + "' and User_Id='" + lcUser_Id + "'")
        If IsNothing(rs) Then
            Get_User_Level = "U"
            Exit Function
        Else
            lcLevel = rs("Access_Level")
        End If
        If lcLevel = "" Then lcLevel = "U"
        If FixNull(rs("ACCESS_RULES")) = "" Then
            Get_User_Level = lcLevel
        Else
            Get_User_Level = UCase(rs("ACCESS_RULES"))
        End If
        'Get_User_Level = lcLevel
    End Function
    Function Get_User_ACCESS_RULES(ByVal lcApp_short_name As String, ByVal lcUser_id As String) As String
        Dim lcApp_id As String = Get_Application_Id(lcApp_short_name)
        Dim lcSql As String = "Select ACCESS_RULES from EBLINKS_APPS_ACCESS where App_Id='" + lcApp_id + "' and USER_ID='" + lcUser_id + "'"
        Get_User_ACCESS_RULES = getDataColumn(EBDB_CS, lcSql)
    End Function


    Function GetUserLevel(ByVal lcUserid As String, ByVal lcSystem As String) As String
        Dim lcLevel As String = ""
        lcLevel = getDataColumn(EBDB_CS, "Select SACLEVEL from EBLIB_EBSACPF where SACSYS='" + lcSystem + "' and SACUID='" + lcUserid + "'", "SACLEVEL")
        If lcLevel = "" Then
            lcLevel = "U"
        End If
        GetUserLevel = lcLevel
    End Function

    Sub AddJQueryLinks(ByRef p As Web.UI.Page, Optional llAddFontsOwsome As Boolean = False)
        Dim laServerUrl As String() = Split(p.Request.Url.ToString, "/")
        'Dim lcServerUrl As String = laServerUrl(0) + "//" + laServerUrl(2) + GetAppPath()
        Dim lcServerUrl As String = Get_Server_Url()
        AddClientCss(p, lcServerUrl + "/JQ/jquery-ui.css?" + Format(Now, "hhMMss"))
        AddClientScript(p, lcServerUrl + "/JQ/jquery.js")
        AddClientScript(p, lcServerUrl + "/JQ/jquery-ui.js")
        AddClientScript(p, lcServerUrl + "/JQ/Dialogs/jquery.dialogextend.js")
        AddClientScript(p, lcServerUrl + "/JQ/Dialogs/jsDialog.js?11124")
        AddClientScript(p, lcServerUrl + "/JQ/jquery.blockUI.js")

        AddClientCss(p, lcServerUrl + "/JQ/toastr/toastr.css")
        AddClientScript(p, lcServerUrl + "/JQ/toastr/toastr.js")
        AddClientScript(p, lcServerUrl + "/JQ/Masks/masks.js")
        If llAddFontsOwsome = True Then
            AddClientCss(p, lcServerUrl + "/css/font-awesome-4.7.0/css/font-awesome.min.css")
        End If
        '        ppp = p
    End Sub

    Sub AddClientCss(ByVal ThePage As Web.UI.Page, ByVal CssFile As String)
        Dim css As HtmlLink = New HtmlLink()
        css.Href = CssFile
        css.Attributes("rel") = "stylesheet"
        css.Attributes("type") = "text/css"
        css.Attributes("media") = "all"
        ThePage.Header.Controls.Add(css)
    End Sub

    Sub AddClientScript(ByVal ThePage As Web.UI.Page, ByVal ScriptFile As String)
        Dim js As HtmlGenericControl = New HtmlGenericControl("script")
        js.Attributes("type") = "text/javascript"
        js.Attributes("src") = ScriptFile
        ThePage.Header.Controls.Add(js)
    End Sub

    Sub AppendClientScript(ByVal ThePage As Web.UI.Page, ByVal ScriptFile As String)
        Dim js As HtmlGenericControl = New HtmlGenericControl("script")
        js.Attributes("type") = "text/javascript"
        js.Attributes("src") = ScriptFile
    End Sub

    Public Function CloseDialogJQ(p As Web.UI.Page, Optional lcDialogID As String = "") As String
        If lcDialogID = "" Then
            Dim context As System.Web.HttpContext = System.Web.HttpContext.Current
            lcDialogID = context.Request("DIALOG_ID")
        End If
        If lcDialogID <> "" Then
            Dim CSM As ClientScriptManager = p.ClientScript
            CSM.RegisterStartupScript(Me.GetType, Guid.NewGuid().ToString(), "<script language=""JavaScript"">CloseDialog('" + lcDialogID + "');</script>")
        Else
            Dim CSM As ClientScriptManager = p.ClientScript
            CSM.RegisterStartupScript(Me.GetType, Guid.NewGuid().ToString(), "<script language=""JavaScript"">GoBack();</script>")
        End If
        CloseDialogJQ = ""
    End Function

    Function AddAppPath(ByVal lcPath As String) As String
        If Left(lcPath, 1) = "/" Then
            AddAppPath = GetAppPath() + lcPath
        Else
            AddAppPath = lcPath
        End If
    End Function

    Public Function SetOnClientClick(bt As Web.UI.WebControls.Button, ByVal lcUrl As String, Optional ByVal lnHeight As Integer = 600, Optional ByVal lnWidth As Integer = 800, Optional lcTitle As String = "", Optional SidePannel As Boolean = False) As String
        Dim cc As String = Replace(ClientScript.GetPostBackEventReference(bt, "").ToString(), "'", Chr(34))
        If SidePannel Then
            bt.OnClientClick = CreateSidePannel(lcUrl, lnWidth, lcTitle, cc)
        Else
            bt.OnClientClick = GetMntDialogJQ(lcUrl, lnHeight, lnWidth, lcTitle, cc)
        End If
    End Function

    Public Function SetOnClientClick(bt As ImageButton, ByVal lcUrl As String, Optional ByVal lnHeight As Integer = 600, Optional ByVal lnWidth As Integer = 800, Optional lcTitle As String = "", Optional SidePannel As Boolean = False) As String
        Dim cc As String = Replace(ClientScript.GetPostBackEventReference(bt, "").ToString(), "'", Chr(34))
        If SidePannel Then
            bt.OnClientClick = CreateSidePannel(lcUrl, lnWidth, lcTitle, cc)
        Else
            bt.OnClientClick = GetMntDialogJQ(lcUrl, lnHeight, lnWidth, lcTitle, cc)
        End If
    End Function

    Public Function SetOnClientClick(bt As Web.UI.WebControls.LinkButton, ByVal lcUrl As String, Optional ByVal lnHeight As Integer = 600, Optional ByVal lnWidth As Integer = 800, Optional lcTitle As String = "", Optional SidePannel As Boolean = False) As String
        Dim cc As String = Replace(ClientScript.GetPostBackEventReference(bt, "").ToString(), "'", Chr(34))
        If SidePannel Then
            bt.OnClientClick = CreateSidePannel(lcUrl, lnWidth, lcTitle, cc)
        Else
            bt.OnClientClick = GetMntDialogJQ(lcUrl, lnHeight, lnWidth, lcTitle, cc)
        End If
    End Function

    Public Function CreateSidePannel(ByVal lcUrl As String, Optional ByVal lnWidth As Integer = 800, Optional lcTitle As String = "", Optional callBackCode As String = "") As String
        Dim lcScript As String = ""
        Dim lcDialogId As String = Guid.NewGuid().ToString
        CreateSidePannel = lcScript
        Dim lcWidth As String = CStr(lnWidth)
        lcScript = ""
        lcUrl = Replace(lcUrl, "[AND]", "&")
        If InStr(lcUrl, "?") > 0 Then
            lcUrl = lcUrl + "&DIALOG_ID=" + lcDialogId
        Else
            lcUrl = lcUrl + "?DIALOG_ID=" + lcDialogId
        End If
        lcScript = lcScript + "return CreateSidePannel('" + lcTitle + "','" + lcUrl + "'," + lcWidth + ",'" + callBackCode + "','" + lcDialogId + "')"
        CreateSidePannel = lcScript
    End Function

    Public Function GetMntDialogJQ(ByVal lcUrl As String, Optional ByVal lnHeight As Integer = 600, Optional ByVal lnWidth As Integer = 800, Optional lcTitle As String = "", Optional callBackCode As String = "") As String
        Dim lcScript As String = ""
        Dim lcDialogId As String = Guid.NewGuid().ToString
        GetMntDialogJQ = lcScript
        Dim lcHeight As String = CStr(lnHeight)
        Dim lcWidth As String = CStr(lnWidth)
        lcScript = ""
        lcUrl = Replace(lcUrl, "[AND]", "&")
        If InStr(lcUrl, "?") > 0 Then
            lcUrl = lcUrl + "&DIALOG_ID=" + lcDialogId
        Else
            lcUrl = lcUrl + "?DIALOG_ID=" + lcDialogId
        End If
        lcScript = lcScript + "return ShowInDialog('" + lcTitle + "','" + lcUrl + "'," + lcHeight + "," + lcWidth + ",'" + callBackCode + "','" + lcDialogId + "')"
        GetMntDialogJQ = lcScript
    End Function


    Public Function GetMntWindowJQ(ByVal lcUrl As String, Optional ByVal lnHeight As Integer = 600, Optional ByVal lnWidth As Integer = 800, Optional lcTitle As String = "", Optional callBackCode As String = "") As String
        Dim lcScript As String = ""
        Dim lcDialogId As String = Guid.NewGuid().ToString
        GetMntWindowJQ = lcScript
        Dim lcHeight As String = CStr(lnHeight)
        Dim lcWidth As String = CStr(lnWidth)
        lcScript = ""
        lcUrl = Replace(lcUrl, "[AND]", "&")
        If InStr(lcUrl, "?") > 0 Then
            lcUrl = lcUrl + "&DIALOG_ID=" + lcDialogId
        Else
            lcUrl = lcUrl + "?DIALOG_ID=" + lcDialogId
        End If
        lcScript = lcScript + "return ShowInWindow('" + lcTitle + "','" + lcUrl + "'," + lcHeight + "," + lcWidth + ",'" + callBackCode + "','" + lcDialogId + "')"
        GetMntWindowJQ = lcScript
    End Function

    Public Sub CloseThisWindows(ByRef P As Web.UI.Page)
        Dim CSM As ClientScriptManager = P.ClientScript
        CSM.RegisterStartupScript(Me.GetType, Guid.NewGuid().ToString(), "<script language=""JavaScript"">top.window.close();</script>")
    End Sub

    'Dim oc As New OleDbConnection("Provider=IBMDA400.1;Data Source=192.168.1.3;User ID=inap;Password=inapinap")
    '  Public pAS400 As New cwbx.AS400System

    Function CheckDataRow(lcConnectionString As String, lcSql As String) As Boolean
        CheckDataRow = False
        Dim dt As Data.DataTable = GetDataTable(lcConnectionString, lcSql)
        If dt.Rows.Count > 0 Then
            CheckDataRow = True
        End If
    End Function

    Sub GetFieldsAndSetValues(c As Collection, ByRef lcSetFlds As String)
        lcSetFlds = ""
        For k = 1 To c.Count
            Dim myPair As CollectionPair = c(k)
            Dim lcVal As String = CStr(myPair.TheValue)
            If Mid(lcVal, 1, 1) = "'" Then
                lcVal = "'" + Strings.Replace(lcVal, "'", "") + "'"
            End If
            lcSetFlds = lcSetFlds + IIf(lcSetFlds = "", "", ",") + myPair.TheKey + "=" + lcVal
        Next
    End Sub

    Sub UpdateTableRecord(ByVal lcConnectionString As String, ByRef c As Collection, ByVal lcTable As String, lcCondition As String)
        Dim oc As New OleDb.OleDbConnection(lcConnectionString)
        oc.Open()
        Dim lcSetFlds As String = ""
        GetFieldsAndSetValues(c, lcSetFlds)
        Dim lcSql2 As String = "Update " + lcTable + " SET " + lcSetFlds + IIf(lcCondition = "", "", " Where " + lcCondition + " ")
        Dim cmd2 As New OleDb.OleDbCommand(lcSql2, oc)
        cmd2.ExecuteNonQuery()
        cmd2.Dispose()
        oc.Close()
    End Sub

    Public Function RunSql(ByVal lcConnectonString As String, ByVal lcSQL As String) As Boolean
        Dim oc As New OleDb.OleDbConnection(lcConnectonString)
        oc.Open()
        Dim cmd2 As New OleDb.OleDbCommand(lcSQL, oc)
        cmd2.ExecuteNonQuery()
        cmd2.Dispose()
        oc.Close()
        RunSql = True
    End Function

    Public Function RunSql_OC(ByRef oc As OleDb.OleDbConnection, ByVal lcSQL As String) As Boolean
        Dim cmd2 As New OleDb.OleDbCommand(lcSQL, oc)
        cmd2.ExecuteNonQuery()
        cmd2.Dispose()
    End Function
    Function Get_Add_Command(ByRef c As Collection, ByVal lcTable As String)
        Dim lcVals As String = ""
        Dim lcFlds As String = ""
        GetFieldsAndValues(c, lcFlds, lcVals)
        Dim lcSql2 As String = "Insert into " + lcTable + " (" + lcFlds + ") values(" + lcVals + ") "
        Get_Add_Command = lcSql2
    End Function

    Sub InsertTableRecord(ByVal lcConnectonString As String, ByRef c As Collection, ByVal lcTable As String)
        Dim oc As New OleDb.OleDbConnection(If(lcConnectonString <> "", lcConnectonString, EBDB_CS()))
        oc.Open()
        Dim lcVals As String = ""
        Dim lcFlds As String = ""
        GetFieldsAndValues(c, lcFlds, lcVals)
        Dim lcSql2 As String = "Insert into " + lcTable + " (" + lcFlds + ") values(" + lcVals + ") "
        Dim cmd2 As New OleDb.OleDbCommand(lcSql2, oc)
        cmd2.ExecuteNonQuery()
        cmd2.Dispose()
        oc.Close()
    End Sub

    Sub GetFieldsAndValues(c As Collection, ByRef lcFlds As String, ByRef lcVals As String)
        lcFlds = ""
        lcVals = ""
        For k = 1 To c.Count
            Dim myPair As CollectionPair = c(k)
            lcFlds = lcFlds + IIf(lcFlds = "", "", ",") + myPair.TheKey
            Dim lcVal As String = CStr(myPair.TheValue)
            If Mid(lcVal, 1, 1) = "'" Then
                lcVal = "'" + Strings.Replace(lcVal, "'", "") + "'"
            End If
            lcVals = lcVals + IIf(lcVals = "", "", ",") + lcVal
        Next
    End Sub

    Sub Add2Collection(ByRef c As Collection, ByVal v As Object, ByVal lcKey As String, Optional lnLen As Integer = 0)
        Dim laKeys As String() = Split(lcKey, ";")
        For Each lcKey1 In laKeys
            If lcKey1 <> "" Then
                Dim cp As New CollectionPair
                cp.TheKey = lcKey1
                cp.TheValue = v
                If lnLen <> 0 Then cp.TheValue = Mid(v, 1, lnLen)
                c.Add(cp, lcKey1)
            End If
        Next
    End Sub

    Sub Add2Collection_Numb(ByRef c As Collection, ByVal v As Object, ByVal lcKey As String, Optional lnLen As Integer = 0)
        Dim laKeys As String() = Split(lcKey, ";")
        For Each lcKey1 In laKeys
            If lcKey1 <> "" Then
                If v = "" Then v = "0"
                Dim cp As New CollectionPair
                cp.TheKey = lcKey1
                cp.TheValue = v
                If lnLen <> 0 Then cp.TheValue = Mid(v, 1, lnLen)
                c.Add(cp, lcKey1)
            End If
        Next
    End Sub

    Sub Add2Collection_Char(ByRef c As Collection, ByVal v As Object, ByVal lcKey As String, Optional lnLen As Integer = 0)
        Dim laKeys As String() = Split(lcKey, ";")
        For Each lcKey1 In laKeys
            If lcKey1 <> "" Then
                Dim cp As New CollectionPair
                cp.TheKey = lcKey1
                v = Replace(v, "&", "")
                cp.TheValue = "'" + v + "'"
                If lnLen <> 0 Then cp.TheValue = Mid(v, 1, lnLen)
                c.Add(cp, lcKey1)
            End If
        Next
    End Sub

    'Public Function SetOnClientClickPage(bt As LinkButton, ByVal lcUrl As String) As String
    '    Dim cc As String = Replace(ClientScript.GetPostBackEventReference(bt, "").ToString(), "'", Chr(34))
    '    bt.Attributes.Add("href", lcUrl)
    'End Function


    Public Function SetOnClientClickPage(bt As LinkButton, ByVal lcUrl As String) As String
        Dim cc As String = Replace(ClientScript.GetPostBackEventReference(bt, "").ToString(), "'", Chr(34))
        'bt.Attributes.Add("href", lcUrl)
        bt.OnClientClick = GetWindow(lcUrl)
    End Function

    'Public Function SetOnClientClickPage(bt As System.Web.UI.WebControls.HyperLink, ByVal lcUrl As String) As String
    '    Dim cc As String = Replace(ClientScript.GetPostBackEventReference(bt, "").ToString(), "'", Chr(34))
    '    bt.NavigateUrl = GetWindow(lcUrl)
    '    'bt.Attributes.Add("href", lcUrl)
    'End Function

    Public Function SetOnClientClickPage(bt As ImageButton, ByVal lcUrl As String) As String
        Dim cc As String = Replace(ClientScript.GetPostBackEventReference(bt, "").ToString(), "'", Chr(34))
        bt.OnClientClick = GetWindow(lcUrl)
        'bt.Attributes.Add("href", lcUrl)
    End Function
    Function SaveSessionVariable(ByVal lcSID As String, lcVar As String, lcVal As String, Optional lcSource As String = "", Optional Valid_Until As String = "")
        'lcSID = GetCurrentUser()
        lcSID = GetWindowsUser()
        Dim mr As New Collection
        If Not CheckDataRow(EBDB_CS, "Select * From EB_SessionsVariables where SID='" + lcSID + "' AND SVAR='" + lcVar + "'") Then
            Add2Collection_Char(mr, lcSID, "SID")
            Add2Collection_Char(mr, lcVar, "SVAR")
            Add2Collection_Char(mr, lcVal, "SVAL")
            Add2Collection_Char(mr, lcSource, "Source")
            Add2Collection_Char(mr, Valid_Until, "VALID_UNTIL")
            InsertTableRecord(EBDB_CS, mr, "EB_SessionsVariables")
        Else
            Add2Collection_Char(mr, lcVal, "SVAL")
            Add2Collection_Char(mr, lcSource, "Source")
            Add2Collection_Char(mr, Valid_Until, "VALID_UNTIL")
            UpdateTableRecord(EBDB_CS, mr, "EB_SessionsVariables", "SID='" + lcSID + "' AND SVAR='" + lcVar + "'")
        End If
    End Function

    Function GetSessionVariable(ByVal lcSID As String, lcVar As String, Optional Validate_Expiry As Boolean = False) As String
        'lcSID = GetCurrentUser()
        lcSID = GetWindowsUser()
        'Dim a = Session.CookieMode
        Dim dr As Data.DataRow = getDataRow(EBDB_CS, "Select * From EB_SessionsVariables where SID='" + lcSID + "' AND SVAR='" + lcVar + "'")
        If IsNothing(dr) Then
            GetSessionVariable = ""
        Else
            GetSessionVariable = FixNull(dr("SVAL"))
            If Validate_Expiry Then
                If FixNull(dr("VALID_UNTIL")) = "" Then
                    GetSessionVariable = ""
                Else
                    If Format(dr("VALID_UNTIL"), "yyyyMMdd") < Format(Now, "yyyyMMdd") Then
                        GetSessionVariable = ""
                    End If
                End If
            End If
        End If
    End Function

    'Sub FillCombo_DT(ByVal dt As Data.DataTable, ByRef CB As DropDownList, Optional ByVal lcDefault As String = "", Optional ByVal AddItemAll As Boolean = False, Optional ByVal DisplayOnly As Boolean = False, Optional ByVal AddOthers As Boolean = False)
    '    Dim i As Integer
    '    Dim lcItemText As String = ""
    '    Dim lcItemValue As String = ""
    '    CB.Items.Clear()
    '    For Each rw As DataRow In dt.Rows
    '        lcItemValue = CStr(rw(0))
    '        lcItemText = CStr(lcItemValue)
    '        If rw.ItemArray.Count > 1 Then lcItemText = CStr(FixNull(rw(1)))
    '        Dim ii As New ListItem
    '        ii.Text = lcItemText
    '        ii.Value = lcItemValue
    '        If DisplayOnly Then
    '            If lcDefault = "" Then
    '                CB.Items.Add(ii)
    '                CB.Items(0).Selected = True
    '                Exit For
    '            End If
    '            If UCase(lcDefault) = UCase(lcItemValue) Then
    '                CB.Items.Add(ii)
    '                CB.Items(0).Selected = True
    '            End If
    '        Else
    '            CB.Items.Add(ii)
    '            If UCase(lcDefault) = UCase(lcItemValue) Then CB.Items(i).Selected = True
    '        End If
    '    Next
    '    If AddItemAll Then
    '        CB.Items.Insert(0, "ALL")
    '        CB.Items(0).Value = "ALL"
    '    End If
    '    If AddOthers Then
    '        CB.Items.Insert(0, "Others")
    '        CB.Items(0).Value = "Others"
    '    End If
    '    If lcDefault = "" Then
    '        If CB.Items.Count <> -1 Then
    '            Try
    '                CB.Items(0).Selected = True
    '            Catch
    '                Dim aaa = 0
    '            End Try
    '        End If
    '    End If
    'End Sub

    Sub FillCombo(ByVal a As String(), ByRef CB As DropDownList, Optional ByVal lcDefault As String = "", Optional ByVal AddItemAll As Boolean = False, Optional ByVal DisplayOnly As Boolean = False, Optional ByVal AddOthers As Boolean = False, Optional ByVal AddNone As Boolean = False)
        Dim i As Integer
        Dim lcItemText As String
        Dim lcItemValue As String

        If AddItemAll Then
            a = Split("ALL^ALL" + vbCrLf + Join(a, vbCrLf), vbCrLf)
            'CB.Items.Insert(0, "ALL")
            'CB.Items(0).Value = "ALL"
        End If
        If AddOthers Then
            a = Split("Others^Others" + vbCrLf + Join(a, vbCrLf), vbCrLf)
            'CB.Items.Insert(0, "Others")
            'CB.Items(0).Value = "Others"
        End If
        If AddNone Then
            a = Split("None^None" + vbCrLf + Join(a, vbCrLf), vbCrLf)
            'CB.Items.Insert(0, "None")
            'CB.Items(0).Value = "None"
        End If

        CB.Items.Clear()
        For i = LBound(a) To UBound(a)
            If a(i) <> "" Then
                a(i) = Replace(a(i), "^", vbTab)
                lcItemValue = Split(a(i), vbTab)(0)
                Try
                    lcItemText = Split(a(i), vbTab)(1)
                Catch
                    lcItemText = lcItemValue
                End Try
                If DisplayOnly Then
                    If lcDefault = "" Then
                        CB.Items.Add(lcItemText)
                        CB.Items(0).Value = lcItemValue
                        CB.Items(0).Selected = True
                        Exit For
                    End If
                    If UCase(lcDefault) = UCase(lcItemValue) Then
                        CB.Items.Add(lcItemText)
                        CB.Items(0).Value = lcItemValue
                        CB.Items(0).Selected = True
                    End If
                Else
                    CB.Items.Add(lcItemText)
                    CB.Items(i).Value = lcItemValue
                    If UCase(lcDefault) = UCase(lcItemValue) Then CB.Items(i).Selected = True
                End If
            End If
        Next
        If lcDefault = "" Then
            If CB.Items.Count <> -1 Then
                Try
                    CB.Items(0).Selected = True
                Catch
                End Try
            End If
        End If
    End Sub

    'Sub Fill_ListBox_DT(ByVal dt As Data.DataTable, ByRef CB As System.Web.UI.WebControls.ListBox, Optional ByVal lcDefault As String = "")
    '    Dim i As Integer
    '    Dim lcItemText As String = ""
    '    Dim lcItemValue As String = ""
    '    CB.Items.Clear()
    '    For Each rw As DataRow In dt.Rows
    '        lcItemValue = CStr(rw(0))
    '        lcItemText = CStr(lcItemValue)
    '        If rw.ItemArray.Count > 1 Then lcItemText = CStr(FixNull(rw(1)))
    '        Dim ii As New ListItem
    '        ii.Text = lcItemText
    '        ii.Value = lcItemValue
    '        CB.Items.Add(ii)
    '        If UCase(lcDefault) = UCase(lcItemValue) Then CB.Items(i).Selected = True
    '    Next
    '    If lcDefault = "" Then
    '        If CB.Items.Count <> -1 Then
    '            Try
    '                CB.Items(0).Selected = True
    '            Catch
    '                Dim aaa = 0
    '            End Try
    '        End If
    '    End If
    'End Sub
    'Public Function GetMntDialog(ByVal lcUrl As String, Optional ByVal lnHeight As Integer = 600, Optional ByVal lnWidth As Integer = 800) As String
    'Dim lcScript As String = ""
    '    GetMntDialog = lcScript
    '    Dim lcHeight As String = CStr(lnHeight) + "px"
    '    Dim lcWidth As String = CStr(lnWidth) + "px"
    '    lcScript = ""
    '    lcScript = lcScript + "window.showModalDialog('"
    '    lcScript = lcScript + GetAppPath() + "/ShowIFFixed.aspx?QS=" + lcUrl + "','newwin','dialogHeight:" + lcHeight + ";dialogLeft:100px;dialogTop:50px;dialogWidth:" + lcWidth + ";center:yes;dialogHide:yes;edge:raised;help:no;resizable:yes;scroll:no;status:yes;unadorned:yes' );"
    '    GetMntDialog = lcScript
    'End Function

    Function ReadTextFile(ByVal lcFileName As String) As String
        'Open a file for reading
        Dim FILENAME As String = lcFileName
        'Get a StreamReader class that can be used to read the file
        Dim objStreamReader As StreamReader
        objStreamReader = File.OpenText(FILENAME)
        'Now, read the entire file into a string
        Dim contents As String = objStreamReader.ReadToEnd()
        ReadTextFile = contents
        objStreamReader.Close()
    End Function

    Function ReadTextFile1(ByVal lcFileName As String) As String
        'Open a file for reading
        Dim FILENAME As String = Server.MapPath(lcFileName)
        'Get a StreamReader class that can be used to read the file
        Dim objStreamReader As StreamReader
        objStreamReader = File.OpenText(FILENAME)
        'Now, read the entire file into a string
        Dim contents As String = objStreamReader.ReadToEnd()
        ReadTextFile1 = contents
        objStreamReader.Close()
    End Function

    Public Sub ExportGv2Excel(ByVal gv As GridView, ByVal lcFileName As String, ByVal p As System.Web.UI.Page)
        If gv.Rows.Count.ToString + 1 < 65536 Then
            gv.AllowPaging = "False"
            gv.DataBind()
            Dim sw As New StringWriter()
            Dim hw As New System.Web.UI.HtmlTextWriter(sw)
            Dim frm As HtmlForm = New HtmlForm()
            p.Response.ContentType = "application/vnd.ms-excel"
            p.Response.AddHeader("content-disposition", "attachment;filename=" & lcFileName & ".xls")
            p.Response.Charset = ""
            p.EnableViewState = False
            p.Controls.Add(frm)
            frm.Controls.Add(gv)
            frm.RenderControl(hw)
            p.Response.Write(sw.ToString())
            p.Response.End()
            gv.AllowPaging = "True"
            gv.DataBind()
        End If
    End Sub


    Function GetServerPath() As String
        '        GetServerPath = Request.Url.Host
        'GetServerPath = Request.ServerVariables("LOCAL_ADDR")
        GetServerPath = System.Net.Dns.GetHostAddresses(Request.Url.Host)(0).ToString()
    End Function

    Function Get_server_Url(ByRef p As System.Web.UI.Page) As String
        Dim laServerUrl As String() = Split(p.Request.Url.ToString, "/")
        Dim lcServerUrl As String = laServerUrl(0) + "//" + laServerUrl(2) + GetAppPath()
        Get_server_Url = lcServerUrl
    End Function

    'Public Sub ExportToExcel(ByVal dtTemp As Data.DataTable, ByVal filepath As String)
    '    Dim strFileName As String = filepath
    '    If System.IO.File.Exists(strFileName) Then
    '        'If (MessageBox.Show("Do you want to replace from the existing file?", "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = System.Windows.Forms.DialogResult.Yes) Then
    '        '    System.IO.File.Delete(strFileName)
    '        'Else
    '        '    Return
    '        'End If
    '    End If
    '    Dim _excel As New Application
    '    Dim wBook As Excel.Workbook
    '    Dim wSheet As Excel.Worksheet

    '    wBook = _excel.Workbooks.Add()
    '    wSheet = wBook.ActiveSheet()

    '    Dim dt As System.Data.DataTable = dtTemp
    '    Dim dc As System.Data.DataColumn
    '    Dim dr As System.Data.DataRow
    '    Dim colIndex As Integer = 0
    '    Dim rowIndex As Integer = 0
    '    For Each dc In dt.Columns
    '        colIndex = colIndex + 1
    '        wSheet.Cells(1, colIndex) = dc.ColumnName
    '    Next
    '    For Each dr In dt.Rows
    '        rowIndex = rowIndex + 1
    '        colIndex = 0
    '        For Each dc In dt.Columns
    '            colIndex = colIndex + 1
    '            wSheet.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
    '        Next
    '    Next
    '    wSheet.Columns.AutoFit()
    '    wBook.SaveAs(strFileName)

    '    ReleaseObject(wSheet)
    '    wBook.Close(False)
    '    ReleaseObject(wBook)
    '    _excel.Quit()
    '    ReleaseObject(_excel)
    '    GC.Collect()

    '    'MessageBox.Show("File Export Successfully!")
    'End Sub
    Private Sub ReleaseObject(ByVal o As Object)
        Try
            While (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0)
            End While
        Catch
        Finally
            o = Nothing
        End Try
    End Sub

    Function GetWindowsUser() As String
        Try
            Dim extendedUserName As String = Web.HttpContext.Current.User.Identity.Name.ToString
            Dim lcUser = Split(extendedUserName, "\")(UBound(Split(extendedUserName, "\")))
            GetWindowsUser = lcUser
        Catch ex As Exception
            GetWindowsUser = ""
        End Try
    End Function

    Function GetCurrentUser() As String
        Dim lcUser As String = GetSessionVariable("", "User_ID", True)
        GetCurrentUser = lcUser
        'GetCurrentUser = "1091"
        'GetCurrentUser = "2326"
        'Try
        '    Dim extendedUserName As String = Web.HttpContext.Current.User.Identity.Name.ToString
        '    'Thread.CurrentPrincipal()
        '    Dim lcUser = Split(extendedUserName, "\")(UBound(Split(extendedUserName, "\")))
        '    GetCurrentUser = lcUser
        '    'GetCurrentUser = 2322
        'Catch ex As Exception
        '    GetCurrentUser = ""
        'End Try
        'GetCurrentUser = "2185"
    End Function

    Sub FillCheckBoxes(ByVal a As Array, ByRef CB As CheckBoxList, Optional ByVal lcSelected As String = "")
        Dim i As Integer
        Dim lcItemText As String
        Dim lcItemValue As String
        CB.Items.Clear()
        For i = LBound(a) To UBound(a)
            If a(i) <> "" Then
                lcItemValue = Split(a(i), vbTab)(0)
                lcItemText = Split(a(i), vbTab)(1)
                CB.Items.Add(lcItemText)
                CB.Items(i).Value = lcItemValue
                CB.Items(i).Selected = False
                If InStr(ReplaceAll("~" + lcSelected + "~", "~~", "~"), "~" + lcItemValue + "~") > 0 Then
                    CB.Items(i).Selected = True
                End If
            End If
        Next
    End Sub

    Function GetSqlArray(ByVal lcConnectionString As String, ByVal lcSql As String) As String()
        Dim dt As Data.DataTable = GetDataTable(lcConnectionString, lcSql)
        Dim lcLines As String = ""
        For Each rs In dt.Rows
            Dim lcCode As String = FixNull(rs(0))
            If dt.Columns.Count > 1 Then
                Dim lcPrompt As String = FixNull(rs(1))
                If lcPrompt = "" Then
                    lcLines = lcLines + lcCode + vbTab + lcCode + vbCrLf
                Else
                    lcLines = lcLines + lcCode + vbTab + lcPrompt + vbCrLf
                End If
            Else
                lcLines = lcLines + lcCode + vbTab + lcCode + vbCrLf
            End If
        Next
        Return Split(lcLines, vbCrLf)
    End Function

    Function ReplaceAll(ByVal lcText As String, ByVal S1 As String, ByVal S2 As String) As String
        Dim lcText1 = lcText
        Do While InStr(lcText1, S1) > 0
            lcText1 = Replace(lcText1, S1, S2)
        Loop
        ReplaceAll = lcText1
    End Function

    Sub GetDatesRange(ByVal lcDateRange As String, ByRef Date1 As String, ByRef Date2 As String)
        Select Case lcDateRange
            Case "T"
                Date1 = Format(Now.Date, "yyyy-MM-dd")
                Date2 = Format(Now.Date, "yyyy-MM-dd")
            Case "W", "TW"
                Dim dow As Integer = ((Now.Date.DayOfWeek() + 2 - 1) Mod 7) + 1
                Date1 = IIf(dow = 1, Format(Now.Date, "yyyy-MM-dd"), Format(Now.Date.AddDays(-1 * (dow - 1)), "yyyy-MM-dd"))
                Date2 = Format(Now.Date, "yyyy-MM-dd")
            Case "M", "TM"
                Date1 = Format(Now.Date, "yyyy-MM-01")
                Date2 = Format(Now.Date, "yyyy-MM-31")
            Case "LM"
                If Now.Month = 1 Then
                    Dim LastYear As Integer = Now.Year - 1
                    Date1 = CStr(LastYear) + "-12-01"
                    Date2 = CStr(LastYear) + "-12-31"
                Else
                    Dim LastMonth As Integer = Now.Month - 1
                    Date1 = CStr(Now.Year) + "-" + Right("00" + CStr(LastMonth), 2) + "-01"
                    Date2 = CStr(Now.Year) + "-" + Right("00" + CStr(LastMonth), 2) + "-31"
                End If
            Case "Y", "TY"
                Date1 = Format(Now.Date, "yyyy-01-01")
                Date2 = Format(Now.Date, "yyyy-12-31")
            Case "LY"
                Dim LastYear As Integer = Now.Year - 1
                Date1 = CStr(LastYear) + "-01-01"
                Date2 = CStr(LastYear) + "-12-31"
        End Select
    End Sub

    'Function API_V01()
    '    Dim V_Data_Text As String = ""
    '    V_Data_Text = V_Data_Text + "A1_BRANCH=1~"                    '   aa
    '    V_Data_Text = V_Data_Text + "A1_ACCOUNT=0000110693048013~"    '   ab
    '    V_Data_Text = V_Data_Text + "A1_CURRENCY=48~"                 '   ad
    '    ' V_Data_Text = V_Data_Text + "A1_NAME=AQEEL MAYOOF~"           '   ac
    '    V_Data_Text = V_Data_Text + "M1_AMOUNT=123.000~"              '   ma
    '    V_Data_Text = V_Data_Text + "T1_TEXT=TEST01~"                 '   ni
    '    V_Data_Text = V_Data_Text + "TODAY_DATE=30-12-2015~"          '   a1
    '    V_Data_Text = V_Data_Text + "LOCAL_BRANCH=1~"                 '   a3
    '    V_Data_Text = V_Data_Text + "D1_VAL_DATE=30-12-2015~"         '   mq
    '    V_Data_Text = V_Data_Text + "F1_COMMISSION=0.0~"              '   pa
    '    V_Data_Text = V_Data_Text + "LOCAL_BR_DESC=TEST~"             '   a4
    '    V_Data_Text = V_Data_Text + "A1_ADDRESS_1=TEST~"              '   am
    '    V_Data_Text = V_Data_Text + "A1_ADDRESS_2=TEST~"              '   aj
    '    V_Data_Text = V_Data_Text + "A1_ADDRESS_3=TEST~"              '   ak
    '    'V_Data_Text = V_Data_Text + "TRSH_BLK.TRSH_NUM=0~"             '   a8
    '    V_Data_Text = V_Data_Text + "A1_CLIENT_NAME=TEST~"             '   ah
    '    ''
    '    Dim oc As New OleDb.OleDbConnection(ICBS_CS)
    '    oc.Open()
    '    Dim myCMD As OleDbCommand = New OleDbCommand("Bs_Genopr", oc)
    '    myCMD.CommandType = CommandType.StoredProcedure
    '    myCMD.Parameters.Add(New OleDbParameter("P_MODCODE", OleDbType.VarChar)).Value = "GEN"
    '    myCMD.Parameters.Add(New OleDbParameter("P_SMODCODE", OleDbType.VarChar)).Value = "DEP"
    '    myCMD.Parameters.Add(New OleDbParameter("P_OPRERATION", OleDbType.VarChar)).Value = "V01"
    '    myCMD.Parameters.Add(New OleDbParameter("P_TODAYDATE", OleDbType.VarChar)).Value = "30-12-2015"
    '    myCMD.Parameters.Add(New OleDbParameter("P_BRANCH", OleDbType.Numeric)).Value = 1
    '    myCMD.Parameters.Add(New OleDbParameter("P_NEEDLINE", OleDbType.Numeric)).Value = 1
    '    myCMD.Parameters.Add(New OleDbParameter("P_NEEDADVICE", OleDbType.Numeric)).Value = 1
    '    myCMD.Parameters.Add(New OleDbParameter("P_DATATEXT", OleDbType.VarChar)).Value = V_Data_Text
    '    myCMD.Parameters.Add(New OleDbParameter("P_USERID", OleDbType.VarChar)).Value = "ICBS"
    '    myCMD.Parameters.Add(New OleDbParameter("P_PROGRAMID", OleDbType.VarChar)).Value = "ICBS"
    '    myCMD.Parameters.Add(New OleDbParameter("P_NODEID", OleDbType.VarChar)).Value = "ICBS"
    '    myCMD.Parameters.Add(New OleDbParameter("P_TIMESTAMP", OleDbType.VarChar)).Value = "30-12-2015"
    '    myCMD.Parameters.Add(New OleDbParameter("P_TRSNBR", OleDbType.Numeric)).Value = 0
    '    myCMD.Parameters.Add(New OleDbParameter("P_ERROR", OleDbType.VarChar, 200)).Value = ""
    '    myCMD.Parameters("P_TRSNBR").Direction = ParameterDirection.InputOutput
    '    myCMD.Parameters("P_ERROR").Direction = ParameterDirection.InputOutput
    '    Try
    '        myCMD.ExecuteScalar()
    '        Dim lError As String = myCMD.Parameters("P_ERROR").Value
    '        If lError <> "" Then
    '            API_V01 = "Error:" + CStr(myCMD.Parameters("P_ERROR").Value)
    '        Else
    '            API_V01 = myCMD.Parameters("P_TRSNBR").Value
    '        End If
    '        'MsgBox(myCMD.Parameters("P_TRSNBR").Value)
    '    Catch ex As Exception
    '        Dim lError As String = myCMD.Parameters("P_ERROR").Value
    '        If lError <> "" Then
    '            API_V01 = "Error:" + CStr(myCMD.Parameters("P_ERROR").Value)
    '        Else
    '            API_V01 = ex.Message
    '        End If
    '        'MsgBox(myCMD.Parameters("P_ERROR").Value)
    '        'MsgBox(ex.Message)
    '    End Try
    '    oc.Close()
    'End Function

    Function DB2_CS() As String
        Dim lcServerName As String = Get_Server_Name()
        If InStr(UCase(Left(lcServerName, 20)), "DRWEB") > 0 Then
            DB2_CS = "Provider=IBMDA400.1;Data Source=192.168.3.3;User ID=inap;Password=inapinap"
        Else
            DB2_CS = "Provider=IBMDA400.1;Data Source=192.168.1.3;User ID=inap;Password=inapinap"
        End If
    End Function

    'Function GetDB2ConnectionString() As String
    '    GetDB2ConnectionString = "Provider=IBMDA400.1;Data Source=192.168.1.3;User ID=inap;Password=inapinap"
    'End Function
    'Function EBDB_CN() As String
    '    EBDB_CN = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=S0320)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=EBDB )));User Id=inap ;Password=inap"
    'End Function

    Function EBDB_CS(Optional lcDatabase As String = "EBDB", Optional lcUserID As String = "inap", Optional lcUserPW As String = "inap") As String
        'EBDB_CS = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=S0230)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=" + lcDatabase + ")));User Id=" + lcUserID + ";Password=" + lcUserPW + ";"
        'Dim lcServer_ip As String = GetServerPath()
        'Dim lcServer_ip As String = Get_server_Url()
        EBDB_CS = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=S0344)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=" + lcDatabase + ")));User Id=" + lcUserID + ";Password=" + lcUserPW + ";"
        Dim lcServerName As String = Get_Server_Name()
        If InStr(UCase(Left(lcServerName, 20)), "DRWEB") > 0 Then
            EBDB_CS = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=S0344)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=" + lcDatabase + ")));User Id=" + lcUserID + ";Password=" + lcUserPW + ";"
        End If
        If InStr(UCase(Left(lcServerName, 20)), "S0305") > 0 Then
            EBDB_CS = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=S0344)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=" + lcDatabase + ")));User Id=" + lcUserID + ";Password=" + lcUserPW + ";"
        End If
        If InStr(UCase(Left(lcServerName, 20)), "S0307") > 0 Then
            EBDB_CS = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=S0320)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=" + lcDatabase + ")));User Id=" + lcUserID + ";Password=" + lcUserPW + ";"
        End If
        If InStr(UCase(Left(lcServerName, 20)), "LOCALHOST") > 0 Then
            EBDB_CS = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=S0320)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=" + lcDatabase + ")));User Id=" + lcUserID + ";Password=" + lcUserPW + ";"
            'EBDB_CS = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=S0344)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=" + lcDatabase + ")));User Id=" + lcUserID + ";Password=" + lcUserPW + ";"
        End If
        'EBDB_CS = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=S0320)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=" + lcDatabase + ")));User Id=" + lcUserID + ";Password=" + lcUserPW + ";"
    End Function

    Function EBDB_CS_Prod(Optional lcDatabase As String = "EBDB", Optional lcUserID As String = "inap", Optional lcUserPW As String = "inap") As String
        EBDB_CS_Prod = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=S0344)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=" + lcDatabase + ")));User Id=" + lcUserID + ";Password=" + lcUserPW + ";"
    End Function

    Function ICBS_CS_Prod() As String
        'Dim lcServerName As String = Get_Server_Name()
        'ICBS_CS_Prod = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=scan-eskan)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBSSRV)));User Id=icbs;Password=icbs2019;"
        'If InStr(UCase(Left(lcServerName, 20)), "DRWEB") > 0 Then
        '    ICBS_CS_Prod = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=drdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBSDR)));User Id=ICBS;Password=icbs2019;"
        'End If
        ICBS_CS_Prod = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=testdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBSTST1)));User Id=ICBS;Password=icbs2020;"
    End Function

    Function EBDB_CS_Test(Optional lcDatabase As String = "EBDB", Optional lcUserID As String = "inap", Optional lcUserPW As String = "inap") As String
        EBDB_CS_test = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=S0320)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=" + lcDatabase + ")));User Id=" + lcUserID + ";Password=" + lcUserPW + ";"
    End Function

    Function ICBS_CS_Test() As String
        'ICBS_CS_Test = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=testdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBS)));User Id=ICBS;Password=icbs2020;"
        'ICBS_CS_Test = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=testdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBS1)));User Id=ICBS;Password=icbs2020;"
        ICBS_CS_Test = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=testdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBS)));User Id=ICBS;Password=icbs2020;"
        'ICBS_CS_Test = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=testdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBS16)));User Id=ICBS;Password=icbs2020;"
        'ICBS_CS_Test = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=testdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBS1)));User Id=ICBS;Password=icbs2020;"
    End Function

    Function ICBS_CS() As String
        'GetOracleConnectionString = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=localhost)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBSTRN1)));User Id=ICBS;Password=ICBS;"
        '        GetOracleConnectionString = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=localhost)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBSTRN1)));User Id=ICBS;Password=ICBS;"
        'ICBS_CS = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=s0320.EskanBank.com)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBSTRN3)));User Id=ICBS;Password=sbci;"
        'ICBS_CS = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=testdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBSTEST)));User Id=ICBS;Password=icbs;"
        ICBS_CS = ICBS_CS1()
        'ICBS_CS = ICBS_CS_Test()
    End Function

    'Function ICBS_CS1(Optional lcDatabase As String = "") As String
    '    Select Case UCase(lcDatabase)
    '        Case "UAT"
    '            ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=s0320.EskanBank.com)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBSTRN3)));User Id=ICBS;Password=sbci;"
    '        Case "ICBS16"
    '            ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=testdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBS)));User Id=ICBS;Password=icbs01;"
    '        Case "TEST"
    '            ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=testdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBS)));User Id=ICBS;Password=icbs01;"
    '        Case "PRODUCTION"
    '            ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.20.12)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBS)));User Id=INAP;Password=inap;"
    '        Case "DR"
    '            ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.20.12)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBS)));User Id=INAP;Password=inap;"
    '        Case Else
    '            'ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=testdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBS)));User Id=INAP;Password=inap;"
    '            ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=testdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBS)));User Id=ICBS;Password=icbs01;"
    '    End Select
    'End Function

    Function ICBS_CS1(Optional lcDatabase As String = "") As String
        'Select Case UCase(lcDatabase)
        '    Case "UAT"
        '        ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=s0320.EskanBank.com)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBSTRN3)));User Id=ICBS;Password=sbci;"
        '    Case "ICBS16", "TEST"
        '        'ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=testdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBS1)));User Id=ICBS;Password=icbs2020;"
        '        ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=testdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBS)));User Id=ICBS;Password=icbs2020;"
        '    Case "ICBSTST1"
        '        'ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=testdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBS1)));User Id=ICBS;Password=icbs2020;"
        '        ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=testdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBSTST1)));User Id=ICBS;Password=icbs2020;"
        '    Case "PRODUCTION"
        '        'ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=DB2)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBS)));User Id=INAP;Password=inap;"
        '        ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=scan-eskan)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBSSRV)));User Id=ICBS;Password=icbs2019;"
        '    Case "DR"
        '        ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=drdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBSDR)));User Id=INAP;Password=inap;"
        '    Case Else
        '        'ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=testdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBS)));User Id=INAP;Password=inap;"
        '        'ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=scan-eskan)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBSSRV)));User Id=ICBS;Password=icbs2019;"
        '        Dim lcServerName As String = Get_Server_Name()
        '        If InStr(UCase(Left(lcServerName, 20)), "DRWEB") > 0 Then
        '            ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=drdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBSDR)));User Id=ICBS;Password=icbs2019;"
        '        End If
        '        If InStr(UCase(Left(lcServerName, 20)), "S0307") > 0 Then
        '            ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=testdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBS)));User Id=ICBS;Password=icbs2020;"
        '        End If
        '        If InStr(UCase(Left(lcServerName, 20)), "LOCALHOST") > 0 Then
        '            ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=testdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBS)));User Id=ICBS;Password=icbs2020;"
        '            ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=testdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBSTST1)));User Id=ICBS;Password=icbs2020;"
        '            'ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=scan-eskan)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBSSRV)));User Id=ICBS;Password=icbs2019;"               ' to be removed
        '        End If
        'End Select
        ICBS_CS1 = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=testdb)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBSTST1)));User Id=ICBS;Password=icbs2020;"
    End Function

    Function ICBS_CS_Instant(Optional lcInstant As String = "") As String
        Select Case UCase(lcInstant)
            Case "1"
                ICBS_CS_Instant = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=DB1)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBS)));User Id=INAP;Password=inap;"
            Case "2"
                ICBS_CS_Instant = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=DB2)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBS)));User Id=INAP;Password=inap;"
        End Select
    End Function


    Public Function GetDataTable(ByVal lcConnection As String, lcSql As String) As Data.DataTable
        Dim dt As Data.DataTable = New Data.DataTable
        Dim oda As OleDbDataAdapter = New OleDbDataAdapter(lcSql, lcConnection)
        oda.Fill(dt)
        GetDataTable = dt
    End Function

    'Public Function GetDataTable_oc(ByVal lcConnection As String, lcSql As String) As Data.DataTable
    '    Dim dt As Data.DataTable = New Data.DataTable
    '    Dim oda As OleDbDataAdapter = New OleDbDataAdapter(lcSql, lcConnection)
    '    oda.Fill(dt)
    '    GetDataTable = dt
    'End Function

    Public Function getDataColumn(lcConnectionString As String, lcSql As String, Optional lcColumn As String = "") As String
        getDataColumn = ""
        Dim dt As Data.DataTable = GetDataTable(lcConnectionString, lcSql)
        If dt.Rows.Count > 0 Then
            If lcColumn = "" Then
                getDataColumn = CStr(FixNull(dt.Rows(0)(0)))
            Else
                getDataColumn = CStr(FixNull(dt.Rows(0)(lcColumn)))
            End If
        End If
    End Function

    Function Base64ToImage(ByVal base64string As String) As System.Drawing.Image
        'Setup image and get data stream together3.        
        Dim img As System.Drawing.Image
        Dim MS As System.IO.MemoryStream = New System.IO.MemoryStream
        Dim b64 As String = base64string.Replace(" ", "+")
        'Dim b64 As String = base64string
        Dim b() As Byte
        'Converts the base64 encoded msg to image data
        b = Convert.FromBase64String(b64)
        MS = New System.IO.MemoryStream(b)
        'creates image
        img = System.Drawing.Image.FromStream(MS)
        Return img
    End Function

    Function Base64_To_Text(ByVal base64string As String) As String
        Base64_To_Text = ""
        If base64string <> "" Then
            Dim data() As Byte
            data = System.Convert.FromBase64String(base64string)
            Base64_To_Text = System.Text.UnicodeEncoding.Unicode.GetString(data)
        End If
    End Function

    Function Base64_From_Text(ByVal lcText As String) As String
        Base64_From_Text = ""
        If lcText <> "" Then
            Base64_From_Text = Convert.ToBase64String(System.Text.UnicodeEncoding.Unicode.GetBytes(lcText))
        End If
    End Function

    Function GetAppPath() As String
        Dim lcPath As String = HttpRuntime.AppDomainAppVirtualPath
        GetAppPath = lcPath
        If lcPath = "\" Or lcPath = "/" Then
            GetAppPath = ""
        End If
    End Function

    Function GetCustomerImage(lcCPR As String) As String
        Dim dt As Data.DataTable = getDB2DataTable("SELECT IDPHOTO FROM EBLIB.EIDINF0PF WHERE IDCPR='" + lcCPR + "'")
        If dt.Rows.Count > 0 Then
            Dim rs As DataRow = dt.Rows(0)
            Dim img As System.Drawing.Image = Base64ToImage(Trim(rs("IDPHOTO")))
            Dim lcFileName = Server.MapPath(GetAppPath() + "\TempImages\" + lcCPR + "_Photo.Jpeg")
            Try
                img.Save(lcFileName, System.Drawing.Imaging.ImageFormat.Jpeg)
            Catch ex As Exception

            End Try
            GetCustomerImage = GetAppPath() + "\TempImages\" + lcCPR + "_Photo.Jpeg"
            'GetCustomerImage = "TempImages/" + lcCPR + "_Photo.Jpeg"
        End If
    End Function

    Function GetIBAN(ByVal lcAccount As String) As String
        'Dim db As New ADODB.Connection
        'Dim rs As New ADODB.Recordset
        'Dim lcSql As String = "Select NEIBAN from NEPF where NEAB||NEAN||NEAS='@ACC' "
        'lcSql = Replace(lcSql, "@ACC", Replace(lcAccount, "-", ""))
        'Try
        '    Dim lcIBan = GetSqlColumn("KFILLIV", lcSql, 0)
        '    GetIBAN = Mid(lcIBan, 1, 4) + " " + _
        '                Mid(lcIBan, 5, 4) + " " + _
        '                Mid(lcIBan, 9, 4) + " " + _
        '                Mid(lcIBan, 13, 4) + " " + _
        '                Mid(lcIBan, 17, 4) + " " + _
        '                Mid(lcIBan, 21, 2)
        'Catch ex As Exception
        '    GetIBAN = ""
        'End Try
    End Function

    Function GetCustomerName(ByVal lcCustomer As String) As String
        'Dim db As New ADODB.Connection
        'Dim rs As New ADODB.Recordset
        'Dim lcSql As String = "Select GFCUN from GFPF where GFCUS='@CUS' "
        'lcSql = Replace(lcSql, "@CUS", lcCustomer)
        'GetCustomerName = GetSqlColumn("KFILLIV", lcSql, 0)
    End Function

    Function Lookup(ByVal cl As Collection, ByVal lcKey As String) As String
        Lookup = ""
        If cl.Contains(lcKey) Then Lookup = cl(UCase(lcKey))
    End Function

    Function FormatFromEqDate(ByVal lcEqDate As String) As String
        lcEqDate = Right("0000000" + lcEqDate, 7)
        FormatFromEqDate = Mid(lcEqDate, 6, 2) + "-" + Mid(lcEqDate, 4, 2) + "-" + CStr(1900 + Val(Mid(lcEqDate, 1, 3)))
    End Function

    Public Function FixNullDate(ByVal dbvalue As Object, lcFormat As String) As String
        If dbvalue Is DBNull.Value Then
            Return ""
        Else
            'NOTE: This will cast value to string if it isn't a string.
            Return Format(dbvalue, lcFormat)
        End If
    End Function
    Public Function FixNull(ByVal dbvalue) As String
        If dbvalue Is DBNull.Value Then
            Return ""
        Else
            'NOTE: This will cast value to string if it isn't a string.
            Return dbvalue.ToString
        End If
    End Function

    'Sub DT2COLLECTION(ByVal DT As Data.DataTable, ByRef cl As Collection, Optional ByVal IgnoreDuplicates As Boolean = False)
    '    Dim row As Data.DataRow
    '    Dim lcKey As String
    '    cl.Clear()
    '    For Each row In DT.Rows
    '        lcKey = row.Item(0)
    '        If IgnoreDuplicates Then
    '            Try
    '                cl.Add(FixNull(row.Item(1)), UCase(lcKey))
    '            Catch ex As Exception

    '            End Try
    '        Else
    '            cl.Add(FixNull(row.Item(1)), UCase(lcKey))
    '        End If
    '    Next
    'End Sub

    Sub DT_COLLECTION(ByRef DT As Data.DataTable, _
                  ByRef cl As Collection, _
                  ByVal lcPrefix As String, _
                  ByVal laKeys As String())
        Dim row As Data.DataRow
        'cl.Clear()
        For Each row In DT.Rows
            Dim lcKey As String = ""
            For ii = LBound(laKeys) To UBound(laKeys)
                lcKey = lcKey + IIf(lcKey <> "", "-", "") + FixNull(row.Item(laKeys(ii)))
            Next
            If lcPrefix <> "" Then lcKey = lcPrefix + "-" + lcKey + ""
            'Dim lcLine As String = lcKey + vbTab
            Dim lcLine As String = ""
            For ii = LBound(row.ItemArray) To UBound(row.ItemArray)
                lcLine = lcLine + CStr(FixNull(row.Item(ii))) + vbTab
            Next
            If cl.Contains(lcKey) Then
                Dim lcOldData As String = cl(lcKey)
                cl.Remove(lcKey)
                cl.Add(lcOldData + "^" + lcLine, lcKey)
            Else
                cl.Add(lcLine, UCase(lcKey))
            End If
        Next
    End Sub

    'Public Sub LoadGridView(ByVal gv As GridView, ByVal sSql As String)
    '    dt = New Data.DataTable
    '    oda = New OleDbDataAdapter(sSql, oc)
    '    oda.Fill(dt)
    '    gv.DataSource = dt
    '    gv.DataBind()
    '    oc.Close()
    'End Sub

    Public Sub LoadGridView(ByVal dsn As String, ByVal gv As GridView, ByVal sSql As String)
        sSql = Replace(sSql, "[", Chr(34))
        sSql = Replace(sSql, "]", Chr(34))
        gv.DataSource = getOleDataTable(sSql)
        gv.DataBind()
    End Sub

    'Public Sub LoadDV(ByVal dv As DetailsView, ByVal sSql As String)
    '    dt = New Data.DataTable
    '    oda = New OleDbDataAdapter(sSql, oc)
    '    oda.Fill(dt)
    '    dv.DataSource = dt
    '    dv.DataBind()
    '    oc.Close()
    'End Sub

    Public Sub LoadDV(ByVal lcConnectionString As String, ByVal dv As DetailsView, ByVal lcSql As String)
        'dt = New Data.DataTable
        'oda = New OleDbDataAdapter(sSql, oc)
        'oda.Fill(dt)
        'dv.DataSource = dt
        'dv.DataBind()
        'oc.Close()
        dv.DataSource = GetDataTable(lcConnectionString, lcSql)
        dv.DataBind()
    End Sub

    Public Sub LoadDDL(ByVal ddl As DropDownList, ByVal sSql As String)
        Dim oc As New System.Data.OleDb.OleDbConnection("Provider=IBMDA400.1;Data Source=192.168.1.3;User ID=inap;Password=inapinap")
        cmd = New OleDbCommand(sSql, oc)
        Try
            oc.Open()
            odr = cmd.ExecuteReader
            Do While odr.Read
                If Not ddl.Items.Contains(odr(0)) Then ddl.Items.Add(odr(0))
            Loop
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: " & Err.Number)
        Finally
            odr.Close()
            oc.Close()
        End Try
    End Sub

    'Public Function ExecuteScalar(ByVal sSql As String)
    '    ExecuteScalar = 0
    '    Dim oc As New OleDb.OleDbConnection(ICBS_CS)
    '    Try
    '        oc.Open()
    '        cmd = New OleDb.OleDbCommand(sSql, oc)
    '        ExecuteScalar = cmd.ExecuteScalar
    '    Catch ex As Exception
    '        Dim aaa = 0
    '        'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: " & Err.Number)
    '    Finally
    '        oc.Close()
    '        If IsDBNull(ExecuteScalar) Then ExecuteScalar = 0
    '    End Try
    'End Function

    Public Function getOleDataTable(sSql As String, Optional lcSystem As String = "") As Data.DataTable
        Dim oc As New System.Data.OleDb.OleDbConnection(ICBS_CS)
        Dim dt As Data.DataTable = New Data.DataTable
        Dim oda As OleDbDataAdapter = New OleDbDataAdapter(sSql, oc)
        oda.Fill(dt)
        getOleDataTable = dt
    End Function

    Public Function getDB2DataTable(sSql As String, Optional lcSystem As String = "") As Data.DataTable
        Dim oc As New System.Data.OleDb.OleDbConnection(DB2_CS)
        Dim dt As Data.DataTable = New Data.DataTable
        Dim oda As OleDbDataAdapter = New OleDbDataAdapter(sSql, oc)
        oda.Fill(dt)
        getDB2DataTable = dt
    End Function

    'Public Function ExecuteNonQuery(ByVal sSql As String)
    '    ExecuteNonQuery = 0
    '    cmd = New OleDbCommand(sSql, oc)
    '    Try
    '        oc.Open()
    '        ExecuteNonQuery = cmd.ExecuteNonQuery
    '    Catch ex As Exception
    '        'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: " & Err.Number)
    '        '
    '    Finally
    '        oc.Close()
    '        If IsDBNull(ExecuteNonQuery) Then ExecuteNonQuery = 0
    '    End Try
    'End Function

    '    Public Sub ShowMsgBox(ByVal message As String, ByRef P As Page)
    ' Dim CSM As ClientScriptManager = P.ClientScript
    '     CSM.RegisterStartupScript(Me.GetType, Guid.NewGuid().ToString(), "<script language=""JavaScript"">" & GetAlertScript(message) & "</script>")
    'End Sub


    'Public Function OpenDsn(ByRef con As ADODB.Connection, ByRef rs As ADODB.Recordset, ByVal lcDSN As String, ByVal lcSQL As String) As Boolean
    '    '    If InStr(lcDSN, "KFIL") > 0 Or InStr(lcDSN, "EBLIB") > 0 Then lcDSN = lcDSN + ";" + "User ID=INAP;Password=INAPINAP"
    '    '    con.Mode = ADODB.ConnectModeEnum.adModeRead
    '    '    con.ConnectionString = "DSN=" + lcDSN
    '    '    con.Open()
    '    '    If InStr(LCase(con.ConnectionString), ".mdb") > 0 Then
    '    '        lcSQL = Replace(lcSQL, "||", "+")
    '    '        lcSQL = Replace(UCase(lcSQL), "SUBSTR", "MID")
    '    '        lcSQL = Replace(UCase(lcSQL), "UPPER", "UCASE")
    '    '    End If
    '    '    rs.Open(lcSQL, con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '    '    OpenDsn = True
    '    '    If (rs.EOF() And rs.BOF()) Then OpenDsn = False
    'End Function

    'Public Function OpenDsnRw(ByRef con As ADODB.Connection, ByRef rs As ADODB.Recordset, ByVal lcDSN As String, ByVal lcSQL As String) As Boolean
    '    If InStr(lcDSN, "KFIL") > 0 Or InStr(lcDSN, "EBLIB") > 0 Then lcDSN = lcDSN + ";" + "User ID=INAP;Password=INAPINAP"
    '    con.Mode = ADODB.ConnectModeEnum.adModeReadWrite
    '    con.ConnectionString = "DSN=" + lcDSN
    '    con.Open()
    '    If InStr(LCase(con.ConnectionString), ".mdb") > 0 Then
    '        lcSQL = Replace(lcSQL, "||", "+")
    '        lcSQL = Replace(UCase(lcSQL), "SUBSTR", "MID")
    '        lcSQL = Replace(UCase(lcSQL), "UPPER", "UCASE")
    '    End If
    '    rs.Open(lcSQL, con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '    OpenDsnRw = True
    '    If (rs.EOF() And rs.BOF()) Then OpenDsnRw = False
    'End Function

    Function GetSqlColumn(ByVal lcDSN As String, ByVal lcSql As String, Optional ByVal lnCOL As Integer = 0) As String
        Dim dt As Data.DataTable = GetDataTable(lcDSN, lcSql)
        GetSqlColumn = ""
        If dt.Rows.Count > 0 Then
            GetSqlColumn = FixNull(dt.Rows(0)(dt.Columns(lnCOL)))
        End If
    End Function


    Function GetSqlDataTable(ByVal lcDSN As String, ByVal lcSql As String) As Data.DataTable
        GetSqlDataTable = GetDataTable(lcDSN, lcSql)
        '    Dim db As New ADODB.Connection
        '    Dim rs As New ADODB.Recordset
        '    Dim DT As New Data.DataTable()
        '    Dim Row1 As DataRow
        '    Dim I As Integer
        '    Dim lcFldName As String

        '    OpenDsn(db, rs, lcDSN, lcSql)

        '    ' Add all coulumns to the table
        '    For I = 0 To rs.Fields.Count - 1
        '        lcFldName = rs.Fields(I).Name
        '        DT.Columns.Add(New DataColumn(lcFldName, System.Type.GetType("System.String")))
        '    Next

        '    ' Add all resords to the data table
        '    Dim lcKeys As String = ""
        '    Do While Not rs.EOF()
        '        Row1 = DT.NewRow()
        '        For I = 0 To rs.Fields.Count - 1
        '            lcFldName = rs.Fields(I).Name
        '            Row1(lcFldName) = rs(lcFldName).Value
        '        Next
        '        DT.Rows.Add(Row1)
        '        rs.MoveNext()
        '    Loop
        '    rs.Close()
        '    db.Close()
        '    Return DT
    End Function



    Sub AssignValue(ByVal r As GridViewRow, ByVal lcLabel As String, ByVal lcValue As Object)
        Dim lbl As System.Web.UI.WebControls.Label
        lbl = r.FindControl(lcLabel)
        If Not IsNothing(lbl) Then
            lbl.Text = FixNull(lcValue)
        End If
    End Sub

    Sub TryAssignValue(ByVal r As GridViewRow, ByVal lcLabel As String, ByVal lcField As String, Optional lcFormat As String = "")
        Dim lbl As System.Web.UI.WebControls.Label
        lbl = r.FindControl(lcLabel)
        If Not IsNothing(lbl) Then
            If Not IsDBNull(r.DataItem(lcField)) Then
                If lcFormat = "" Then
                    lbl.Text = r.DataItem(lcField)
                Else
                    lbl.Text = Format(Val(CStr(r.DataItem(lcField))), lcFormat)
                End If
            Else
                lbl.Text = "?"
                lbl.BackColor = System.Drawing.Color.Red
            End If
        End If
    End Sub


End Class

Public Class ClientSidePage
    Inherits System.Web.UI.Page

    'Public Sub DisplayAJAXMessage(ByVal p As System.Web.UI.Page, ByVal msg As String)
    '    msg = Replace(msg, vbCrLf, " ")
    '    msg = Replace(msg, vbCr, " ")
    '    msg = Replace(msg, vbLf, " ")
    '    Dim myScript As String = "alert('" + msg + "');"
    '    'ScriptManager.RegisterStartupScript(Page, p.GetType(), "MyScript", myScript, True)
    '    System.Web.UI.ScriptManager.RegisterClientScriptBlock(p, p.GetType(), "Alert_Box", myScript, True)
    'End Sub

    Public Sub CloseAjaxWindow(ByVal p As System.Web.UI.Page)
        '        Dim myScript As String = "opener=self;window.close();"
        'ScriptManager.r()
        'ScriptManager.RegisterClientScriptBlock(Page, p.GetType(), "MyScript", myScript, True)
        'ScriptManager.RegisterStartupScript(Page, p.GetType(), "MyScript", myScript, True)
        '       System.Web.UI.ScriptManager.RegisterClientScriptBlock(p, p.GetType(), "Close_Window", "self.close();", True)

        Dim CSM As ClientScriptManager = p.ClientScript
        CSM.RegisterStartupScript(Me.GetType, Guid.NewGuid().ToString(), "<script language=""JavaScript"">top.self.close();</script>")

    End Sub

    Public Sub DisplayAlert(ByVal message As String, ByRef p As System.Web.UI.Page)
        Dim CSM As ClientScriptManager = p.ClientScript
        'CSM.RegisterClientScriptBlock(Me.GetType, Guid.NewGuid().ToString(), "<script language=""JavaScript"">" & GetAlertScript(message) & "</script>")

    End Sub

    Public Sub DisplayAlertAfterLoad(ByVal message As String, ByRef P As System.Web.UI.Page)
        Dim CSM As ClientScriptManager = P.ClientScript
        CSM.RegisterStartupScript(Me.GetType, Guid.NewGuid().ToString(), "<script language=""JavaScript"">" & GetAlertScript(message) & "</script>")
    End Sub

    Public Sub CloseThisWindows(ByRef P As System.Web.UI.Page)
        Dim CSM As ClientScriptManager = P.ClientScript
        CSM.RegisterStartupScript(Me.GetType, Guid.NewGuid().ToString(), "<script language=""JavaScript""> opener=self;window.close();</script>")
        'System.Web.UI.ScriptManager.RegisterClientScriptBlock(P, P.GetType(), "Close_Window", "self.close();", True)
    End Sub

    Public Function GetAlertScript(ByVal message As String) As String
        'Return "HidePleaseWait();alert('" & message.Replace("'", "\'") & "');"
        Return "alert( '" & message.Replace("'", "\'") & "');"
    End Function

    Public Sub ShowInDialog(ByVal lcUrl As String, ByRef P As System.Web.UI.Page)
        Dim lcScript As String
        Dim CSM As ClientScriptManager = P.ClientScript
        lcScript = ""
        lcScript = lcScript + "window.showModalDialog('"
        lcScript = lcScript + "/Enquiry/Dialogs/ShowIFFixed.aspx?QS=" + lcUrl + "','newwin','dialogHeight:600px;dialogLeft:100px;dialogTop:50px;dialogWidth:800px;center:yes;dialogHide:yes;edge:raised;help:no;resizable:yes;scroll:no;status:yes;unadorned:yes' );"
        CSM.RegisterStartupScript(Me.GetType, Guid.NewGuid().ToString(), "<script language='JavaScript'>" & lcScript & "</script>")
    End Sub

    Public Sub ShowInDialogNew(ByVal lcUrl As String, ByRef P As System.Web.UI.Page)
        Dim lcScript As String
        Dim CSM As ClientScriptManager = P.ClientScript
        lcScript = ""
        lcScript = lcScript + "window.showModalDialog('"
        lcScript = lcScript + lcUrl + "','newwin','dialogHeight:400px;dialogLeft:100px;dialogTop:50px;dialogWidth:600px;center:yes;dialogHide:yes;edge:raised;help:no;resizable:yes;scroll:no;status:yes;unadorned:yes' );"
        CSM.RegisterStartupScript(Me.GetType, Guid.NewGuid().ToString(), "<script language='JavaScript'>" & lcScript & "</script>")
    End Sub

    Public Function GetDialogString(ByVal lcUrl As String, ByVal lcArgs As String) As String
        Dim lcScript As String = ""
        lcScript = lcScript + " var Parameters = new Object();" + vbCrLf
        lcScript = lcScript + " Parameters.P1= " + Chr(34) + lcArgs + Chr(34) + ";" + vbCrLf
        'lcScript = lcScript + " document.form1.txtPOPUP.value= " + Chr(34) + lcArgs + Chr(34) + ";" + vbCrLf
        lcScript = lcScript + " window.showModalDialog('" + lcUrl + "',Parameters"
        lcScript = lcScript + ",'dialogHeight:400px;dialogLeft:100px;dialogTop:50px;dialogWidth:600px;center:yes;dialogHide:yes;edge:raised;help:no;resizable:yes;scroll:no;status:yes;unadorned:yes' );"
        GetDialogString = lcScript
    End Function

    Public Function GetLookupDialog(ByVal lcSql As String) As String
        Dim lcScript As String = ""
        Dim lcUrl As String = "/Enquiry/Dialogs/LookupDlg.aspx?QS=" + lcSql
        lcScript = lcScript + " window.showModalDialog('" + lcUrl + "',''"
        lcScript = lcScript + ",'dialogHeight:400px;dialogLeft:100px;dialogTop:50px;dialogWidth:600px;center:yes;dialogHide:yes;edge:raised;help:no;resizable:yes;scroll:no;status:yes;unadorned:yes' );"
        GetLookupDialog = lcScript
    End Function

    Public Function GetMntDialog(ByVal lcUrl As String, Optional ByVal lnHeight As Integer = 600, Optional ByVal lnWidth As Integer = 800) As String
        Dim lcScript As String = ""
        GetMntDialog = lcScript
        Dim lcHeight As String = CStr(lnHeight) + "px"
        Dim lcWidth As String = CStr(lnWidth) + "px"
        lcScript = ""
        lcScript = lcScript + "window.showModalDialog('"
        lcScript = lcScript + "/Enquiry/Dialogs/ShowIFFixed.aspx?QS=" + lcUrl + "','newwin','dialogHeight:" + lcHeight + ";dialogLeft:100px;dialogTop:50px;dialogWidth:" + lcWidth + ";center:yes;dialogHide:yes;edge:raised;help:no;resizable:yes;scroll:no;status:yes;unadorned:yes' );"
        GetMntDialog = lcScript
    End Function

    'Public Function GetNewWindow(ByVal lcUrl As String, Optional ByVal lnHeight As Integer = 600, Optional ByVal lnWidth As Integer = 800) As String
    '    Dim lcScript As String = ""
    '    GetNewWindow = lcScript
    '    Dim lcHeight As String = CStr(lnHeight) + "px"
    '    Dim lcWidth As String = CStr(lnWidth) + "px"
    '    lcScript = ""
    '    lcScript = lcScript + "window.open('"
    '    lcScript = lcScript + "/Enquiry/Dialogs/ShowIFFixed.aspx?QS=" + lcUrl + "','','height=" + lcHeight + ",left=100px,top=50px,Width=" + lcWidth + ",center=yes,edge=raised,help=no,resizable=yes,scroll=no,status=yes,unadorned=yes' );"
    '    GetNewWindow = lcScript
    'End Function

    Public Function GetNewWindow(ByVal lcUrl As String, Optional ByVal lnHeight As Integer = 600, Optional ByVal lnWidth As Integer = 800) As String
        Dim lcScript As String = ""
        GetNewWindow = lcScript
        Dim lcHeight As String = CStr(lnHeight) + "px"
        Dim lcWidth As String = CStr(lnWidth) + "px"
        lcScript = ""
        lcScript = lcScript + "window.open('"
        lcScript = lcScript + GetAppPath() + "/Dialogs/ShowIFFixed.aspx?QS=" + lcUrl + "','','height=" + lcHeight + ",left=100px,top=50px,Width=" + lcWidth + ",center=yes,edge=raised,help=no,resizable=yes,scroll=no,status=yes,unadorned=yes' );"
        GetNewWindow = lcScript
    End Function

    Function GetAppPath() As String
        Dim lcPath As String = HttpRuntime.AppDomainAppVirtualPath
        HttpRuntime.UnloadAppDomain()
        GetAppPath = lcPath
        If lcPath = "\" Or lcPath = "/" Then
            GetAppPath = ""
        End If
    End Function


    Public Function GetWindow(ByVal lcUrl As String, Optional ByVal lcHeight As String = "600px", Optional ByVal lcWidth As String = "800px") As String
        Dim lcScript As String = ""
        GetWindow = lcScript
        lcScript = ""
        lcScript = lcScript + "window.open('"
        lcScript = lcScript + "/Enquiry/Dialogs/ShowIFFixed.aspx?QS=" + lcUrl + "',null,'dialogHeight:" + lcHeight + ";dialogLeft:100px;dialogTop:50px;dialogWidth:" + lcWidth + ";center:yes;dialogHide:yes;edge:raised;help:no;resizable:yes;scroll:no;status:yes;unadorned:yes' );"
        GetWindow = lcScript
    End Function

    Public Function GetRequestDialog(ByVal lcUrl As String, Optional ByVal lnHeight As Integer = 600, Optional ByVal lnWidth As Integer = 800) As String
        Dim lcScript As String = ""
        GetRequestDialog = lcScript
        Dim lcWidth As String = CStr(lnWidth) + "px"
        Dim lcHeight As String = CStr(lnHeight) + "px"
        lcScript = ""
        lcScript = lcScript + "window.showModalDialog('"
        lcScript = lcScript + "/Enquiry/Dialogs/RequestDialog.aspx?" + lcUrl + " ','newwin','dialogHeight:" + lcHeight + ";dialogLeft:100px;dialogTop:50px;dialogWidth:" + lcWidth + ";center:yes;dialogHide:yes;edge:raised;help:no;resizable:yes;scroll:no;status:yes;unadorned:yes' );"
        GetRequestDialog = lcScript
    End Function

    Public Function GetPopUpDialog(ByVal lcUrl As String, Optional ByVal lcHeight As String = "600px", Optional ByVal lcWidth As String = "800px") As String
        Dim lcScript As String = ""
        lcScript = ""
        lcScript = lcScript + "window.showModalDialog('"
        lcScript = lcScript + "/Enquiry/Dialogs/ShowIFFixed.aspx?QS=" + lcUrl + "','POPUP','dialogHeight:" + lcHeight + ";dialogLeft:100px;dialogTop:50px;dialogWidth:" + lcWidth + ";center:yes;dialogHide:yes;edge:raised;help:no;resizable:yes;scroll:no;status:yes;unadorned:yes' );"
        GetPopUpDialog = lcScript
    End Function

    Public Sub ShowRequestDialog(ByVal lcUrl As String, ByRef P As System.Web.UI.Page)
        Dim lcScript As String
        Dim CSM As ClientScriptManager = P.ClientScript
        lcScript = ""
        lcScript = lcScript + "window.showModalDialog('"
        lcScript = lcScript + "/Enquiry/ShowIFFixed.aspx?QS=" + lcUrl + "','newwin','dialogHeight:800px;dialogLeft:100px;dialogTop:50px;dialogWidth:1000px;center:yes;dialogHide:yes;edge:raised;help:no;resizable:yes;scroll:no;status:yes;unadorned:yes' );"
        CSM.RegisterStartupScript(Me.GetType, Guid.NewGuid().ToString(), "<script language='JavaScript'>" & lcScript & "</script>")
    End Sub

    Public Sub CreateConfirmBox(ByRef btn As WebControls.Button, ByVal strMessage As String)
        btn.Attributes.Add("onclick", "return confirm('" & strMessage & "');")
    End Sub

    Public Sub CreateConfirmBoxLinkBtn(ByRef btn As WebControls.LinkButton, ByVal strMessage As String)
        btn.Attributes.Add("onclick", "return confirm('" & strMessage & "');")
    End Sub


End Class

Public Class Host
    Dim udf As New GUDF
    Dim AccountsTypes As New Collection
    Dim CustomersTypes As New Collection
    Public TransactionsTypes As New Collection
    Dim LoansTypes As New Collection
    Dim RespCodes As New Collection
    Function GetCurrency(ByVal lcCCY As String) As String
        'Dim db As New ADODB.Connection
        'Dim rs As New ADODB.Recordset
        'GetCurrency = ""
        'Dim lcSQL = "SELECT C8CUR FROM C8PF WHERE C8CCY='" + lcCCY + "'"
        'Try
        '    If udf.OpenDsn(db, rs, "KFILLIV", lcSQL) Then GetCurrency = Trim(rs("C8CUR").Value)
        'Catch
        'End Try
        'rs.Close()
        'db.Close()
    End Function

    Function GetAccountType(ByVal lcType As String) As String
        'GetAccountType = ""
        'If AccountsTypes.Contains(lcType) Then
        '    GetAccountType = AccountsTypes(lcType)
        'Else
        '    Dim db As New ADODB.Connection
        '    Dim rs As New ADODB.Recordset
        '    GetAccountType = ""
        '    Dim lcSQL = "SELECT C5ATD FROM C5PF WHERE C5ATP='" + lcType + "'"
        '    Try
        '        If udf.OpenDsn(db, rs, "KFILLIV", lcSQL) Then GetAccountType = Trim(rs("C5ATD").Value)
        '    Catch
        '    End Try
        '    rs.Close()
        '    db.Close()
        'End If
    End Function

    Public Sub GetTransactionsTypes()
        'udf.DT2COLLECTION(udf.GetSqlDataTable("KFILLIV", "SELECT CTTCD,CTTCN FROM CTPF"), TransactionsTypes)
    End Sub

    Function GetAccountsList(ByVal lcCIF As String) As String()
        'Dim DB As New ADODB.Connection
        'Dim RS As New ADODB.Recordset
        'Dim lcAccounts = ""
        'Dim lcSQL = "SELECT SCAB,SCAn,SCAS FROM SCPF WHERE SCAN='" + lcCIF + "' and SCACT in ('EB','CA','EA','CY')"
        'udf.OpenDsn(DB, RS, "KFILLIV", lcSQL)
        'Do While Not RS.EOF
        '    lcAccounts = lcAccounts + "~" + RS("SCAB").Value + "-" + RS("SCAN").Value + "-" + RS("SCAS").Value + vbTab _
        '               + RS("SCAB").Value + "-" + RS("SCAN").Value + "-" + RS("SCAS").Value
        '    RS.MoveNext()
        'Loop
        'RS.Close()
        'DB.Close()
        'GetAccountsList = Split(Mid(lcAccounts, 2), "~")
    End Function

    'Public Sub GetAccountsTypes()
    '    udf.DT2COLLECTION(udf.GetSqlDataTable("KFILLIV", "SELECT C5ATP, C5ATD FROM C5PF"), AccountsTypes)
    'End Sub

    'Public Sub GetCustomersTypes()
    '    udf.DT2COLLECTION(udf.GetSqlDataTable("KFILLIV", "SELECT C4CTP,C4CTD FROM C4PF"), CustomersTypes)

    'End Sub

    Function FormatDate(ByVal lcDate As String) As String
        Dim lcdate1 As String = CStr(Val(CStr(lcDate)) + 19000000)
        FormatDate = Mid(lcdate1, 1, 4) + "-" + Mid(lcdate1, 5, 2) + "-" + Mid(lcdate1, 7, 2)
    End Function

  
End Class

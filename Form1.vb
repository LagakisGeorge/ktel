﻿Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Xml
Imports System.IO
Imports System.Data.OleDb
Imports System.Xml.Schema
Imports System.Data.SqlClient
Imports System.Web
Imports System.CodeDom.Compiler
Imports Newtonsoft.Json



Imports Newtonsoft.Json.Linq
Imports CefSharp.DevTools.DOM
Imports CefSharp.Enums
Imports System.Drawing.Drawing2D
Imports System.Globalization
Imports Microsoft.VisualBasic.ApplicationServices


'Imports System.Net.Http.Headers
'Imports System.Text
'Imports System.Net.Http

'ΑΛΛΑΓΕΣ ΣΤΟ ΤΙΜ 
'    [ENTITYUID] [varchar](40) NULL,
'	[ENTITYMARK] [varchar](43) NULL,
'	[ENTITY] [int] NULL,
'	[AADEKAU] [float] NULL,
'	[AADEFPA] [float] NULL,
'	[ENTLINEN] [int] NULL,
'	[INCMARK] [nvarchar](43) NULL,

Public Class Form1



    Dim reader As JsonTextReader
    Public openedFileStream As System.IO.Stream
    Public dataBytes() As Byte
    Public gConnect As String
    Public gSQLCon As String
    Public sqlDT As New DataTable
    Public sqlDT2 As New DataTable
    Public gUserId As String
    Public gSubKey As String

    Public fbearer As String



    ' Public Property HttpUtility As Object
    '  SELECT ENTITY ,ATIM,ENTITYUID,ENTITYMARK ,HME FROM TIM WHERE ENTITY>0 ORDER BY ENTITY'

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MakeRequest()

    End Sub


    Private Async Sub MakeRequest()
        ' ΣΤΕΛΝΩ ΣΤΟ SENDINVOICES 
        ' "c:\txtfiles\inv.xml").ToString ' "--> εκει έχω αποθηκεύσει το xml που εφτιαξα"
        ' H APANTHSH EINAI STO  "c:\txtfiles\apantSendInv.XML")




        Dim client = New HttpClient()
        'Dim queryString = HttpUtility.ParseQueryString(String.Empty)
        Try
            client.DefaultRequestHeaders.Add("aade-user-id", gUserId) '"glagakis2")
            client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", gSubKey) ' "555bc57c80634243958f62b629316aaa")

            Dim uri = "https://mydata-dev.azure-api.net/SendInvoices" ' + queryString.ToString

            Dim response As HttpResponseMessage
            Dim xl As String = XDocument.Load("c:\txtfiles\inv.xml").ToString ' "--> εκει έχω αποθηκεύσει το xml που εφτιαξα"
            Dim byteData As Byte() = Encoding.UTF8.GetBytes(xl)

            Using content = New ByteArrayContent(byteData)
                content.Headers.ContentType = New MediaTypeHeaderValue("application/xml")
                response = Await client.PostAsync(uri, content)
                Dim result = Await response.Content.ReadAsStringAsync()
                Dim MF = "c:\txtfiles\sendinv\apantSendInv" + Format(Now, "yyyyddMMHHmm") + ".xml"
                FileOpen(1, MF, OpenMode.Output)
                PrintLine(1, result.ToString)
                FileClose(1)
                TextBox2.Text = result.ToString
                ' "είναι το textbox πανω στη φόρμα που σου επιστρέφει το response xml"
                'Dim byteData2 As Byte() = File.ReadAllBytes("c:\txtfiles\inv.xml")
                ' sept 2020 debug   Rename("c:\txtfiles\inv.xml", "c:\txtfiles\inv" + Format(Now, "yyyyddMMHHmm") + ".xml")
                FileCopy(MF, "c:\txtfiles\apantSendInv.XML")

            End Using


        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        MakeIncomeRequest()
    End Sub

    Private Async Sub MakeIncomeRequest()
        'ΠΑΙΡΝΩ ΤΟ ΑΡΧΕΙΟ  c:\txtfiles\inc.xml"
        ' ΚΑΙ ΣΤΕΛΝΩ ΤΟ SendIncomeClassification" 
        ' Η ΑΠΑΝΤΗΣΗ ΕΙΝΑΙ ΤΟ "c:\txtfiles\apantiNCOMe" + Format(Now, "yyyyddMMHHmm") + ".xml"





        ListBox2.Items.Clear()

        Dim client = New HttpClient()
        'Dim queryString = HttpUtility.ParseQueryString(String.Empty)
        Try
            client.DefaultRequestHeaders.Add("aade-user-id", "glagakis2")
            client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", "555bc57c80634243958f62b629316aaa")

            Dim uri = "https://mydata-dev.azure-api.net/SendIncomeClassification" ' + queryString.ToString

            Dim response As HttpResponseMessage
            Dim xl = XDocument.Load("c:\txtfiles\inc.xml").ToString ' "--> εκει έχω αποθηκεύσει το xml που εφτιαξα"
            Dim byteData As Byte() = Encoding.UTF8.GetBytes(xl)

            Using content = New ByteArrayContent(byteData)
                content.Headers.ContentType = New MediaTypeHeaderValue("application/xml")
                response = Await client.PostAsync(uri, content)
                Dim result = Await response.Content.ReadAsStringAsync()

                TextBox2.Text = result.ToString
                ' "είναι το textbox πανω στη φόρμα που σου επιστρέφει το response xml"

                Dim MF = "c:\txtfiles\incomes\ApantIncome" + Format(Now, "yyyyddMMHHmm") + ".xml"
                FileOpen(1, MF, OpenMode.Output)
                PrintLine(1, result.ToString)
                FileClose(1)
                TextBox2.Text = result.ToString
                ' "είναι το textbox πανω στη φόρμα που σου επιστρέφει το response xml"
                'Dim byteData2 As Byte() = File.ReadAllBytes("c:\txtfiles\inv.xml")
                ' Rename("c:\txtfiles\inv.xml", "c:\txtfiles\inv" + Format(Now, "yyyyddMMHHmm") + ".xml")
                FileCopy(MF, "c:\txtfiles\ApantIncome.XML")

            End Using


        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


    End Sub





    Function FINDTYPOS(C As String) As String
        Dim sqlDT4 As New DataTable
        ExecuteSQLQuery("select ETIK from PARASTAT where EIDOS='" + C + "'", sqlDT4)
        If IsDBNull(sqlDT4(0)(0)) Then
            FINDTYPOS = ""
        Else
            FINDTYPOS = sqlDT4(0)(0).ToString
        End If


        FINDTYPOS = FINDTYPOS.Trim()

        ' If InStr(FINDTYPOS, ";") = 0 Then
        FINDTYPOS = FINDTYPOS + ";;" ' gia na mhn skaei to split()
        ' Else


        ' End If


    End Function
    Function FindTRP(C As String) As String
        Dim sqlDT4 As New DataTable
        ExecuteSQLQuery("select N1 from PINAKES where TYPOS=12 AND  AYJON=" + C + "", sqlDT4)
        If IsDBNull(sqlDT4(0)(0)) Then
            FindTRP = ""
        Else
            FindTRP = sqlDT4(0)(0).ToString
        End If

    End Function



    Function Get_AJ_ASCII(ByRef pol As String,
                          ByVal polepis As String,
                          ByVal ago As String,
                          ByVal AGOEPIS As String) As Boolean

        '<EhHeader>


        '</EhHeader>



        ' Dim R As New ADODB.Recordset
        Dim x As String

        'If gConnect = "Access" Then
        '   Set db = OpenDatabase(gDir, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If

        ExecuteSQLQuery("select POL,EIDOS,AJIA_APOU from PARASTAT", sqlDT2)

        pol = " "

        Dim row As Integer
        For row = 0 To sqlDT2.Rows.Count - 1

            If IsDBNull(sqlDT2.Rows(row)("eidos")) Or IsDBNull(sqlDT2.Rows(row)("pol")) Or IsDBNull(sqlDT2.Rows(row)("ajia_apou")) Then

            Else

                If sqlDT2.Rows(row)("pol") = "1" And sqlDT2.Rows(row)("ajia_apou") = "3" Then
                    pol = pol + "'" + sqlDT2.Rows(row)("eidos") + "',"
                End If

                If sqlDT2.Rows(row)("pol") = "1" And sqlDT2.Rows(row)("ajia_apou") = "4" Then
                    polepis = polepis + "'" + sqlDT2.Rows(row)("eidos") + "',"
                End If

                If sqlDT2.Rows(row)("pol") = "2" And sqlDT2.Rows(row)("ajia_apou") = "1" Then
                    ago = ago + "'" + sqlDT2.Rows(row)("eidos") + "',"
                End If

                If sqlDT2.Rows(row)("pol") = "2" And sqlDT2.Rows(row)("ajia_apou") = "2" Then
                    AGOEPIS = AGOEPIS + "'" + sqlDT2.Rows(row)("eidos") + "',"
                End If




            End If
        Next

240:    pol = Mid(pol, 1, Len(pol) - 1)

250:    If Len(polepis) > 0 Then
260:        polepis = Mid(polepis, 1, Len(polepis) - 1)
        Else
270:        polepis = "999"  'ME KENO DHMIOYRGEI PROBLHMA
        End If

280:
290:    ago = Mid(ago, 1, Len(ago) - 1)
300:    Get_AJ_ASCII = True

350:    If Len(AGOEPIS) > 0 Then
360:        AGOEPIS = Mid(AGOEPIS, 1, Len(AGOEPIS) - 1)
        Else
370:        AGOEPIS = "999" 'ME KENO DHMIOYRGEI PROBLHMA
        End If


    End Function




    Private Sub ToXML_Click(sender As Object, e As EventArgs) Handles toXML.Click
        Dim I As Integer = ToXMLsub()
    End Sub
    Private Function ToXMLsub() As Integer

        '===ΒΓΑΖΩ ΤΟ XML ΓΙΑ ΤΑ ΠΑΡΑΣΤΑΤΙΚΑ =================================================================================
        'WHERE (ENTITYMARK IS NULL OR ENTITYMARK='ERROR' ) AND   
        'Left(ATIM, 1) In     (  " + PAR + "  )    And 
        'HME>='" + Format(APO.Value, "MM/dd/yyyy") + "'  AND HME<='" + Format(EOS.Value, "MM/dd/yyyy") + "'  "
        '<correlatedInvoices>400000017716190</correlatedInvoices>
        ToXMLsub = 1



        Dim pol, polepis, ago, AGOEPIS As String


        If checkServer(0) Then
            ' MsgBox("OK")
        End If
        ExecuteSQLQuery("UPDATE TIM SET ENTITY=0,ENTLINEN=0")

        Get_AJ_ASCII(pol, polepis, ago, AGOEPIS)

        Dim PAR, SYNT As String
        PAR = pol + polepis
        Dim SQL As String
        SYNT = ""
        SQL = "SELECT  ID_NUM, AJ1  ,AJ2 , AJ3,AJ4,AJ5,AJI,FPA1,FPA2,FPA3,FPA4,ATIM,"
        SQL = SQL + "HME,PEL.EPO,PEL.AFM,KPE,PEL.DIE,PEL.XRVMA"    '"CONVERT(CHAR(10),HME,3) AS HMEP
        SQL = SQL + ",PEL.EPA,PEL.POL,AJ6,FPA6,AJ7,FPA7,TRP,APALAGIFPA "

        SQL = SQL + "   FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD "
        SQL = SQL + " WHERE (ENTITYMARK IS NULL OR ENTITYMARK='ERROR' ) AND    LEFT(ATIM,1) IN     (  " + PAR + "  )    and HME>='" + Format(APO.Value, "MM/dd/yyyy") + "'  AND HME<='" + Format(EOS.Value, "MM/dd/yyyy") + "'  "
        SQL = SQL + "  AND AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7>0  " + SYNT
        SQL = SQL + " order by HME"       '  OR INCMARK IS NULL OR INCMARK='ERROR'




        '  SQL = "SELECT  top 20  AJ1 ,AJ2  from TIM  order by HME"

        ExecuteSQLQuery(SQL)

        If sqlDT.Rows.Count = 0 Then
            MsgBox("ΔΕΝ ΒΡΕΘΗΚΑΝ ΕΓΓΡΑΦΕΣ")
            ToXMLsub = 0
            Exit Function
        End If



        Dim ff As String = "c:\txtfiles\inv.xml"  'c:\mercvb\m" + Format(Now, "yyyyddmmHHMM") + ".export" ' "\\Logisthrio\333\pr.export" '
        Dim writer As New XmlTextWriter(ff, System.Text.Encoding.UTF8)
        writer.WriteStartDocument(True)
        writer.Formatting = Xml.Formatting.Indented
        writer.Indentation = 2
        writer.WriteStartElement("InvoicesDoc")
        writer.WriteAttributeString("xmlns", "http://www.aade.gr/myDATA/invoice/v1.0")
        writer.WriteAttributeString("xsi:schemaLocation", "http://www.aade.gr/myDATA/invoice/v1.0 schema.xsd")
        writer.WriteAttributeString("xmlns:N1", "https://www.aade.gr/myDATA/incomeClassificaton/v1.0")
        writer.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")



        Dim i As Integer
        Dim sumNet As Single
        Dim sumFpa As Single
        Dim ctypos As String

        '======================================= ΠΕΡΠΑΤΑΩ ΤΟ ΤΙΜ ====================================
        For i = 0 To sqlDT.Rows.Count - 1



            If checkIntegrity(i) = False Then

                MsgBox(" ΠΡΟΒΛΗΜΑ ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ " + sqlDT(0)("ATIM").ToString)
            End If



            Dim EGGTIM As New DataTable
            Me.Text = "ΠΑΡΑΣΤΑΤΙΚΑ " + Str(i)
            ExecuteSQLQuery("SELECT KODE,POSO,TIMM,EKPT,FPA,ISNULL(KAU_AJIA,0) AS KAU_AJIA,ISNULL(MIK_AJIA,0) AS MIK_AJIA FROM EGGTIM WHERE TIMM<>0 AND POSO<>0 AND ID_NUM=" + sqlDT(i)("ID_NUM").ToString, EGGTIM)
            Dim DUM As New DataTable
            ExecuteSQLQuery("UPDATE TIM SET ENTITY=" + Str(i + 1) + " WHERE ID_NUM=" + sqlDT(i)("ID_NUM").ToString, DUM)




            sumNet = sqlDT(i)("aj1") + sqlDT(i)("aj2") + sqlDT(i)("aj3") + sqlDT(i)("aj4") + sqlDT(i)("aj5") + sqlDT(i)("aj6") + sqlDT(i)("aj7")
            sumFpa = sqlDT(i)("fpa1") + sqlDT(i)("fpa2") + sqlDT(i)("fpa3") + sqlDT(i)("fpa4") + sqlDT(i)("fpa6") + sqlDT(i)("fpa7")

            '1.1;E3_561_001;category1_1
            ctypos = FINDTYPOS(Mid(sqlDT(i)("ATIM"), 1, 1)) ' Split(tmpStr, ":")(0)
            If Len(Trim(Split(ctypos, ";")(1))) = 0 Or Len(Trim(Split(ctypos, ";")(2))) = 0 Then
                writer.Close()
                MsgBox("δεν εχουν ορισθει παραμετροι ΜΥDATA στο παρ/κό " + sqlDT(i)("ATIM"))
                Exit Function

            End If


            writer.WriteComment(sqlDT(i)("ATIM") + " " + Format(sqlDT(i)("HME"), "dd/MM/yyyy"))
            writer.WriteStartElement("invoice")
            crNode("uid", Str(i + 1), writer)
            'crNode("mark", "", writer)


            '---------------------------------------------εκδοτης
            writer.WriteStartElement("issuer")
            crNode("vatNumber", "028783755", writer)
            crNode("country", "GR", writer)
            crNode("branch", "0", writer)
            ' writer.WriteStartElement("address")
            ' crNode("postalCode", """66100""", writer)
            ' crNode("city", """ΔΡΑΜΑ""", writer)
            ' writer.WriteEndElement() '/address
            writer.WriteEndElement() '/issuer

            '--------------------------------------------- πελατης
            If Mid(Split(ctypos, ";")(0), 1, 2) = "11" Then
                'lianikh den xreiazetai pelaths
            Else
                ' End If
                writer.WriteStartElement("counterpart")
                crNode("vatNumber", Trim(sqlDT(i)("AFM")), writer)  ' crNode("vatNumber", "026677115", writer)
                crNode("country", "GR", writer)
                crNode("branch", "0", writer)
                writer.WriteStartElement("address")
                crNode("postalCode", """66100""", writer)  'crNode("postalCode", """66100""", writer)
                crNode("city", sqlDT(i)("POL"), writer)  ' crNode("city", """ΔΡΑΜΑ""", writer)
                writer.WriteEndElement() ' /address
                writer.WriteEndElement() ' /counterpart
            End If


            '----------------------------------------------- header
            writer.WriteStartElement("invoiceHeader")
            crNode("series", "0", writer)
            crNode("aa", Mid(sqlDT(i)("ATIM"), 2, 6), writer)   '  crNode("aa", "15", writer)
            crNode("issueDate", Format(sqlDT(i)("hme"), "yyyy-MM-dd"), writer) ' crNode("issueDate", "2019-12-15", writer)


            'ctypos = FINDTYPOS(Mid(sqlDT(i)("ATIM"), 1, 1)) ' Split(tmpStr, ":")(0)
            'If Len(Trim(Split(ctypos, ";")(1))) = 0 Or Len(Trim(Split(ctypos, ";")(2))) = 0 Then
            '    writer.Close()
            '    MsgBox("δεν εχουν ορισθει παραμετροι ΜΥDATA στο παρ/κό " + sqlDT(i)("ATIM"))
            '    Exit Function

            'End If


            crNode("invoiceType", Split(ctypos, ";")(0), writer)   ' ειδος παραστατικού
            crNode("currency", "EUR", writer)
            crNode("exchangeRate", "1.0", writer)
            writer.WriteEndElement() ' /invoiceHeader

            '          <paymentMethods>
            '	<paymentMethodDetails>
            '		<type>3</type>
            '		<amount>66.53</amount>
            '	</paymentMethodDetails>
            '</paymentMethods>
            '----------------------------------------------- paymentMethods
            writer.WriteStartElement("paymentMethods")
            writer.WriteStartElement("paymentMethodDetails")
            Dim cTrp As String = FindTRP(Mid(sqlDT(i)("TRP"), 1, 1))
            If Len(cTrp) = 0 Then
                MsgBox("ΔΕΝ ΕΧΩ ΑΝΤΙΣΤΟΙΧΙΣΗ ΣΤΟΝ ΤΡΟΠΟ ΠΛΗΡΩΜΗΣ " + sqlDT(i)("TRP"))
                writer.Close()
                Exit Function
            End If
            crNode("type", cTrp, writer)
            crNode("amount", Format(sumNet + sumFpa, "######0.##"), writer)   '  crNode("aa", "15", writer)

            writer.WriteEndElement() ' /paymentMethodDetails
            writer.WriteEndElement() ' /paymentMethods


            Dim SYN_KAU, SYN_FPA As Double
            SYN_KAU = 0
            SYN_FPA = 0
            Dim fpaRow As Double

            '======================================= ΠΕΡΠΑΤΑΩ ΤΟ EGGΤΙΜ ====================================
            For L As Integer = 0 To EGGTIM.Rows.Count - 1

                Dim AJ As Single
                If IsDBNull(EGGTIM(L)("TIMM")) Then
                    AJ = 0
                Else
                    AJ = EGGTIM(L)("KAU_AJIA")  ' Math.Round(EGGTIM(L)("POSO") * EGGTIM(L)("TIMM") * (1 - EGGTIM(L)("EKPT") / 100), 2)
                End If

                Dim VAT As String
                '1 ΦΠΑ συντελεστής 24% 24% 
                '2 ΦΠΑ συντελεστής 13% 13% 
                '3 ΦΠΑ συντελεστής 6% 6% 
                '4 ΦΠΑ συντελεστής 17% 17% 
                '5 ΦΠΑ συντελεστής 9% 9% 
                '6 ΦΠΑ συντελεστής 4% 4% 
                '7 Άνευ Φ.Π.Α. 0%  
                '8 Εγγραφές χωρίς ΦΠΑ  (πχ Μισθοδοσία, Αποσβέσεις) 



                If EGGTIM(L)("FPA") = 1 Then '13%
                    VAT = "2"
                    SYN_KAU = SYN_KAU + AJ
                    fpaRow = EGGTIM(L)("MIK_AJIA") - EGGTIM(L)("KAU_AJIA") ' AJ * 0.13
                    SYN_FPA = SYN_FPA + fpaRow

                    'ElseIf EGGTIM(L)("FPA") = 2 Then
                    '   VAT = "1"
                ElseIf EGGTIM(L)("FPA") = 2 Then
                    VAT = "1"
                    SYN_KAU = SYN_KAU + AJ
                    fpaRow = EGGTIM(L)("MIK_AJIA") - EGGTIM(L)("KAU_AJIA") 'AJ * 0.24
                    SYN_FPA = SYN_FPA + fpaRow

                    ' SYN_FPA = SYN_FPA + AJ * 0.24

                ElseIf EGGTIM(L)("FPA") = 5 Then
                    VAT = "7"
                    SYN_KAU = SYN_KAU + AJ
                    fpaRow = 0
                    SYN_FPA = SYN_FPA + fpaRow

                ElseIf EGGTIM(L)("FPA") = 6 Then
                    VAT = "1"
                    SYN_KAU = SYN_KAU + AJ
                    ' SYN_FPA = SYN_FPA + AJ * 0.24

                    fpaRow = EGGTIM(L)("MIK_AJIA") - EGGTIM(L)("KAU_AJIA")  ' AJ * 0.24
                    SYN_FPA = SYN_FPA + fpaRow




                ElseIf EGGTIM(L)("FPA") = 4 Then
                    VAT = "4"
                    SYN_KAU = SYN_KAU + AJ
                    fpaRow = EGGTIM(L)("MIK_AJIA") - EGGTIM(L)("KAU_AJIA")  ' AJ * 0.06
                    SYN_FPA = SYN_FPA + fpaRow
                Else ' If EGGTIM(L)("FPA") = 2 Then
                    VAT = "1"
                End If
                '-----------------------------------------------  invoiceDetails
                writer.WriteStartElement("invoiceDetails")
                crNode("lineNumber", Str(L + 1), writer) '  crNode("lineNumber", "1", writer)
                '    crNode("quantity", EGGTIM(L)("POSO").ToString, writer)
                '    crNode("measurementUnit", "1", writer)
                crNode("netValue", Format(AJ, "######0.##"), writer)  ' crNode("netValue", "100", writer)

                crNode("vatCategory", VAT, writer) '1=24%   2=13%   

                crNode("vatAmount", Format(fpaRow, "######0.##"), writer)  ' c

                If fpaRow = 0 Then

                    crNode("vatExemptionCategory", "1", writer) 'APALAGIFPA

                End If

                writer.WriteStartElement("incomeClassification")
                crNode("N1:classificationType", Split(ctypos, ";")(1), writer)
                crNode("N1:classificationCategory", Split(ctypos, ";")(2), writer)
                crNode("N1:amount", Format(AJ, "######0.##"), writer)

                writer.WriteEndElement() '/incomeClassification





                '                <invoiceDetails>
                '  <lineNumber> 1</lineNumber>
                '  <netValue>1185</netValue>
                '  <vatCategory>7</vatCategory>
                '  <vatAmount>0</vatAmount>
                '<vatExemptionCategory>1</vatExemptionCategory>
                '  <incomeClassification>
                '    <N1:classificationType> E3_561_001</N1:classificationType>
                '    <N1:classificationCategory>category1_1</N1:classificationCategory>
                '    <N1:amount> 1185</N1:amount>
                '  </incomeClassification>

                '</invoiceDetails>



                '            	<incomeClassification>
                '<N1:classificationType> E3_561_001</N1:classificationType>
                '            <N1:classificationCategory>category1_1</N1:classificationCategory>
                '<N1:amount> 100.0</N1:amount>
                '    </incomeClassification>


                writer.WriteEndElement()   ' /invoiceDetails
            Next

            ExecuteSQLQuery("UPDATE TIM SET AADEKAU=" + Replace(Format(SYN_KAU, "######0.#####"), ",", ".") + ",AADEFPA=" + Replace(Format(SYN_FPA, "######0.#####"), ",", ".") +
                            " WHERE ID_NUM=" + sqlDT(i)("ID_NUM").ToString, DUM)
            '------------------------------------------------ InvoiceSummary 
            writer.WriteStartElement("invoiceSummary")
            crNode("totalNetValue", Format(SYN_KAU, "######0.##"), writer)  ' crNode("totalNetValue", "100", writer)
            crNode("totalVatAmount", Format(SYN_FPA, "######0.##"), writer)  '  crNode("totalVatAmount", "24", writer)
            crNode("totalWithheldAmount", "0", writer)
            crNode("totalFeesAmount", "0", writer)
            crNode("totalStampDutyAmount", "0", writer)
            crNode("totalOtherTaxesAmount", "0", writer)
            crNode("totalDeductionsAmount", "0", writer)
            crNode("totalGrossValue", Format(SYN_KAU + SYN_FPA, "######0.##"), writer)


            writer.WriteStartElement("incomeClassification")
            crNode("N1:classificationType", Split(ctypos, ";")(1), writer)
            crNode("N1:classificationCategory", Split(ctypos, ";")(2), writer)
            crNode("N1:amount", Format(SYN_KAU, "######0.##"), writer)
            writer.WriteEndElement() '  /invoicesummary




            writer.WriteEndElement() '  /invoicesummary
            '=========================================================
            writer.WriteEndElement() ' / Invoice

        Next





        writer.WriteEndElement() 'InvoicesDoc








        writer.WriteEndDocument()
        writer.Close()
        '  MsgBox("ok")


        ListBox2.Items.Clear()



        '------ τοπικος ελεγχος xml που τον καταργησα γιατι μπηκε και το "https://www.aade.gr/myDATA/incomeClassificaton/v1.0
        'FileOpen(1, "C:\TXTFILES\CHECKXSD.TXT", OpenMode.Output)
        ''        Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim myDocument As New XmlDocument
        'myDocument.Load(ff) ' m_filename)  ' "C:\somefile.xml"
        'myDocument.Schemas.Axmlns:N1dd("http://www.aade.gr/myDATA/invoice/v1.0", "c:\txtfiles\invoicesDoc-v0.6.xsd") 'namespace here or empty string
        'Dim eventHandler As ValidationEventHandler = New ValidationEventHandler(AddressOf ValidationEventHandler)
        'myDocument.Validate(eventHandler)
        ''       MsgBox("ok ελεγχος")
        'For n As Integer = 0 To ListBox2.Items.Count - 1
        '    PrintLine(1, ListBox2.Items(n).ToString)
        'Next
        'FileClose(1)






        paint_ergasies(DataGridView1, "SELECT   ATIM,HME,ENTITY,AADEKAU,AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7 AS KAUTIM,AADEFPA,FPA1+FPA2+FPA3+FPA4+FPA6+FPA7 AS FPATIM,ENTITYUID,ENTITYMARK FROM TIM WHERE ENTITY>0")

    End Function

    Private Function checkIntegrity(I As Long) As Boolean
        Dim S_AJ As Double = sqlDT(I)("aj1") + sqlDT(I)("aj2") + sqlDT(I)("aj3") + sqlDT(I)("aj4") + sqlDT(I)("aj5") + sqlDT(I)("aj6") + sqlDT(I)("aj7")
        Dim S_FPA As Double = sqlDT(I)("fpa1") + sqlDT(I)("fpa2") + sqlDT(I)("FPA3") + sqlDT(I)("FPA4") + sqlDT(I)("FPA6") + sqlDT(I)("FPA7")
        Dim DIFF As Double = sqlDT(I)("aji") - (S_AJ + S_FPA)
        If DIFF = 0 Then
            'OK
        ElseIf Math.Abs(DIFF) < 0.05 Then ' ΕΧΟΥΝ ΜΙΚΡΗ ΔΙΑΦΟΡΑ ΚΑΙ ΤΗΝ ΚΑΛΥΠΤΩ ΑΛΛΑΖΟΝΤΑ ΤΟ AJI
            Dim DUM As New DataTable
            ExecuteSQLQuery("UPDATE TIM SET AJI=AJI+" + Replace(Str(DIFF), ",", ".") + "   WHERE ID_NUM=" + sqlDT(I)("ID_NUM").ToString, DUM)
        Else
            ' Dim DUM As New DataTable
            ' ExecuteSQLQuery("SELECT SUM(POSO*TIMM*(100-EKPT)/100)   WHERE ID_NUM=" + sqlDT(I)("ID_NUM").ToString, DUM)
            checkIntegrity = False
            Exit Function
        End If

        '    Dim DUM150 As New DataTable
        ' ExecuteSQLQuery("UPDATE EGGTIM 
        '        SET MIK_AJIA=ROUND(POSO*TIMM*(100-EKPT)/100*1.24,2)
        '     ,KAU_AJIA=ROUND(POSO*TIMM*(100-EKPT)/100,2)  ", DUM150)





        Dim DUM10 As New DataTable
        ExecuteSQLQuery("SELECT SUM(KAU_AJIA) AS KAU,SUM(MIK_AJIA) AS MIK,
             SUM(CASE WHEN FPA=1 THEN KAU_AJIA ELSE 0 END ) AS KAU1,
             SUM(CASE WHEN FPA=2 THEN KAU_AJIA ELSE 0 END ) AS KAU2,
             SUM(CASE WHEN FPA=3 THEN KAU_AJIA ELSE 0 END ) AS KAU3,
             SUM(CASE WHEN FPA=4 THEN KAU_AJIA ELSE 0 END ) AS KAU4,
             SUM(CASE WHEN FPA=5 THEN KAU_AJIA ELSE 0 END ) AS KAU5,
             SUM(CASE WHEN FPA=6 THEN KAU_AJIA ELSE 0 END ) AS KAU6,
             SUM(CASE WHEN FPA=7 THEN KAU_AJIA ELSE 0 END ) AS KAU7,
             SUM(CASE WHEN FPA=1 THEN MIK_AJIA ELSE 0 END ) AS MIK1,
             SUM(CASE WHEN FPA=2 THEN MIK_AJIA ELSE 0 END ) AS MIK2,
             SUM(CASE WHEN FPA=3 THEN MIK_AJIA ELSE 0 END ) AS MIK3,
             SUM(CASE WHEN FPA=4 THEN MIK_AJIA ELSE 0 END ) AS MIK4,
             SUM(CASE WHEN FPA=5 THEN MIK_AJIA ELSE 0 END ) AS MIK5,
             SUM(CASE WHEN FPA=6 THEN MIK_AJIA ELSE 0 END ) AS MIK6,
             SUM(CASE WHEN FPA=7 THEN MIK_AJIA ELSE 0 END ) AS MIK7
             FROM EGGTIM    WHERE ID_NUM=" + sqlDT(I)("ID_NUM").ToString, DUM10)

        'ΚΑΘΑΡΕΣ ΑΞΙΕΣ ΕΛΕΓΧΟΣ
        If DUM10(0)("KAU") = S_AJ Then 'OK   SYMFVNEI Η ΚΑΘΑΡΗ ΑΞΙΑ
            'OK


        Else ' ΚΑΛΥΠΤΩ ΤΗΝ ΔΙΑΦΟΡΑ EGGTIM - TIM ΑΝΑ ΦΠΑ
            If Math.Abs(sqlDT(I)("aj1") - DUM10(0)("KAU1")) > 0.005 Then
                Dim DUM111 As New DataTable
                Dim DIFF1 = sqlDT(I)("aj1") - DUM10(0)("KAU1")
                ExecuteSQLQuery("SELECT * FROM EGGTIM    WHERE FPA=1 AND ID_NUM=" + sqlDT(I)("ID_NUM").ToString + "ORDER BY KAU_AJIA DESC", DUM111)
                ExecuteSQLQuery("UPDATE EGGTIM SET KAU_AJIA=KAU_AJIA+" + Replace(Str(DIFF1), ",", ".") + " WHERE ID=" + DUM111(0)("ID").ToString, DUM111)
            End If
            If Math.Abs(sqlDT(I)("aj2") - DUM10(0)("KAU2")) > 0.005 Then
                Dim DUM111 As New DataTable
                Dim DIFF1 = sqlDT(I)("aj2") - DUM10(0)("KAU2")
                ExecuteSQLQuery("SELECT * FROM EGGTIM    WHERE FPA=2 AND ID_NUM=" + sqlDT(I)("ID_NUM").ToString + "ORDER BY KAU_AJIA DESC", DUM111)
                ExecuteSQLQuery("UPDATE EGGTIM SET KAU_AJIA=KAU_AJIA+" + Replace(Str(DIFF1), ",", ".") + " WHERE ID=" + DUM111(0)("ID").ToString, DUM111)
            End If
            If Math.Abs(sqlDT(I)("aj3") - DUM10(0)("KAU3")) > 0.005 Then
                Dim DUM111 As New DataTable
                Dim DIFF1 = sqlDT(I)("aj3") - DUM10(0)("KAU3")
                ExecuteSQLQuery("SELECT * FROM EGGTIM    WHERE FPA=3 AND ID_NUM=" + sqlDT(I)("ID_NUM").ToString + "ORDER BY KAU_AJIA DESC", DUM111)
                ExecuteSQLQuery("UPDATE EGGTIM SET KAU_AJIA=KAU_AJIA+" + Replace(Str(DIFF1), ",", ".") + " WHERE ID=" + DUM111(0)("ID").ToString, DUM111)
            End If
            If Math.Abs(sqlDT(I)("aj4") - DUM10(0)("KAU4")) > 0.005 Then
                Dim DUM111 As New DataTable
                Dim DIFF1 = sqlDT(I)("aj4") - DUM10(0)("KAU4")
                ExecuteSQLQuery("SELECT * FROM EGGTIM    WHERE FPA=4 AND ID_NUM=" + sqlDT(I)("ID_NUM").ToString + "ORDER BY KAU_AJIA DESC", DUM111)
                ExecuteSQLQuery("UPDATE EGGTIM SET KAU_AJIA=KAU_AJIA+" + Replace(Str(DIFF1), ",", ".") + " WHERE ID=" + DUM111(0)("ID").ToString, DUM111)
            End If
            If Math.Abs(sqlDT(I)("aj5") - DUM10(0)("KAU5")) > 0.005 Then
                Dim DUM111 As New DataTable
                Dim DIFF1 = sqlDT(I)("aj5") - DUM10(0)("KAU5")
                ExecuteSQLQuery("SELECT * FROM EGGTIM    WHERE FPA=5 AND ID_NUM=" + sqlDT(I)("ID_NUM").ToString + "ORDER BY KAU_AJIA DESC", DUM111)
                ExecuteSQLQuery("UPDATE EGGTIM SET KAU_AJIA=KAU_AJIA+" + Replace(Str(DIFF1), ",", ".") + " WHERE ID=" + DUM111(0)("ID").ToString, DUM111)
            End If
            If Math.Abs(sqlDT(I)("aj6") - DUM10(0)("KAU6")) > 0.005 Then
                Dim DUM111 As New DataTable
                Dim DIFF1 = sqlDT(I)("aj6") - DUM10(0)("KAU6")
                ExecuteSQLQuery("SELECT * FROM EGGTIM    WHERE FPA=6 AND ID_NUM=" + sqlDT(I)("ID_NUM").ToString + "ORDER BY KAU_AJIA DESC", DUM111)
                ExecuteSQLQuery("UPDATE EGGTIM SET KAU_AJIA=KAU_AJIA+" + Replace(Str(DIFF1), ",", ".") + " WHERE ID=" + DUM111(0)("ID").ToString, DUM111)
            End If
            If Math.Abs(sqlDT(I)("aj7") - DUM10(0)("KAU7")) > 0.005 Then
                Dim DUM111 As New DataTable
                Dim DIFF1 = sqlDT(I)("aj7") - DUM10(0)("KAU7")
                ExecuteSQLQuery("SELECT * FROM EGGTIM    WHERE FPA=7 AND ID_NUM=" + sqlDT(I)("ID_NUM").ToString + "ORDER BY KAU_AJIA DESC", DUM111)
                ExecuteSQLQuery("UPDATE EGGTIM SET KAU_AJIA=KAU_AJIA+" + Replace(Str(DIFF1), ",", ".") + " WHERE ID=" + DUM111(0)("ID").ToString, DUM111)
            End If
        End If

        'ΚΑΘΑΡΕΣ MIKTES ΑΞΙΕΣ ΕΛΕΓΧΟΣ
        If DUM10(0)("MIK") = S_AJ + S_FPA Then 'OK   SYMFVNEI Η ΑΞΙΑ ME FPA
            'OK


        Else ' ΚΑΛΥΠΤΩ ΤΗΝ ΔΙΑΦΟΡΑ EGGTIM - TIM ΑΝΑ ΦΠΑ
            If Math.Abs(sqlDT(I)("aj1") + sqlDT(I)("FPA1") - DUM10(0)("MIK1")) > 0.005 Then
                Dim DUM111 As New DataTable
                Dim DIFF1 = sqlDT(I)("aj1") + sqlDT(I)("FPA1") - DUM10(0)("MIK1")
                ExecuteSQLQuery("SELECT * FROM EGGTIM    WHERE FPA=1 AND ID_NUM=" + sqlDT(I)("ID_NUM").ToString + "ORDER BY KAU_AJIA DESC", DUM111)
                ExecuteSQLQuery("UPDATE EGGTIM SET MIK_AJIA=MIK_AJIA+" + Replace(Str(DIFF1), ",", ".") + " WHERE ID=" + DUM111(0)("ID").ToString, DUM111)
            End If
            If Math.Abs(sqlDT(I)("aj2") + sqlDT(I)("FPA2") - DUM10(0)("MIK2")) > 0.005 Then
                Dim DUM111 As New DataTable
                Dim DIFF1 = sqlDT(I)("aj2") + sqlDT(I)("FPA2") - DUM10(0)("MIK2")
                ExecuteSQLQuery("SELECT * FROM EGGTIM    WHERE FPA=2 AND ID_NUM=" + sqlDT(I)("ID_NUM").ToString + "ORDER BY KAU_AJIA DESC", DUM111)
                ExecuteSQLQuery("UPDATE EGGTIM SET MIK_AJIA=MIK_AJIA+" + Replace(Str(DIFF1), ",", ".") + " WHERE ID=" + DUM111(0)("ID").ToString, DUM111)
            End If
            If Math.Abs(sqlDT(I)("aj3") + sqlDT(I)("FPA3") - DUM10(0)("MIK3")) > 0.005 Then
                Dim DUM111 As New DataTable
                Dim DIFF1 = sqlDT(I)("aj3") + sqlDT(I)("FPA3") - DUM10(0)("MIK3")
                ExecuteSQLQuery("SELECT * FROM EGGTIM    WHERE FPA=3 AND ID_NUM=" + sqlDT(I)("ID_NUM").ToString + "ORDER BY KAU_AJIA DESC", DUM111)
                ExecuteSQLQuery("UPDATE EGGTIM SET MIK_AJIA=MIK_AJIA+" + Replace(Str(DIFF1), ",", ".") + " WHERE ID=" + DUM111(0)("ID").ToString, DUM111)
            End If
            If Math.Abs(sqlDT(I)("aj4") + sqlDT(I)("FPA4") - DUM10(0)("MIK4")) > 0.005 Then
                Dim DUM111 As New DataTable
                Dim DIFF1 = sqlDT(I)("aj4") + sqlDT(I)("FPA4") - DUM10(0)("MIK4")
                ExecuteSQLQuery("SELECT * FROM EGGTIM    WHERE FPA=4 AND ID_NUM=" + sqlDT(I)("ID_NUM").ToString + "ORDER BY KAU_AJIA DESC", DUM111)
                ExecuteSQLQuery("UPDATE EGGTIM SET MIK_AJIA=MIK_AJIA+" + Replace(Str(DIFF1), ",", ".") + " WHERE ID=" + DUM111(0)("ID").ToString, DUM111)
            End If
            If Math.Abs(sqlDT(I)("aj6") + sqlDT(I)("FPA6") - DUM10(0)("MIK6")) > 0.005 Then
                Dim DUM111 As New DataTable
                Dim DIFF1 = sqlDT(I)("aj6") + sqlDT(I)("FPA6") - DUM10(0)("MIK6")
                ExecuteSQLQuery("SELECT * FROM EGGTIM    WHERE FPA=6 AND ID_NUM=" + sqlDT(I)("ID_NUM").ToString + "ORDER BY KAU_AJIA DESC", DUM111)
                ExecuteSQLQuery("UPDATE EGGTIM SET MIK_AJIA=MIK_AJIA+" + Replace(Str(DIFF1), ",", ".") + " WHERE ID=" + DUM111(0)("ID").ToString, DUM111)
            End If

            If Math.Abs(sqlDT(I)("aj7") + sqlDT(I)("FPA7") - DUM10(0)("MIK7")) > 0.005 Then
                Dim DUM111 As New DataTable
                Dim DIFF1 = sqlDT(I)("aj7") + sqlDT(I)("FPA7") - DUM10(0)("MIK7")
                ExecuteSQLQuery("SELECT * FROM EGGTIM    WHERE FPA=7 AND ID_NUM=" + sqlDT(I)("ID_NUM").ToString + "ORDER BY KAU_AJIA DESC", DUM111)
                ExecuteSQLQuery("UPDATE EGGTIM SET MIK_AJIA=MIK_AJIA+" + Replace(Str(DIFF1), ",", ".") + " WHERE ID=" + DUM111(0)("ID").ToString, DUM111)
            End If

        End If

        checkIntegrity = True

    End Function

    Private Sub Test()
        Dim objWorkingXML As New System.Xml.XmlDocument
        Dim objValidateXML As System.Xml.XmlValidatingReader
        Dim objSchemasColl As New System.Xml.Schema.XmlSchemaCollection

        objSchemasColl.Add("http://www.aade.gr/myDATA/invoice/v1.0", "c:\txtfiles\invoicesDoc-v0.6.xsd")
        objValidateXML = New System.Xml.XmlValidatingReader(New System.Xml.XmlTextReader("c:\txtfiles\inv2.xml"))

        AddHandler objValidateXML.ValidationEventHandler, AddressOf ValidationCallBack
        objValidateXML.Schemas.Add(objSchemasColl)

        'This is WHERE the validation occurs.. WHEN the XML Document READS through the validating reader

        objWorkingXML.Load(objValidateXML)

        'Close the stream
        objValidateXML.Close()


        'The document is valid
        MsgBox("THE DOCUMENT IS VALID")

    End Sub



    Private Shared Sub ValidationCallBack(ByVal sender As Object, ByVal e As System.Xml.Schema.ValidationEventArgs)
        Dim test As String
        MsgBox(e.Message)
        ' Throw e.Exception

    End Sub







    Private Sub ValidationEventHandler(ByVal sender As Object, ByVal e As ValidationEventArgs)
        Select Case e.Severity
            Case XmlSeverityType.Error
                Debug.WriteLine("Error: {0}", e.Message)
                ListBox2.Items.Add("ERROR " + e.Message)
            Case XmlSeverityType.Warning
                Debug.WriteLine("Warning {0}", e.Message)
                ListBox2.Items.Add("warning " + e.Message)
        End Select
    End Sub




    Private Sub crNode(ByVal pName As String, ByVal cValue As String, ByVal writer As XmlTextWriter)
        writer.WriteStartElement(pName)
        writer.WriteString(cValue)
        writer.WriteEndElement()
    End Sub



    Private Sub EditConnString_Click(sender As Object, e As EventArgs) Handles EditConnString.Click
        If checkServer(1) Then
            MsgBox("OK")
            ExecuteSQLQuery("IF COL_LENGTH('dbo.TIM', 'AADEFPA') IS  NULL  BEGIN; ALTER TABLE TIM ADD AADEFPA FLOAT NULL;END")
            ExecuteSQLQuery("IF COL_LENGTH('dbo.TIM', 'AADEKAU') IS  NULL  BEGIN; ALTER TABLE TIM ADD AADEKAU FLOAT NULL;END")
            ExecuteSQLQuery("IF COL_LENGTH('dbo.TIM', 'ENTITY') IS  NULL  BEGIN;  ALTER TABLE TIM ADD ENTITY  INT NULL;END")
            ExecuteSQLQuery("IF COL_LENGTH('dbo.TIM', 'ENTITYUID') IS  NULL  BEGIN;  ALTER TABLE TIM ADD ENTITYUID VARCHAR(40) NULL;END")
            ExecuteSQLQuery("IF COL_LENGTH('dbo.TIM', 'ENTITYMARK') IS  NULL  BEGIN; ALTER TABLE TIM ADD ENTITYMARK VARCHAR(18) NULL;END")

        End If
    End Sub

    'Δουλεύει κανονικά, αν έχεις κάποια απορία στείλε μου email στο manol14@hotmail.com









    '  <totalNetValue>100</totalNetValue>
    '    <totalVatAmount>24</totalVatAmount>
    '    <totalWithheldAmount>0</totalWithheldAmount>
    '    <totalFeesAmount>0</totalFeesAmount>
    '    <totalStampDutyAmount>0</totalStampDutyAmount>
    '    <totalOtherTaxesAmount>0</totalOtherTaxesAmount>
    '    <totalDeductionsAmount>0</totalDeductionsAmount>
    '    <totalGrossValue>124</totalGrossValue>
    '  </invoiceSummary>
    '</invoice>


    Public Function ExecuteSQLQuery(ByVal SQLQuery As String) As DataTable
        Try
            Dim sqlCon As New OleDbConnection(gConnect)

            Dim sqlDA As New OleDbDataAdapter(SQLQuery, sqlCon)

            Dim sqlCB As New OleDbCommandBuilder(sqlDA)
            sqlDT.Reset() ' refresh 
            sqlDA.Fill(sqlDT)
            'rowsAffected = command.ExecuteNonQuery();
            ' sqlDA.Fill(sqlDaTaSet, "PEL")

        Catch ex As Exception
            MsgBox("Error: " & ex.ToString)
            If Err.Number = 5 Then
                MsgBox("Invalid Database, Configure TCP/IP", MsgBoxStyle.Exclamation, "Sales and Inventory")
            Else
                MsgBox("Error : " & ex.Message)
            End If
            ERROR_WRITE(Format(Now, "dd/MM/yy") + "Error No. " & Err.Number & " Invalid database or no database found !! Adjust settings first")
            ERROR_WRITE(SQLQuery)
            'MsgBox(SQLQuery)
            End


        End Try
        Return sqlDT
    End Function

    Public Sub ExecuteSQLQuery(ByVal SQLQuery As String, ByRef SQLDT As DataTable)
        'αν χρησιμοποιώ  byref  tote prepei να δηλωθεί   
        'Dim DTI As New DataTable


        Try
            Dim sqlCon As New OleDbConnection(gConnect)

            Dim sqlDA As New OleDbDataAdapter(SQLQuery, sqlCon)

            Dim sqlCB As New OleDbCommandBuilder(sqlDA)
            'SQLDT.Reset() ' refresh 
            sqlDA.Fill(SQLDT)
            ' sqlDA.Fill(sqlDaTaSet, "PEL")

        Catch ex As Exception
            MsgBox("Error: " & ex.ToString)
            If Err.Number = 5 Then
                MsgBox("Invalid Database, Configure TCP/IP", MsgBoxStyle.Exclamation, "Sales and Inventory")
            Else
                MsgBox("Error : " & ex.Message)
            End If
            MsgBox("Error No. " & Err.Number & " Invalid database or no database found !! Adjust settings first", MsgBoxStyle.Critical, "Sales And Inventory")
            MsgBox(SQLQuery)
            ERROR_WRITE(Format(Now, "dd/MM/yy") + "Error No. " & Err.Number & " Invalid database or no database found !! Adjust settings first")
            ERROR_WRITE(SQLQuery)
            End

        End Try
        'Return sqlDT
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Test()
    End Sub

    Private Sub UPDATE_TIM_Click(sender As Object, e As EventArgs) Handles UPDATE_TIM.Click
        UpdateTim()
    End Sub

    Private Sub UpdateTim()
        'ΠΑΙΡΝΩ ΤΟ  "c:\txtfiles\apantSendInv.XML")
        ' ΚΑΙ ΕΝΗΜΕΡΩΝΩ ΤΟ ΤΙΜ ΚΑΙ ΜΕΤΑ
        '' ΕΔΩ ΔΗΜΙΟΥΡΓΩ ΤΟ INCOME ΑΡΧΕΙΟ ΜΕ ΠΡΟΣΔΙΟΡΙΣΜΟ ΤΗΣ ΚΑΘΕ ΕΓΓΡΑΦΗΣ    "c:\txtfiles\inC.xml"





        If checkServer(0) Then
            ' MsgBox("OK")
        End If



        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load("C:\TXTFILES\apantSendInv.xml")
        Dim nodes As XmlNodeList = xmlDoc.DocumentElement.SelectNodes("/ResponseDoc/response")

        'ExecuteSQLQuery("delete from oc_category_filter")
        Me.Text = "Ελεγχος αριθμου εγγραφών"
        Application.DoEvents()

        Dim NF As Long = 0
        For Each node As XmlNode In nodes
            NF = NF + 1
        Next
        Application.DoEvents()
        Me.Text = NF


        ' ΕΔΩ ΔΗΜΙΟΥΡΓΩ ΤΟ INCOME ΑΡΧΕΙΟ ΜΕ ΠΡΟΣΔΙΟΡΙΣΜΟ ΤΗΣ ΚΑΘΕ ΕΓΓΡΑΦΗΣ

        'Dim ff As String = "c:\txtfiles\inC.xml"  'c:\mercvb\m" + Format(Now, "yyyyddmmHHMM") + ".export" ' "\\Logisthrio\333\pr.export" '
        'Dim writer As New XmlTextWriter(ff, System.Text.Encoding.UTF8)
        'writer.WriteStartDocument(True)
        'writer.Formatting = Formatting.Indented
        'writer.Indentation = 2
        'writer.WriteStartElement("IncomeClassificationsDoc")
        'writer.WriteAttributeString("xmlns", "https://www.aade.gr/myDATA/incomeClassificaton/v1.0")
        'writer.WriteAttributeString("xsi:schemaLocation", "https://www.aade.gr/myDATA/incomeClassificaton/v1.0 schema.xsd")
        'writer.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")







        Dim SuccessCounter = 0
        Dim cSuccessCounter As String = "0"
        Dim line As Integer
        Dim invoiceUid As String
        Dim invoiceMark As String
        For Each node As XmlNode In nodes
            Dim Status As String = node.SelectSingleNode("statusCode").InnerText
            Dim merror As String
            line = node.SelectSingleNode("index").InnerText

            Dim temp As New DataTable
            If Status = "Success" Then
                invoiceUid = node.SelectSingleNode("invoiceUid").InnerText
                invoiceMark = node.SelectSingleNode("invoiceMark").InnerText
                SuccessCounter = SuccessCounter + 1
                cSuccessCounter = Str(SuccessCounter)
                'ΑΝ ΕΧΕΙ ΑΠΟΣΤΑΛΕΙ ΤΟ ΤΙΜΟΛΟΓΙΟ ΜΕ ΕΠΙΤΥΧΙΑ ΠΑΙΡΝΩ ΤΗΝ ΕΥΚΑΙΡΙΑ
                'ΝΑ ΣΤΕΙΛΩ ΚΑΙ ΤΟΝ ΤΥΠΟ ΤΟΥ ΕΣΟΔΟΥ

                ExecuteSQLQuery("select AADEKAU,ID_NUM,ATIM,HME FROM TIM   WHERE ENTITY=" + Str(line), temp)

                Dim EGGTIM As New DataTable
                ExecuteSQLQuery("select round(POSO*TIMM*(100-EKPT)/100,2) AS AJ FROM EGGTIM   WHERE POSO*TIMM<>0 AND ID_NUM=" + Str(temp(0)(1)), EGGTIM)

                ' writer.WriteComment(temp(0)("ATIM") + " " + Format(temp(0)("HME"), "dd/MM/yyyy"))

                ' writer.WriteStartElement("incomeInvoiceClassification") '---------------------------
                'crNode("invoiceMark", invoiceMark, writer)

                For L As Integer = 0 To EGGTIM.Rows.Count - 1

                    '   writer.WriteStartElement("invoicesIncomeClassificationDetails") '====
                    '  crNode("lineNumber", Str(L + 1), writer)
                    ' writer.WriteStartElement("incomeClassificationDetailData") '***
                    ' crNode("classificationType", "101", writer)
                    ' crNode("classificationCategory", "1", writer)
                    ' crNode("amount", Format(EGGTIM(L)(0), "#####0.#####"), writer)
                    ' writer.WriteEndElement() 'incomeClassificationDetailData    '****
                    ' writer.WriteEndElement() ' invoicesIncomeClassificationDetails   ======

                Next
                'writer.WriteEndElement() ' incomeInvoiceClassification---------------------

            Else 'ΕΧΕΙ ΛΑΘΟΣ ΟΠΟΤΕ ΑΠΟΘΗΚΕΥΩ ΤΟ ΛΑΘΟΣ ΣΤΟ ΤΙΜ.invoiceUid
                invoiceUid = node.SelectSingleNode("errors/error/message").InnerText
                invoiceMark = "ERROR"
                cSuccessCounter = "0"
                MsgBox("ΛΑΘΟΣ ΣΤΗΝ ΑΠΟΣΤΟΛΗ ΣΤΟ ΤΙΜΟΛΟΓΙΟ " + temp(0)(1)("ATIM"))
            End If

            ExecuteSQLQuery("update TIM SET ENTITYUID='" + Mid(invoiceUid, 1, 40) + "' , ENTITYMARK='" + Mid(invoiceMark, 1, 16) + "',ENTLINEN=" + cSuccessCounter + "  WHERE ENTITY=" + Str(line))


        Next

        'writer.WriteEndDocument()
        'writer.Close()

    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged

    End Sub

    Public Sub paint_ergasies(GridView1 As DataGridView, sqlQuery As String)



        Dim connectionString As String = gSQLCon '"Data Source=.;Initial Catalog=pubs;Integrated Security=True"
        Dim sql As String = sqlQuery ' "SELECT * FROM TIM"
        Dim connection As New SqlConnection(connectionString)
        Dim dataadapter As New SqlDataAdapter(sql, connection)
        Dim ds As New DataSet()
        connection.Open()
        dataadapter.Fill(ds, "Authors_table")
        connection.Close()
        GridView1.DataSource = ds
        GridView1.DataMember = "Authors_table"










        'Dim con As New OleDb.OleDbConnection
        'con.ConnectionString = gConnect
        'con.Open()
        'Dim ds As DataSet = New DataSet
        'Dim adapter As New OleDb.OleDbDataAdapter
        'Dim sql As String

        'sql = sqlQuery

        'adapter.SelectCommand = New OleDb.OleDbCommand(sql, con)
        'adapter.Fill(ds)
        'GridView1.DataSource = ds.Tables(0)


































        Exit Sub

        ''Create connection
        'Dim conn As SqlConnection

        ''create data adapter
        'Dim da As New SqlDataAdapter

        ''create dataset
        'Dim ds As DataSet = New DataSet

        'conn = New SqlConnection(gSQLCon)
        'Try
        '    ' Open connection
        '    conn.Open()

        '    da = New SqlDataAdapter(sqlQuery, conn)

        '    'create command builder
        '    Dim cb As SqlCommandBuilder = New SqlCommandBuilder(da)
        '    ds.Clear()
        '    'fill dataset
        '    da.Fill(ds, "PEL")
        '    GridView1.ClearSelection()
        '    GridView1.DataSource = ds
        '    GridView1.DataMember = "PEL"


        '    GridView1.Refresh()







        'Catch ex As SqlException
        '    MsgBox(ex.ToString)
        'Finally
        '    ' Close connection
        '    conn.Close()
        'End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If checkServer(0) Then

        End If
        ' paint_ergasies(DataGridView1, "Select ATIM, HME, ENTITY, AADEKAU, AJ1 + AJ2 + AJ3 + AJ4 + AJ5 + AJ6 + AJ7 As KAUTIM, AADEFPA, FPA1 + FPA2 + FPA3 + FPA4 + FPA6 + FPA7 As FPATIM, ENTITYUID, ENTITYMARK FROM TIM WHERE ENTITY>0")
    End Sub

    ' Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
    Private Async Sub MakeRequest2()
        Dim client = New HttpClient()
        Dim queryString = HttpUtility.ParseQueryString(String.Empty)
        client.DefaultRequestHeaders.Add("aade-user-id", gUserId)
        client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", gSubKey)
        Dim cc As String = InputBox("ΑΠΟ ΠΟΙΟ ΜΑΡΚ ΚΑΙ ΜΕΤΑ;")
        queryString("mark") = cc  ' "1000000006337" ' "{string}"
        '  queryString("nextPartitionKey") = "{string}"
        '   queryString("nextRowKey") = "{string}"
        Dim uri = "https://mydata-dev.azure-api.net/RequestDocs?" & queryString.ToString

        Dim response = Await client.GetAsync(uri)
        Dim result = Await response.Content.ReadAsStringAsync()
        TextBox2.Text = result.ToString

        Dim MF = "c:\txtfiles\apantReqtome.xml"  'Inv" + Format(Now, "yyyyddMMHHmm") + ".xml"
        FileOpen(1, MF, OpenMode.Output)
        PrintLine(1, result.ToString)
        FileClose(1)






    End Sub

    Private Async Sub Make2RequestKTEL()  ' LOGI
        Dim client = New HttpClient()
        Dim queryString = HttpUtility.ParseQueryString(String.Empty)

        ' AYTO XREIAZETAI GIA TA TICKETS
        'client.DefaultRequestHeaders.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", "YOUR_ACCESS_TOKEN")



        'ΟΚ   Dim jsonContent As String = " {""syncAll"":0, ""syncTables"": [1,2,3,8,9,10,11,12],""devicesUuid"":""ERP"",""lastSyncDate"":null,""fetchAllUsers"" : true }"
        Dim jsonContent As String = " {""username"":""0"", ""password"": ""==="",""devicesUuid"":""ERP"" }"

        ' Dim jsonContent As String = "{""syncAll"":""0"", ""syncTables"": ""[1,2,3,8,9,10,11,12"",""devicesUuid"":""ERP"",""lastSyncDate"":""null"",""fetchAllUsers"" : ""true"" }"

        Dim content As New StringContent(jsonContent, Encoding.UTF8, "application/json")

        'client.DefaultRequestHeaders.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", "YOUR_ACCESS_TOKEN")









        ' Dim cc As String = InputBox("ΑΠΟ ΠΟΙΟ ΜΑΡΚ ΚΑΙ ΜΕΤΑ;")
        'queryString("mark") = cc  ' "1000000006337" ' "{string}"
        '  queryString("nextPartitionKey") = "{string}"
        '   queryString("nextRowKey") = "{string}"

        Dim uri = "https://drama.amfare.busnet.gr/amfare-admin/web/api/sync_bo"

        Dim response = Await client.PostAsync(uri, content)
        Dim result = Await response.Content.ReadAsStringAsync()
        TextBox2.Text = result.ToString

        Dim MF = "c:\txtfiles\apantReqtome.xml"  'Inv" + Format(Now, "yyyyddMMHHmm") + ".xml"
        FileOpen(1, MF, OpenMode.Output)
        PrintLine(1, result.ToString)
        FileClose(1)






    End Sub

    Private Async Sub Make3RequestKTELtickets()  ' LOGIN USER

        Dim client = New HttpClient()
        Dim queryString = HttpUtility.ParseQueryString(String.Empty)

        ' AYTO XREIAZETAI GIA TA TICKETS
        client.DefaultRequestHeaders.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", "MiJ9.eyJzdWIiOiJhbWNvRXJwVXNlciIsInVzZXJJZCI6IjM1NyIsImNvbW1hU2VwYXJhdGVkQXV0aG9yaXRpZXMiOiJERUxfSElTVE9SWV9SRUNPTkNJTElBVElPTixERUxfUkVDT05DSUxJQVRJT04sRURJVF9ISVNUT1JZX1JFQ09OQ0lMSUFUSU9OLEVESVRfUkVDT05DSUxJQVRJT04sSE9NRSxORVdfSElTVE9SWV9SRUNPTkNJTElBVElPTixORVdfUkVDT05DSUxJQVRJT04sT0JVX0RSSVZFUixWSUVXX0hJU1RPUllfUkVDT05DSUxJQVRJT04sVklFV19SRUNPTkNJTElBVElPTixWSUVXX1VTRVJTLFZJRVdfWl9FVkVOVFMiLCJleHAiOjE3MDYwMjY5NjV9.S6rbSTKpF397NsVdVLfQXJWKZNauh6omo-pyEcmh7hvjgR90UNWX-6bRez_Dj4xetfpwpLMREOJAxbGqObc9eA")



        'ΟΚ   Dim jsonContent As String = " {""syncAll"":0, ""syncTables"": [1,2,3,8,9,10,11,12],""devicesUuid"":""ERP"",""lastSyncDate"":null,""fetchAllUsers"" : true }"
        '""deviceUuid"":""ERP"",  

        Dim jsonContent As String = "{  ""carrierId"":108, ""carrierIdentifier"": null,""language"":""el"",""transactionTimestampFrom"":""07/06/2024 00:00:00"",""transactionTimestampTo"" : ""07/06/2024 22:00:00"", ""userId"": [] }"

        Dim content As New StringContent(jsonContent, Encoding.UTF8, "application/json")

        'client.DefaultRequestHeaders.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", "YOUR_ACCESS_TOKEN")

        '        Select Case U.lastname,F.code ,S.revenue/100 As TIMH,* 
        ' From [EISITIRIA] S 
        ' INNER Join Users U on S.userId=U.id 
        ' INNER Join FareProducts F on S.fareProductId=F.id 
        ' where userId = 9







        ' Dim cc As String = InputBox("ΑΠΟ ΠΟΙΟ ΜΑΡΚ ΚΑΙ ΜΕΤΑ;")
        'queryString("mark") = cc  ' "1000000006337" ' "{string}"
        '  queryString("nextPartitionKey") = "{string}"
        '   queryString("nextRowKey") = "{string}"

        Dim uri = "https://drama.amfare.busnet.gr/amfare-admin/web/pub/logistics/get_ticket_transactions" ' "https://drama.amfare.busnet.gr/amfare-admin/web/api/sync_bo"

        Dim response = Await client.PostAsync(uri, content)
        Dim result = Await response.Content.ReadAsStringAsync()
        TextBox2.Text = result.ToString

        Dim MF = "c:\txtfiles\ticketsapantReqtome.xml"  'Inv" + Format(Now, "yyyyddMMHHmm") + ".xml"
        FileOpen(1, MF, OpenMode.Output)
        PrintLine(1, result.ToString)
        FileClose(1)
        Dim json As String = result.ToString
        Dim ser As JObject = JObject.Parse(json)
        Dim c As String = ""
        Dim SQLDT As New DataTable
        Execute2SQLQuery("delete from EISITIRIA", SQLDT)
        For K As Integer = 0 To ser("transactions").Count - 1


            Dim userId As String = ser("transactions")(K)("userId").ToString
            Dim quantity As String = ser("transactions")(K)("quantity").ToString
            Dim fareProductId As String = ser("transactions")(K)("fareProductId").ToString
            Dim paymentMethodId As String = ser("transactions")(K)("paymentMethodId").ToString
            'cissueDateTimestamp
            Dim cissueDateTimestamp As String = ser("transactions")(K)("issueDateTimestamp").ToString
            Dim deviceUuid As String = ser("transactions")(K)("deviceUuid").ToString
            Dim userUuid As String = ser("transactions")(K)("userUuid").ToString

            Dim revenue As String = ser("transactions")(K)("revenue").ToString
            Dim totalRevenue As String = ser("transactions")(K)("totalRevenue").ToString
            '
            Execute2SQLQuery("INSERT INTO EISITIRIA(userUuid, deviceUuid,cissueDateTimestamp,userId,quantity,fareProductId,paymentMethodId,revenue,totalRevenue) VALUES ('" + userUuid + "','" + deviceUuid + "','" + cissueDateTimestamp + "'," + userId + "," + quantity + "," + fareProductId + "," + paymentMethodId + "," + revenue + "," + totalRevenue + ")", SQLDT)
        Next

        MsgBox("OK " + ser("transactions").Count.ToString)

        'xreiazomai hmeromhnia ,kvdiko eidoys,kodiko pelTH,ΑΡΙΘΜΟ ΠΑΡΑΣΤΑΤΙΚΟΥ
        'insert into EGGTIM (HME,ATIM,EIDOS,TIMM,PELKOD,KODE,POSO)   select '03/26/2024' AS HME,'T000122' AS ATIM,'e' as e,SUM(revenue/100) as r,userId,fareProductId,quantity from EISITIRIA where userId=9 group by userId,fareProductId,quantity
        ' insert into TIM (HME,ATIM,EIDOS,AJI,KPE,KLEIDI)  select '03/26/2024' AS HME,'T000122' AS ATIM,'e' as e,SUM(revenue/100) as r,'USER' AS AA,'T000122' AS ATIM2 from EISITIRIA where userId=9 



    End Sub

    Sub Execute2SQLQuery(ByVal SQLQuery As String, ByRef SQLDT As DataTable)
        'αν χρησιμοποιώ  byref  tote prepei να δηλωθεί   
        'Dim DTI As New DataTable
        ' Dim gConnect2 As String = "Provider=SQLOLEDB.1;Data Source=LOGISTIRIO\SQLEXPRESS;Integrated Security=True;database=MERCURY"
        Dim cn2String As String = "Provider=SQLOLEDB.1;;Password=p@ssw0rd;Persist Security Info=True ;User Id=sa;Initial Catalog=MERCURY;Data Source=LOGISTIRIO\SQLEXPRESS"
        Try
            Dim sqlCon As New OleDbConnection(cn2String)

            Dim sqlDA As New OleDbDataAdapter(SQLQuery, sqlCon)

            Dim sqlCB As New OleDbCommandBuilder(sqlDA)
            SQLDT.Reset() ' refresh 
            sqlDA.Fill(SQLDT)
            ' sqlDA.Fill(sqlDaTaSet, "PEL")

        Catch 'ex As Exception
            'MsgBox("Error: " & ex.ToString)
            'If Err.Number = 5 Then
            '    MsgBox("Invalid Database, Configure TCP/IP", MsgBoxStyle.Exclamation, "Sales and Inventory")
            'Else
            '    MsgBox("Error : " & ex.Message)
            'End If
            'MsgBox("Error No. " & Err.Number & " Invalid database or no database found !! Adjust settings first", MsgBoxStyle.Critical, "Sales And Inventory")
            MsgBox(Err.Description + Chr(13) + SQLQuery)
        End Try
        'Return sqlDT
    End Sub



    Private Async Sub MakeRequestKTEL()
        Dim client = New HttpClient()
        Dim queryString = HttpUtility.ParseQueryString(String.Empty)

        ' AYTO XREIAZETAI GIA TA TICKETS
        'client.DefaultRequestHeaders.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", "YOUR_ACCESS_TOKEN")

        Dim JSONresponseFromServer As String

        'ΟΚ   Dim jsonContent As String = " {""syncAll"":0, ""syncTables"": [1,2,3,8,9,10,11,12],""devicesUuid"":""ERP"",""lastSyncDate"":null,""fetchAllUsers"" : true }"
        Dim jsonContent As String = " {""syncAll"":0, ""syncTables"": [1,2,3,8,9,10,11,12,13],""devicesUuid"":""ERP"",""lastSyncDate"":null,""fetchAllUsers"" : true }"

        ' Dim jsonContent As String = "{""syncAll"":""0"", ""syncTables"": ""[1,2,3,8,9,10,11,12"",""devicesUuid"":""ERP"",""lastSyncDate"":""null"",""fetchAllUsers"" : ""true"" }"

        Dim content As New StringContent(jsonContent, Encoding.UTF8, "application/json")

        'client.DefaultRequestHeaders.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", "YOUR_ACCESS_TOKEN")









        ' Dim cc As String = InputBox("ΑΠΟ ΠΟΙΟ ΜΑΡΚ ΚΑΙ ΜΕΤΑ;")
        'queryString("mark") = cc  ' "1000000006337" ' "{string}"
        '  queryString("nextPartitionKey") = "{string}"
        '   queryString("nextRowKey") = "{string}"

        Dim uri = "https://drama.amfare.busnet.gr/amfare-admin/web/api/sync_bo"

        Dim response = Await client.PostAsync(uri, content)
        Dim result = Await response.Content.ReadAsStringAsync()
        TextBox2.Text = result.ToString
        JSONresponseFromServer = result.ToString
        Dim MF = "c:\txtfiles\apantReqtome.xml"  'Inv" + Format(Now, "yyyyddMMHHmm") + ".xml"
        FileOpen(1, MF, OpenMode.Output)
        PrintLine(1, result.ToString)
        FileClose(1)

        Dim json As String = JSONresponseFromServer
        Dim ser As JObject = JObject.Parse(json)

        Dim c As String = ""
        Dim SQLDT As New DataTable
        Execute2SQLQuery("delete from FareProducts", SQLDT)
        For K As Integer = 0 To ser("FareProducts").Count - 1


            Dim id As String = ser("FareProducts")(K)("id").ToString
            Dim code As String = ser("FareProducts")(K)("code").ToString



            Execute2SQLQuery("INSERT INTO FareProducts( id,code) VALUES (" + id + ",'" + code + "' )", SQLDT)
        Next


        Execute2SQLQuery("delete from PaymentMethods", SQLDT)
        For K As Integer = 0 To ser("PaymentMethods").Count - 1


            Dim id As String = ser("PaymentMethods")(K)("id").ToString
            Dim code As String = ser("PaymentMethods")(K)("code").ToString


            Execute2SQLQuery("INSERT INTO PaymentMethods( id,code) VALUES (" + id + ",'" + code + "' )", SQLDT)
        Next


        '       "firstname" [nvarchar](50) NULL,
        '"lastname" [nvarchar](50) NULL,
        '"username" [nvarchar](50) NULL,
        '"uuid" [nvarchar](50) NULL,
        '"email" [nvarchar](50) NULL,

        '{
        '    "id": 37,
        '    "username": "077809",
        '    "password": "$2a$10$4fhrrehADIFCbpsUD9bC7eNU2yyawbUslk2IoOyac0emZpObTBSqi",
        '    "salt": null,
        '    "confirmPassword": null,
        '    "email":  "user077809@nomail.com",
        '    "uuid": "1f247151-5046-4b55-9291-2d40d072f81c",
        '    "mobile": "69077809",
        '    "firstname": "ΣΤΥΛΙΑΝΟΣ",
        '    "lastname": "ΧΑΤΖΗΙΩΑΝΝΙΔΗΣ",
        '    "status": "ACTIVE",
        '    "userRoleList": [
        '        {
        '            "userId": 37,
        '            "roleId": 20,
        '            "enabled": true,
        '            "roleName": "DRIVER"
        '        }
        '    ],
        '    "statuses": null,
        '    "userRoleStrToDisplay": null,
        '    "isAdmin":  false,
        '    "carriersId": 108,
        '    "thirdPartiesId": null,
        '    "thirdPartyUserId": null,
        '    "enabled":  "1",
        '    "activationUrlKey": null,
        '    "activationPin": null,
        '    "rfId":  "",
        '    "isSuperAdmin": 0,
        '    "litId": 0,
        '    "fathersName": "",
        '    "logisticsCodeAxia": "",
        '    "logisticsCodeAxiaParcel": "",
        '    "logisticsCodeSum": "",
        '    "logisticsCodeFpa": "",
        '    "logisticsCodeFpaParcel": "",
        '    "logisticsCodePayroll": "",
        '    "isDriver": true,
        '    "isPosUser": false,
        '    "maxZAmount": null,
        '    "debt": null,
        '    "serialNumber": null,
        '    "exceededAllowedDebt":  false,
        '    "exceededMaxZAmount": false,
        '    "relatedDeviceId": null,
        '    "admin":  false,
        '    "logisticsCodeCommission": ""
        '},







        Execute2SQLQuery("delete from Users", SQLDT)
        For K As Integer = 0 To ser("Users").Count - 1


            Dim firstname As String = ser("Users")(K)("firstname").ToString

            Dim id As String = ser("Users")(K)("id").ToString
            Dim lastname As String = ser("Users")(K)("lastname").ToString
            Dim username As String = ser("Users")(K)("username").ToString
            Dim uuid As String = ser("Users")(K)("uuid").ToString
            Dim email As String = ser("Users")(K)("email").ToString

            Execute2SQLQuery("INSERT INTO Users( id,firstname,lastname,username,uuid,email) VALUES (" + id + ",'" + firstname + "','" + lastname + "','" + username + "','" + uuid + "','" + email + "' )", SQLDT)
        Next







        MsgBox("ok")










        ' Dim data As List(Of JToken) = ser.Children().ToList

        'For Each item As JProperty In data
        '    item.CreateReader()
        '    Select Case item.Name
        '        Case "Stops" 'each record is inside the entries array
        '            For Each Entry As JObject In item.Values
        '                Dim naming As String = Entry("userIdentifier").ToList.Item(0)
        '                Dim factor As String = Entry("name").ToList.Item(0)
        '                ' you can continue listing the array items untill you reach the end of you array
        '                ListBox1.Items.Add(naming + "  " + factor)


        '            Next
        '    End Select
        'Next









        'If the Then Json Object item Is Not In an array And you just want the items it returns

        'For Each item As JProperty In data
        '    item.CreateReader()
        '    Dim customerId As String = ser("newsyncdate")
        '    Dim deviceIdAs As String = ser("resp")
        '    ListBox1.Items.Add(customerId + "==" + deviceIdAs)
        'Next
        '***************************  ok ******************************************

        '        ? ser("Stops")(0)("multiLanguageFieldsMap")("STOP_NAME")("en").ToString()
        '"1st STOP MONASTIRAKI"

        '>?  ser("Stops")(22)("name").ToString()
        '"ΑΓΙΑ ΕΙΡΗΝΗ"


        '>?  ser("Stops").count
        '531
        '>>?  ser("Stops")(22).count
        '10
        '>
        '        {
        '            "multiLanguageFieldsMap": {
        '                "STOP_NAME": {
        '                    "el": "1η ΣΤΑΣΗ ΜΟΝΑΣΤΗΡΑΚΙ",
        '                    "en": "1st STOP MONASTIRAKI"
        '                }
        '            },
        '            "id": 2,
        '            "longitude": "24.1623965",
        '            "latitude": "41.1842568",
        '            "carriersId": 108,
        '            "userIdentifier": "02610",
        '            "name": "1η ΣΤΑΣΗ ΜΟΝΑΣΤΗΡΑΚΙ",
        '            "address": null,
        '            "phone": null,
        '            "direction":  "BOTH"
        '        },


        '>?  ser("Stops")(22)("longitude").Tostring
        '"24.199277"









    End Sub


    '        Using System;
    'Using System.Net.Http.Headers;
    'Using System.Text;
    'Using System.Net.Http;
    'Using System.Web;

    'Namespace CSHttpClientSample
    '{
    '    Static Class Program
    '    {
    '        Static void Main()
    '        {
    '            MakeRequest();
    '            Console.WriteLine("Hit ENTER to exit...");
    '            Console.ReadLine();
    '        }

    '        Static Async void MakeRequest()
    '        {
    '            var client = New HttpClient();
    '            var queryString = HttpUtility.ParseQueryString(String.Empty);

    '            // Request headers
    '            client.DefaultRequestHeaders.Add("aade-user-id", "");
    '            client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", "{subscription key}");

    '            // Request parameters
    '            queryString["mark"] = "{string}";
    '            queryString["nextPartitionKey"] = "{string}";
    '            queryString["nextRowKey"] = "{string}";
    '            var uri = "https://mydata-dev.azure-api.net/RequestIssuerInvoices?" + queryString;

    '            var response = await client.GetAsync(uri);
    '        }
    '    }
    '}	
    '    End Sub


    '  Imports System
    '  Imports System.Net.Http.Headers
    '  Imports System.Text
    '  Imports System.Net.Http
    '  Imports System.Web'
    'Namespace CSHttpClientSample
    '   Module' Program
    Private Sub Main2()
        MakeRequest()
        Console.WriteLine("Hit ENTER to exit...")
        Console.ReadLine()
    End Sub

    'Private Async Sub MakeRequest3()
    '    Dim client = New HttpClient()
    '    Dim queryString = HttpUtility.ParseQueryString(String.Empty)
    '    client.DefaultRequestHeaders.Add("aade-user-id", "")
    '    client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", "{subscription key}")
    '    queryString("mark") = "{string}"
    '    Dim uri = "https://mydata-dev.azure-api.net/RequestInvoices?" & queryString
    '    Dim response = Await client.GetAsync(uri)
    'End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        MakeRequest2()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        UpdApantIncome()
    End Sub

    Private Sub UpdApantIncome()
        '======================================================================================
        'διαβαζω το apantIncome.xml για να δω τελικά ποια έχουν πρόβλημα 
        ' και να αποθηκεύσω το αποτέλεσμα στο ΤΙΜ
        Dim cc(1000) As String


        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load("C:\TXTFILES\apantiNCOMe.xml")



        Dim nodes As XmlNodeList = xmlDoc.DocumentElement.SelectNodes("/ResponseDoc/response")

        'ExecuteSQLQuery("delete from oc_category_filter")
        Me.Text = "Ελεγχος αριθμου εγγραφών"
        Application.DoEvents()



        Dim k As Integer = 0
        For Each node As XmlNode In nodes
            k = k + 1
            Dim mEn As String
            Dim Status As String = node.SelectSingleNode("statusCode").InnerText
            cc(k) = Status

            Dim line As String = node.SelectSingleNode("entitylineNumber").InnerText
            Dim entityMark As String


            If Status = "Success" Then
                'OK
                entityMark = node.SelectSingleNode("entityMark").InnerText
            Else
                entityMark = "ERROR"
            End If

            ExecuteSQLQuery("update TIM SET INCMARK='" + entityMark + "' WHERE ENTLINEN=" + line)

        Next




        ' ΔΕΝ ΔΙΑΒΑΖΩ ΤΟ INC.XML ΓΙΑΤΙ ΤΑ ATTRIBUTES ΕΜΠΟΔΙΖΟΥΝ ΤΟ ΔΙΑΒΑΣΜΑ ΤΩΝ NODES
        ' ΕΝΩ ΑΝ ΣΒΗΣΩ ΤΑ ATTRIBUTES ΔΙΠΛΑ ΑΠΟ ΤΟ IncomeClassificationsDoc ΔΟΥΛΕΥΕΙ ΚΑΝΟΝΙΚΑ




    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        File.Delete("C:\TXTFILES\INV.XML")
        File.Delete("C:\TXTFILES\INC.XML")


        Dim I As Integer = ToXMLsub()
        If I = 0 Then
            MsgBox("ΔΕΝ ΕΓΙΝΕ ΑΠΟΣΤΟΛΗ")
            Exit Sub
        End If
        Threading.Thread.Sleep(5000)
        MsgBox("1.ΑΠΕΣΤΑΛΗΣΑΝ ΤΑ ΑΡΧΕΙΑ")
        MakeRequest()
        Threading.Thread.Sleep(5000)
        MsgBox("2.ΑΠΕΣΤΑΛΗΣΑΝ ΤΑ ΑΡΧΕΙΑ")
        UpdateTim()
        MsgBox("3.ΕΝΗΜΕΡΩΘΗΚΕ Η ΒΑΣΗ")


        FileOpen(1, "DATES.TXT", OpenMode.Output)
        PrintLine(1, Format(APO.Value, "dd/MM/yyyy"))
        PrintLine(1, Format(EOS.Value, "dd/MM/yyyy"))

        FileClose(1)


        'Threading.Thread.Sleep(5000)
        'MakeIncomeRequest()
        'MsgBox("3.ΑΠΕΣΤΑΛΗΣΑΝ ΤΑ ΑΡΧΕΙΑ")
        'Threading.Thread.Sleep(5000)
        'UpdApantIncome()
        'MsgBox("4.ΑΠΕΣΤΑΛΗΣΑΝ ΤΑ ΑΡΧΕΙΑ")
    End Sub

    Private Sub CancInv_Click(sender As Object, e As EventArgs) Handles CancInv.Click
        CancelInvoice()
    End Sub





    Private Async Sub CancelInvoice()
        'Dim client = New HttpClient()
        'Dim queryString = HttpUtility.ParseQueryString(String.Empty)
        'client.DefaultRequestHeaders.Add("aade-user-id", "glagakis2")
        'client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", "555bc57c80634243958f62b629316aaa")
        'queryString("mark") = "400000020235191" ' "{string}"
        ''  queryString("nextPartitionKey") = "{string}"
        ''   queryString("nextRowKey") = "{string}"
        '' Dim uri = "https://mydata-dev.azure-api.net/RequestIssuerInvoices?" & queryString.ToString
        'Dim uri As String = "https://mydata-dev.azure-api.net/CancelInvoice?" & queryString.ToString
        'Dim response = Await client.PostAsync(uri, "")
        'Dim result = Await response.Content.ReadAsStringAsync()
        'TextBox2.Text = result.ToString

        'Dim MF = "c:\txtfiles\apantCancInv" + Format(Now, "yyyyddMMHHmm") + ".xml"
        'FileOpen(1, MF, OpenMode.Output)
        'PrintLine(1, result.ToString)
        'FileClose(1)



        Dim client = New HttpClient()
        'Dim queryString = HttpUtility.ParseQueryString(String.Empty)
        Try

            client.DefaultRequestHeaders.Add("aade-user-id", gUserId)
            client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", gSubKey)
            ' Dim cc As String = InputBox("ΑΠΟ ΠΟΙΟ ΜΑΡΚ ΚΑΙ ΜΕΤΑ;")
            ' queryString("mark") = cc  ' "1000000006337" ' "{string}"


            'client.DefaultRequestHeaders.Add("aade-user-id", "glagakis2")
            'client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", "555bc57c80634243958f62b629316aaa")

            Dim uri = "https://mydata-dev.azure-api.net/CancelInvoice?mark=400000020235194"   ' 400000020235191" ' + queryString.ToString

            Dim response As HttpResponseMessage
            Dim xl = XDocument.Load("c:\txtfiles\canc.xml").ToString ' "--> εκει έχω αποθηκεύσει το xml που εφτιαξα"
            Dim byteData As Byte() = Encoding.UTF8.GetBytes(xl)

            Using content = New ByteArrayContent(byteData)
                content.Headers.ContentType = New MediaTypeHeaderValue("application/xml")
                response = Await client.PostAsync(uri, content)
                Dim result = Await response.Content.ReadAsStringAsync()
                Dim MF = "c:\txtfiles\apantSendInv" + Format(Now, "yyyyddMMHHmm") + ".xml"
                FileOpen(1, MF, OpenMode.Output)
                PrintLine(1, result.ToString)
                FileClose(1)
                TextBox2.Text = result.ToString
                ' "είναι το textbox πανω στη φόρμα που σου επιστρέφει το response xml"
                'Dim byteData2 As Byte() = File.ReadAllBytes("c:\txtfiles\inv.xml")
                ' sept 2020 debug   Rename("c:\txtfiles\inv.xml", "c:\txtfiles\inv" + Format(Now, "yyyyddMMHHmm") + ".xml")
                FileCopy(MF, "c:\txtfiles\apantSendInv.XML")

            End Using


        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try




    End Sub

    Private Sub RequestTransmittedDocs_Click(sender As Object, e As EventArgs) Handles RequestTransmittedDocs.Click

        '

        MakeRequest3()


    End Sub

    Private Async Sub MakeRequest3() ' RequestTransmittedDocs_Click  400000074460885
        Dim client = New HttpClient()
        Dim queryString = HttpUtility.ParseQueryString(String.Empty)

        client.DefaultRequestHeaders.Add("aade-user-id", gUserId)
        client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", gSubKey)
        Dim cc As String = InputBox("ΑΠΟ ΠΟΙΟ ΜΑΡΚ ΚΑΙ ΜΕΤΑ;")
        queryString("mark") = cc  ' "1000000006337" ' "{string}"



        ' client.DefaultRequestHeaders.Add("aade-user-id", "glagakis2")
        ' client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", "555bc57c80634243958f62b629316aaa")
        ' queryString("mark") = "1000000006337" ' "{string}"
        '  queryString("nextPartitionKey") = "{string}"
        '   queryString("nextRowKey") = "{string}"
        Dim uri = "https://mydata-dev.azure-api.net/RequestTransmittedDocs?" & queryString.ToString
        Dim response = Await client.GetAsync(uri)
        Dim result = Await response.Content.ReadAsStringAsync()
        TextBox2.Text = result.ToString
        Dim MF = "c:\txtfiles\apantReqInv2" + Format(Now, "yyyyddMMHHmm") + ".xml"
        FileOpen(1, MF, OpenMode.Output)
        PrintLine(1, result.ToString)
        FileClose(1)
    End Sub


    Private Sub ERROR_WRITE(CCC As String)
        Dim objStreamWriter As StreamWriter
        'Pass the file path and the file name to the StreamWriter constructor.
        Dim C As String
        objStreamWriter = New StreamWriter("C:\TXTFILES\ERRORS.TXT", True, System.Text.Encoding.Default)
        objStreamWriter.WriteLine(CCC)
        objStreamWriter.Close()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        xmlread()
    End Sub
    Private Sub xmlread()






        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load("c:\txtfiles\apantreq.xml")  'xmlDoc.ChildNodes.Count =>2
        Dim nodes As XmlNodeList

        'xmlDoc.Load(fs);
        nodes = xmlDoc.GetElementsByTagName("invoice")

        ' Dim nodes As XmlNodeList = xmlDoc.DocumentElement.SelectSingleNode  '/RequestedDoc
        'Dim nodes2 As XmlNodeList = xmlDoc.DocumentElement.SelectNodes("/invoicesDoc")  '/RequestedDoc οκ
        ' Dim nodes3 As XmlNodeList = xmlDoc.DocumentElement.SelectNodes("./RequestedDoc")  '/RequestedDoc

        ' xmlDoc.DocumentElement.SelectSingleNode()

        Me.Text = "Ελεγχος αριθμου εγγραφών"
        Application.DoEvents()

        Dim NF As Long = 0
        For Each node As XmlNode In nodes
            NF = NF + 1
        Next
        Application.DoEvents()
        Me.Text = NF

        For Each node As XmlNode In nodes
            'Dim yparxei As Boolean = False

            'Dim category_id As String = node.ParentNode.SelectSingleNode("./issuer").InnerText
            ' Dim c As String = node.ChildNodes.Item(1).InnerText
            Dim c255 As String = node.ChildNodes.Item("issuer").ChildNodes.Item("vatNumber").InnerText.ToString

            Dim c2 As String = node.ChildNodes.Item(1).Name.ToString 'mark
            Dim c22 As String = node.ChildNodes.Item(2).ChildNodes(0).Name.ToString 'vatnumber (issuer)

            Dim c3 As String = node.ChildNodes.Item(3).Name.ToString
            Dim v As String = ""
            '? node.ChildNodes.Item(4).name  =>"invoiceHeader"

            'Dim pID As String = node.SelectSingleNode("id").InnerText
            ' Dim c2 As String = node.ChildNodes.Item("uid").InnerText
            'Dim parent_id As String = node.SelectSingleNode("group/category/id").InnerText
            'Dim pName As String = node.SelectSingleNode("name").InnerText

            'Dim mPrice As String = node.SelectSingleNode("price").InnerText

            'Dim pImage As String = node.SelectSingleNode("image").InnerText
            'Dim pDescription As String = node.SelectSingleNode("descr").InnerText

            'Dim pSKU As String = node.SelectSingleNode("sku").InnerText
            'Dim pPrice As String = node.SelectSingleNode("group").InnerText
            ''MessageBox.Show(pID & " " & pName & " " & pPrice)
            'Dim pNow As String = Format(Now, "yyyy-MM-dd")
            ''βρισκω το αντιστοιχο id στο δικό μου site

            ''If category_id = 118 Then
            ''    yparxei = False ' dum

            ''End If
        Next



    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load("C:\TXTFILES\apantReqtome.xml") 'C:\TXTFILES\apantSendInv.xml")
        Dim nodes As XmlNodeList = xmlDoc.DocumentElement.SelectNodes("*")

        'ExecuteSQLQuery("delete from oc_category_filter")
        Me.Text = "Ελεγχος αριθμου εγγραφών"
        Application.DoEvents()

        Dim NF As Long = 0
        For Each node As XmlNode In nodes
            NF = NF + 1
            Me.Text = node.InnerText
        Next
        Application.DoEvents()
        Me.Text = NF

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        ' SKROUTZ()
        SKROUTZ_MakeRequest()
        '      curl -X POST https: //www.skroutz.gr/oauth2/token\?\
        'client_id \= your_client_id \&\
        'client_secret\=your_client_secret\&\
        'grant_type \= client_credentials \&\
        'scope\=Public

    End Sub

    Private Async Sub SKROUTZ_MakeRequest()

        Dim client = New HttpClient()

        Try
            client.DefaultRequestHeaders.Add("Accept", "application/vnd.skroutz+json; version=3.0")

            '// client.DefaultRequestHeaders.Add("Authorization", "xobten1524571846")
            client.DefaultRequestHeaders.Add("Authorization", "Bearer BQN-WsvJFdttSjg00rpR1bA6fKquC72hFi3l6DzDYuDFQJb9ksG_yLtcjqeAcgnPNLVJjnWDzdAbeVn1ovB-vQ==")
            ' // client.DefaultRequestHeaders.Add("scope", "public")




            Dim uri = "https://api.skroutz.gr/merchants/ecommerce/orders/210127-9011968"  'hwww.skroutz.gr/oauth2/token\?"







            Dim response As HttpResponseMessage
            Dim xl As String = XDocument.Load("c:\txtfiles\inv.xml").ToString ' "--> εκει έχω αποθηκεύσει το xml που εφτιαξα"
            Dim byteData As Byte() = Encoding.UTF8.GetBytes(xl)

            ' Dim response As HttpResponseMessage
            'response = Await client.PostAsync(uri, "")


            response = Await client.GetAsync(uri)
            Dim result = Await response.Content.ReadAsStringAsync()
            TextBox2.Text = result.ToString

            Dim MF = "c:\txtfiles\sendinv\skr" + Format(Now, "yyyyddMMHHmm") + ".json"
            FileOpen(1, MF, OpenMode.Output)
            PrintLine(1, result.ToString)
            FileClose(1)



            '  {"order"{"code":"210127-5178005",
            '"state""accepted",
            '"customer":
            '  {"id""GQYoapve57","first_name":"Βαγγέλης","last_name":"Ελληνικακης",
            '  "address":{"street_name":"kalomiri","street_number":"22","zip":"83100","city":"samos","region":"Σάμου","pickup_from_collection_point":false}},
            '  "invoice":false,"comments":"","courier":"ACS","courier_voucher":null, "courier_tracking_codes":  [],
            '  "line_items":[
            '  {"id":"MYKOK4gNYn","shop_uid":"1002","product_name":"Motospeed H19 Gaming Headset Grey","quantity":1,"unit_price":20.0,"total_price":20.0,"price_includes_vat":true}],


            '  "created_at":"2021-01-27T13:09:01+02:00","expires_at":"2021-01-28T10:09:01+02:00",
            '  "dispatch_until":"2021-01-29T18:00:00+02:00"}}
            '            order.code = 210127 - 5178005
            '            order.state = accepted
            '            order.customer.id = GQYoapve57
            '            order.customer.first_name = Βαγγέλης
            '            order.customer.last_name = Ελληνικακης
            '            order.customer.address.street_name = kalomiri
            '            order.customer.address.street_number = 22
            '            order.customer.address.zip = 83100
            '            order.customer.address.city = samos
            '            order.customer.address.region = Σάμου
            '            order.customer.address.pickup_from_collection_point = False
            '            order.invoice = False
            '            order.courier = ACS
            '            order.line_items[0].id=MYKOK4gNYn
            'order.line_items[0].shop_uid=1002
            'order.line_items[0].product_name=Motospeed H19 Gaming Headset Grey
            'order.line_items[0].quantity=1
            'order.line_items[0].unit_price=20
            'order.line_items[0].total_price=20
            'order.line_items[0].price_includes_vat=True
            'order.created_at = 27 / 1 / 2021 1:09:01 μμ
            'order.expires_at = 28 / 1 / 2021 10:09:01 πμ
            'order.dispatch_until = 29 / 1 / 2021 6:00:00 μμ


            Dim f, l As String
            FileOpen(1, "c:\txtfiles\jasonstr.txt", OpenMode.Output)
            'PrintLine(1, result.ToString)
            'FileClose(1)
            reader = New JsonTextReader(New StringReader(result))
            Dim c As String
            While reader.Read
                If reader.Value IsNot Nothing Then


                    If InStr(reader.Path, reader.Value) = 0 Then
                        ListBox1.Items.Add(reader.Path + "=" + reader.Value.ToString)
                        PrintLine(1, reader.Path + "=" + reader.Value.ToString)
                    End If
                    c = reader.Value
                    Select Case reader.Path
                        Case "order.customer.first_name"
                            Dim EPO1 As String = c
                        Case "order.customer.last_name"
                            Dim EPO1 As String = c
                        Case "order.customer.address.street_name"
                            Dim EPO1 As String = c
                        Case "order.customer.address.street_number"
                            Dim EPO1 As String = c
                        Case "order.customer.address.zip"
                            Dim EPO1 As String = c
                        Case "order.customer.address.city"
                            Dim EPO1 As String = c
                        Case "order.customer.address.region"
                            Dim EPO1 As String = c
                        Case "order.customer.first_name"
                            Dim EPO1 As String = c
                        Case "order.customer.first_name"
                            Dim EPO1 As String = c
                        Case "order.customer.first_name"
                            Dim EPO1 As String = c





                    End Select



                End If
            End While

            FileClose(1)
















            '  Dim result = Await response.Content.ReadAsStringAsync()




            '   response = Await client.GetAsync(uri)



            'Using content = New ByteArrayContent(byteData)
            'content.Headers.ContentType = New MediaTypeHeaderValue("application/xml")


            ' https://api.skroutz.gr/search?q=apple



            ' response = Await client.PostAsync(uri, content)
            '  Dim result = Await response.Content.ReadAsStringAsync()

            'response = Await client.PostAsync(uri, content)
            'Dim result = Await response.Content.ReadAsStringAsync()
            'Dim MF = "c:\txtfiles\sendinv\apantSendInv" + Format(Now, "yyyyddMMHHmm") + ".xml"
            'FileOpen(1, MF, OpenMode.Output)
            'PrintLine(1, result.ToString)
            'FileClose(1)


            ' End Using


        Catch ex As Exception

            MsgBox(ex.ToString)
        End Try


    End Sub

    Async Sub SKROUTZ()

        Dim client = New HttpClient()
        Dim queryString = HttpUtility.ParseQueryString(String.Empty)



        client.DefaultRequestHeaders.Add("client_id", "netbox")

        client.DefaultRequestHeaders.Add("client_secret", "xobten1524571846")
        client.DefaultRequestHeaders.Add("grant_type", "RaXm-jKHWHAK0P-tws0RVhgMjyfYTzH4-HsZboRJHubXZODRTi4FtRR8SVW7R82oDgRw6EFtIFxW3E5Yh9gxhg==")
        client.DefaultRequestHeaders.Add("scope", "public")




        Dim uri = "hwww.skroutz.gr/oauth2/token\?"
        'client_id \=your_client_id\&\
        'client_secret \=your_client_secret\&\
        ' grant_type \=client_credentials\&\
        'scope \=public
        Dim response As HttpResponseMessage
        'response = Await client.PostAsync(uri, "")
        Dim result = Await response.Content.ReadAsStringAsync()




        response = Await client.GetAsync(uri)

        TextBox2.Text = result.ToString

        Dim MF = "c:\txtfiles\apantReqtome.xml"  'Inv" + Format(Now, "yyyyddMMHHmm") + ".xml"
        FileOpen(1, MF, OpenMode.Output)
        PrintLine(1, result.ToString)
        FileClose(1)

    End Sub

    Function findJ(c As String) As String
        findJ = ""
        If reader.Path.Contains(c) And reader.Path <> c Then
            findJ = reader.Value.ToString

        End If

    End Function



    Public Function checkServer(ByVal check_path As Integer) As Boolean
        Dim c As String
        Dim tmpStr As String
        c = "Config.ini"


        Dim par As String = ""
        Dim mf As String
        mf = c   ' "c:\mercvb\err3.txt"  If System.IO.File.Exists(SavePath) Then
        If Len(Dir(UCase(mf))) = 0 Then
            par = ":DELLAGAKIS\SQL17:sa:12345678:1:EMP"    '" 'G','g','Ξ','D'  "
            par = InputBox("ΒΑΣΗ ΔΕΔΟΜΕΝΩΝ", , par)
            gUserId = InputBox("Χρήστης", gUserId)
            gSubKey = InputBox("Κλειδί", gSubKey)
        Else
            FileOpen(11, mf, OpenMode.Input)
            '   Input(1, par)
            par = LineInput(11)
            gUserId = LineInput(11)
            gSubKey = LineInput(11)
            FileClose(11)
        End If
        If check_path = 1 Then
            par = InputBox("ΒΑΣΗ ΔΕΔΟΜΕΝΩΝ  (CONFIG.INI ΣΤΟΝ ΤΡΕΧΟΝΤΑ ΦΑΚΕΛΟ) ", ":Π.Χ. (local)\sql2012:sa:12345678:1:EMP", par)
            gUserId = InputBox("Χρήστης", gUserId)
            gSubKey = InputBox("Κλειδί", gSubKey)
        End If

        'Input = InputBox("Text:")

        If String.IsNullOrEmpty(par) Then
            ' Cancelled, or empty
            checkServer = False
            ' MsgBox("εξοδος λογω μη σύνδεσης με βάση δεδομένων")
            Exit Function
        Else
            ' Normal
        End If


        FileOpen(7, mf, OpenMode.Output)
        PrintLine(7, par)
        PrintLine(7, gUserId)
        PrintLine(7, gSubKey)

        FileClose(7)

        Dim A As String
        If System.IO.File.Exists("DATES.TXT") Then
            FileOpen(31, "DATES.TXT", OpenMode.Input)
            A = LineInput(31)
            ListBox2.Items.Add(A)
            A = LineInput(31)
            ListBox2.Items.Add(A)
            FileClose(31)

        End If


        ':(local)\sql2012:::2:EMP
        ':(local)\sql2012:sa:12345678:1:EMP





        Try

            ' With FrmSERVERSETTINGS
            OpenFileDialog1.FileName = c
            openedFileStream = OpenFileDialog1.OpenFile()
            'End With

            ReDim dataBytes(openedFileStream.Length - 1) 'Init 
            openedFileStream.Read(dataBytes, 0, openedFileStream.Length)
            openedFileStream.Close()
            tmpStr = par ' System.Text.Encoding.Unicode.GetString(dataBytes)

            '     With FrmSERVERSETTINGS
            If Val(Split(tmpStr, ":")(4)) = 1 Then
                'network
                'gConnect = "Provider=SQLOLEDB.1;" & _
                '           "Data Source=" & Split(tmpStr, ":")(0) & _
                '           ";Network=" & Split(tmpStr, ":")(1) & _
                '           ";Server=" & Split(tmpStr, ":")(1) & _
                '           ";Initial Catalog=" & Trim(Split(tmpStr, ":")(5)) & _
                '           ";User Id=" & Split(tmpStr, ":")(2) & _
                '           ";Password=" & Split(tmpStr, ":")(3)

                gConnect = "Provider=SQLOLEDB.1;;Password=" & Split(tmpStr, ":")(3) &
                ";Persist Security Info=True ;" &
                ";User Id=" & Split(tmpStr, ":")(2) &
                ";Initial Catalog=" & Trim(Split(tmpStr, ":")(5)) &
                ";Data Source=" & Split(tmpStr, ":")(1)

                ''   gConnect = "Provider=SQLOLEDB.1;;Password=" & Split(tmpStr, ":")(3) &
                gSQLCon = "Server=" + Split(tmpStr, ":")(1)
                gSQLCon = gSQLCon + ";Database=" + Trim(Split(tmpStr, ":")(5))
                gSQLCon = gSQLCon + ";Uid=" + Split(tmpStr, ":")(2) + ";Pwd=" + Split(tmpStr, ":")(3)



            Else
                'local
                'MsgBox(Split(tmpStr, ":")(1))
                '  gConnect = "Provider=SQLOLEDB;Server=" & Split(tmpStr, ":")(1) &
                '         ";Database=" & Split(tmpStr, ":")(5) & "; Trusted_Connection=yes;"

                '    gConSQL = "Data Source=" & Split(tmpStr, ":")(1) & ";Integrated Security=True;database=" & Split(tmpStr, ":")(5)
                'cnString = "Data Source=localhost\SQLEXPRESS;Integrated Security=True;database=YGEIA"

            End If
            'End With
            Dim sqlCon As New OleDbConnection
            '
            ' gConnect = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;PWD=12345678;Initial Catalog=D2014;Data Source=logisthrio\sqlexpress"
            'GDB.Open(gConnect)



            'OK
            'gConnect = "Provider=SQLOLEDB.1;;Password=12345678;Persist Security Info=True ;User Id=sa;Initial Catalog=EMP;Data Source=LOGISTHRIO\SQLEXPRESS"
            sqlCon.ConnectionString = gConnect
            sqlCon.Open()
            checkServer = True
            sqlCon.Close()

            '            Dim GDB As New ADODB.Connection

        Catch ex As Exception
            checkServer = False
            MsgBox("εξοδος λογω μη σύνδεσης με βάση δεδομένων")
            'End
        End Try
    End Function

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        MakeRequestKTEL()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Make2RequestKTEL()

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Make3RequestKTELtickets()
    End Sub

    Private Sub tickToParastat_Click(sender As Object, e As EventArgs) Handles tickToParastat.Click
        'xreiazomai hmeromhnia ,kvdiko eidoys,kodiko pelTH,ΑΡΙΘΜΟ ΠΑΡΑΣΤΑΤΙΚΟΥ
        '  insert into EGGTIM (HME,ATIM,EIDOS,TIMM,PELKOD,KODE,POSO)   select '03/26/2024' AS HME,'T000122' AS ATIM,'e' as e,SUM(revenue/100) as r,userId,fareProductId,quantity from EISITIRIA where userId=9 group by userId,fareProductId,quantity
        '  insert into TIM (HME,ATIM,EIDOS,AJI,KPE,KLEIDI)  select '03/26/2024' AS HME,'T000122' AS ATIM,'e' as e,SUM(revenue/100) as r,'USER' AS AA,'T000122' AS ATIM2 from EISITIRIA where userId=9 

        ' METATREPV TIS HMEROMHNIES
        Dim SQLDT As New DataTable
        Execute2SQLQuery("UPDATE EISITIRIA SET [issueDateTimestamp]=CONVERT(DATE,[cissueDateTimestamp],103)", SQLDT)


        Dim SQL55 As New DataTable
        ' ΔΕΝ ΠΑΙΖΕΙ Execute2SQLQuery("SELECT userId from  EISITIRIA where userId not in (SELECT NUM1 FROM PEL )", SQL55)
        Execute2SQLQuery(" Select  userId,P.NUM1,P.EPO from  EISITIRIA E
                       Left Join PEL P ON  E.userId=P.NUM1
                       where P.NUM1 Is NULL", SQL55)
        Dim probl As Integer = 0

        Dim s55ektos As String = ""
            For n55 As Integer = 0 To SQL55.Rows.Count - 1
                s55ektos = s55ektos + SQL55.Rows(0)(0).ToString + " , "
            Next
            If s55ektos = "" Then
            Else
            MsgBox("σεν εχω αντιστοιχία για πελατες:" + s55ektos)
            probl = probl + 1
        End If



        ' Execute2SQLQuery("SELECT fareProductId from  EISITIRIA where fareProductId not in (SELECT NUM1 FROM EID )", SQL55)
        Execute2SQLQuery("SELECT fareProductId from  EISITIRIA S LEFT JOIN EID D ON S.fareProductId=D.NUM1 WHERE D.NUM1 IS NULL", SQL55)
        s55ektos = ""
            For n55 As Integer = 0 To SQL55.Rows.Count - 1
                s55ektos = s55ektos + SQL55.Rows(0)(0).ToString + " , "
            Next
            If s55ektos = "" Then
            Else
            MsgBox("σεν εχω αντιστοιχία για ΕΙΣΙΤΗΡΙΑ:" + s55ektos)
            probl = probl + 1
        End If
        If probl > 0 Then
            MsgBox("αδυνατη η συνεχιση")
            Exit Sub
        End If






        Execute2SQLQuery("UPDATE EISITIRIA SET PELKOD=(SELECT top 1 KOD FROM PEL WHERE NUM1=userId)", SQLDT)
            Execute2SQLQuery("UPDATE EISITIRIA SET KODE=(SELECT top 1 KOD FROM EID WHERE NUM1=fareProductId )", SQLDT)
            UPDATE_ATIM()
            Execute2SQLQuery("delete from EGGTIMTEMP", SQLDT)

            Execute2SQLQuery(" insert into EGGTIMTEMP (TIMM, EIDOS, MONA, MIK_AJIA, POSO, PELKOD, ATIM, KODE, HME   ) SELECT (revenue/100/1.13) as timm,'e','TEM',SUM(totalRevenue/100) as axia, SUM(quantity) as aritmos,PELKOD,ATIM,KODE,convert(date,[issueDateTimestamp]) FROM EISITIRIA GROUP BY PELKOD,ATIM,KODE,convert(date,[issueDateTimestamp]),(revenue/100/1.13)", SQLDT)
        'SELECT SUM(POSO*TIMM) AS AJ1 ,ATIM,PELKOD,HME   FROM [MERCURY].[dbo].[EGGTIMTEMP] GROUP BY ATIM,PELKOD,HME




    End Sub
    Sub UPDATE_ATIM()
        Dim DT As New DataTable
        Dim muserId = -1
        Dim mAtim = "0000000"
        Execute2SQLQuery("SELECT * FROM EISITIRIA ORDER BY userId", DT)
        For K As Integer = 0 To DT.Rows.Count - 1
            Dim userId As Integer = DT.Rows(K)("userId")
            Dim ID As Integer = DT.Rows(K)("ID")

            If userId <> muserId Then
                muserId = userId

                Dim l As String = "C" ' ("ABCDEFGHIJKLMNOPQRSTUVWYXZ").Substring(userID, 1)

                Dim DT3 As New DataTable
                Execute2SQLQuery("UPDATE ARITMISI SET ARITMISI=ARITMISI+1 WHERE ID=55", DT3) ' + Str(50 + userID

                Dim DT2 As New DataTable
                Execute2SQLQuery("SELECT ARITMISI FROM ARITMISI WHERE ID=55", DT2)  ' + Str(50 + userID
                Dim NN As Integer = DT2(0)(0)

                mAtim = l + Format(NN, "000000")
            End If
            Dim DT22 As New DataTable
            Execute2SQLQuery("UPDATE EISITIRIA SET ATIM='" + mAtim + "' WHERE ID=" + Str(ID), DT22)





        Next

        MsgBox("ok")


    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        REQPAROCHOS()

    End Sub

    Async Sub REQPAROCHOS()


        Dim client = New HttpClient()
        Dim queryString = HttpUtility.ParseQueryString(String.Empty)

        ' AYTO XREIAZETAI GIA TA TICKETS
        'client.DefaultRequestHeaders.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", "YOUR_ACCESS_TOKEN")







        'ok    Dim jsonContent As String = " {""username"":""glagakis@gmail.com"", ""email"":""glagakis@gmail.com"",""password"": ""y39213921!!Y"",""subscriptionKey"":""32657FD7929B483AABBD1C633401E945"" }"



        Dim memail As String
        Dim jsonContent As String = " {""username"":""" + memail + """, ""email"":""glagakis@gmail.com"",""password"": ""y39213921!!Y"",""subscriptionKey"":""32657FD7929B483AABBD1C633401E945"" }"




        ' Dim jsonContent As String = "{""syncAll"":""0"", ""syncTables"": ""[1,2,3,8,9,10,11,12"",""devicesUuid"":""ERP"",""lastSyncDate"":""null"",""fetchAllUsers"" : ""true"" }"

        Dim content As New StringContent(jsonContent, Encoding.UTF8, "application/json")

        'client.DefaultRequestHeaders.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", "YOUR_ACCESS_TOKEN")









        ' Dim cc As String = InputBox("ΑΠΟ ΠΟΙΟ ΜΑΡΚ ΚΑΙ ΜΕΤΑ;")
        'queryString("mark") = cc  ' "1000000006337" ' "{string}"
        '  queryString("nextPartitionKey") = "{string}"
        '   queryString("nextRowKey") = "{string}"

        Dim uri = "https://beta-account.parochos.gr/api/account/loginToSubscription"

        Dim response = Await client.PostAsync(uri, content)
        Dim result = Await response.Content.ReadAsStringAsync()
        TextBox2.Text = result.ToString

        Dim MF = "c:\txtfiles\apantPAROCHOS.xml"  'Inv" + Format(Now, "yyyyddMMHHmm") + ".xml"
        FileOpen(1, MF, OpenMode.Output)
        PrintLine(1, result.ToString)
        FileClose(1)


        Dim json As String = result.ToString
        Dim ser As JObject = JObject.Parse(json)
        Dim c As String = ""

        '  For K As Integer = 0 To ser("jwt").Count - 1


        fbearer = ser("jwt").ToString

        '    Next


















        '' VB.NET1

        'Dim vm = New JArray()

        '' foreach(var cycle In cycles)
        ''{
        ''JObject.FromObject(cycle);
        ''   // extend cycleJson ......




        'Dim cycleJson = JObject.FromObject(
        '   New With {
        '        Key .name = "Jeff",
        '        Key .id = 1,
        '        Key .car = "BMW"
        '    }
        ').ToString()

        'vm.Add(cycleJson)




        'Dim cycleJson2 = JObject.Parse("{""name"":""john""}")

        ''//add surname
        'cycleJson2("surname") = "doe"

        ''//add a complex object
        'cycleJson2("complexObj") = JObject.FromObject(New With
        '                                              {
        '                                              Key .id = 1,
        '                                              Key .Name = "test"}
        '                                              )
        ''So the final json will be

        ''{
        ''  "name": "john",
        ''  "surname": "doe",
        ''  "complexObj": {
        ''    "id": 1,
        ''    "name": "test"
        ''  }
        ''}

        'vm.Add(cycleJson2)















    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        sendtimol()
    End Sub
    Async Sub sendtimol()


        Dim iss As New Issuer
        iss.vatNumber = "028783755"
        iss.branch = 0
        iss.street = "Κ.ΠΑΛΑΙΟΛΟΓΟΥ"
        iss.streetNumber = "21"
        iss.postalCode = "66100" '       "postalCode": "postalCode",
        iss.city = "ΔΡΑΜΑ" '       "city": "city",
        iss.country = "GR" '       "country": "GR"






        Dim cou As New Counterpart
        cou.vatNumber = "150144455"
        cou.name = "ΛΑΓΑΚΗΣ ΠΑΡΑΣΚΕΥΑΣ"
        cou.branch = 0
        cou.street = "ΣΑΜΑΚΟΒΙΟΥ"
        cou.streetNumber = "13"
        cou.postalCode = "66100" '       "postalCode": "postalCode",
        cou.city = "ΔΡΑΜΑ" '       "city": "city",
        cou.country = "GR" '
        cou.email = "test@email.com" '


        Dim deliv As New DeliveryAddress
        deliv.street = "ΣΑΜΑΚΟΒΙΟΥ "
        deliv.streetNumber = "13"
        deliv.postalCode = "66100" '       "postalCode": "postalCode",
        deliv.city = "ΔΡΑΜΑ" '       "city": "city",
        deliv.country = "GR" '

        '  Dim det(10) As InvoiceDetail
        Dim newitem(10) As InvoiceDetail

        newitem(0) = New InvoiceDetail()
        newitem(0).lineNumber = 1
        newitem(0).entityName = "KARIDIA"
        'newitem(0).recType = Null
        newitem(0).invoiceDetailType = 0
        newitem(0).quantity = 1
        newitem(0).netValue = 100
        newitem(0).vatCategory = 1
        newitem(0).vatCategoryUbl = "S"
        newitem(0).vatAmount = 24
        newitem(0).vatPercent = 24
        newitem(0).measurementUnit = 1
        newitem(0).measurementUnitUbl = "H87"
        newitem(0).lineComments = "ΣΧΟΛΙΑ"

        newitem(0).totalValue = 124
        newitem(0).classificationCategory = "category1_1"
        newitem(0).classificationType = "E3_561_001"
        newitem(0).cpvCode = "45233222-1"



        newitem(1) = New InvoiceDetail()
        newitem(1).lineNumber = 2
        newitem(1).entityName = "KARIDIA"
        'newitem(1).recType = Null
        newitem(1).invoiceDetailType = 0
        newitem(1).quantity = 1
        newitem(1).netValue = 100
        newitem(1).vatCategory = 1
        newitem(1).vatCategoryUbl = "S"
        newitem(1).vatAmount = 24
        newitem(1).vatPercent = 24
        newitem(1).measurementUnit = 1
        newitem(1).measurementUnitUbl = "H87"
        newitem(1).lineComments = "ΣΧΟΛΙΑ"

        newitem(1).totalValue = 124
        newitem(1).classificationCategory = "category1_1"
        newitem(1).classificationType = "E3_561_001"
        newitem(1).cpvCode = "45233222-1"

        Dim pay(5) As PaymentMethod

        pay(0) = New PaymentMethod()
        pay(0).type = 5
        pay(0).amount = 248





        Dim he As New InvoiceHeader
        he.series = "b"
        he.issueDate = "2024-04-17"
        he.dispatchDate = "2024-04-17"
        he.aa = 204
        he.invoiceCode = "INV000204"
        he.invoiceTypeUbl = "380"
        he.invoiceType = "1.1"
        he.currency = "EUR"
        he.movePurpose = 1
        he.selfPricing = False

        he.fuelInvoice = False




        Dim invs As New InvoiceSummary

        invs.totalVatAmount = 48.0
        invs.totalValue = 248.0
        invs.totalNetValue = 200.0





        Dim inv As New Invoice
        inv.issuer = iss
        inv.counterpart = cou
        inv.invoiceHeader = he
        inv.invoiceDetails.Add(newitem(0))
        inv.invoiceDetails.Add(newitem(1))
        inv.paymentMethods.Add(pay(0))
        inv.invoiceSummary = invs
        inv.deliveryAddress = deliv
        inv.invoiceSummary = invs




        Dim SOU As New Source2
        SOU.invoice = inv

        Dim ex2 As New Example
        ex2.externalSystemId = "187"
        ex2.source = SOU

        Dim ResultJSON As String = JsonConvert.SerializeObject(ex2).ToString
        Debug.Print(ResultJSON) '

        '  Exit Sub





        Dim client = New HttpClient()
        Dim queryString = HttpUtility.ParseQueryString(String.Empty)



        Dim ResultJSON12 As String = " { ""externalSystemId"" : ""185"",  ""source"": {
        ""invoice"": {
            ""issuer"": {
                ""vatNumber"": ""028783755"",
                ""branch"": 0,
                ""street"": ""street"",
                ""streetNumber"": ""streetNumber"",
                ""postalCode"": ""postalCode"",
                ""city"": ""city"",
                ""country"": ""GR""
            },
            ""counterpart"": {
                ""vatNumber"": ""028783755"",
                ""branch"": 0,
                ""name"": ""name"",
                ""street"": ""street"",
                ""streetNumber"": ""streetNumber"",
                ""postalCode"": ""postalCode"",
                ""city"": ""city"",
                ""country"": ""GR"",
                ""email"": ""test@email.com""
            },
            ""representative"": {
                ""vatNumber"": ""000000000"",
                ""name"": ""name""
            },
            ""deliveryAddress"": {
                ""street"": ""street"",
                ""streetNumber"": ""streetNumber"",
                ""postalCode"": ""postalCode"",
                ""city"": ""city"",
                ""country"": ""GR""
            },
           
            ""invoiceHeader"": {
                ""series"": ""INV"",
                ""aa"": 13,
                ""issueDate"": ""2024-04-14"",                
                ""dispatchDate"": ""2024-04-14"",
                ""invoiceCode"": ""INV-00013"",
                ""invoiceType"": ""1.1"",
                ""invoiceTypeUbl"": ""380"",
                ""currency"": ""EUR"",
                ""selfPricing"": false,
                ""movePurpose"": 1,
                ""fuelInvoice"": false
            },
            ""invoiceDetails"": [
                {
                    ""lineNumber"": 1,
                    ""recType"": 1,
                    ""quantity"": 1,
                    ""entityName"": ""string"",
                    ""invoiceDetailType"": 0,
                    ""netValue"": 100,
                    ""totalValue"": 124,
                    ""vatCategory"":  ""24"", 
                   
                  
                    ""vatAmount"": 24,
                    ""vatPercent"": 24,
                    ""measurementUnit"": ""H87"",
                    ""measurementUnitUbl"": ""H87"",
                    ""lineComments"": ""ΣΧΟΛΙΑ"",
                    ""classificationCategory"": ""category1_1"",
                    ""classificationType"": ""E3_561_1"",
                   
                    ""cpvCode"": ""45233222-1""
                   
                }
            ],
           
            ""paymentMethods"": [
                {
                    ""type"": 5,
                    ""amount"": 124
                }
            ],
           
            ""invoiceSummary"": {
                ""totalNetValue"": 100,
                ""totalVatAmount"": 24,
                ""totalValue"": 124
            },
            ""Messages"": [
                {
                    ""type"": 0,
                    ""recipients"": ""test1@email.com;test2@email.com"",
                    ""cc"": ""test3@email.com;test4@email.com"",
                    ""templateIdentifier"": ""0000001""
                },
                {
                    ""type"": 1,
                    ""recipients"": ""6983022568;6983022868""
                }
            ]
        }
    }
}"



        '""taxInfo"": {
        '               ""taxCategory"": 1,
        '               ""taxCategoryUbl"": ""50"",
        '               ""underlyingValue"": 100,
        '               ""taxPercent"": 20
        '           }

        Dim jsonContent As String = ResultJSON ' " {""externa





        Dim content As New StringContent(jsonContent, Encoding.UTF8, "application/json")

        'client.DefaultRequestHeaders.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", "YOUR_ACCESS_TOKEN")


        client.DefaultRequestHeaders.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", fbearer)








        ' Dim cc As String = InputBox("ΑΠΟ ΠΟΙΟ ΜΑΡΚ ΚΑΙ ΜΕΤΑ;")
        'queryString("mark") = cc  ' "1000000006337" ' "{string}"
        '  queryString("nextPartitionKey") = "{string}"
        '   queryString("nextRowKey") = "{string}"

        Dim uri = "https://beta-srv.parochos.gr/api/send"

        Dim response = Await client.PostAsync(uri, content)
        Dim result = Await response.Content.ReadAsStringAsync()
        TextBox2.Text = result.ToString


        ' Exit Sub

        Dim MF = "c:\txtfiles\apanttimPAROCHOS.xml"  'Inv" + Format(Now, "yyyyddMMHHmm") + ".xml"
        FileOpen(1, MF, OpenMode.Output)
        PrintLine(1, result.ToString)
        FileClose(1)


        Dim json As String = result.ToString
        Dim ser As JObject = JObject.Parse(json)
        Dim c As String = ""

        ' For K As Integer = 0 To ser("jwt").Count - 1


        '  fbearer = ser("jwt")(K)("userId").ToString

        ' Next










    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        forol()
    End Sub
    Private Async Sub forol()

        Dim client = New HttpClient()
        Dim queryString = HttpUtility.ParseQueryString(String.Empty)
        ' client.DefaultRequestHeaders.Add("aade-user-id", gUserId)
        '  client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", gSubKey)
        '  Dim cc As String = InputBox("ΑΠΟ ΠΟΙΟ ΜΑΡΚ ΚΑΙ ΜΕΤΑ;")
        '  queryString("mark") = cc  ' "1000000006337" ' "{string}"
        '  queryString("nextPartitionKey") = "{string}"
        '   queryString("nextRowKey") = "{string}"
        Dim uri = "http://localhost:11180/webservice/checkreceipt?recno=000008&format=xml" ' "https://mydata-dev.azure-api.net/RequestDocs?" & queryString.ToString

        Dim response = Await client.GetAsync(uri)
        Dim result = Await response.Content.ReadAsStringAsync()
        TextBox2.Text = result.ToString

        Dim MF = "c:\txtfiles\apantforolReqtome.xml"  'Inv" + Format(Now, "yyyyddMMHHmm") + ".xml"
        FileOpen(1, MF, OpenMode.Output)
        PrintLine(1, result.ToString)
        FileClose(1)

    End Sub
End Class






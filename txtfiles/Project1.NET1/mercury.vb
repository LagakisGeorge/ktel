Option Strict Off
Option Explicit On
Imports System.Windows.Forms
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Xml
Imports System.Text
Imports System.IO

Imports System.Data.OleDb
Imports System.Net.NetworkInformation
Imports Microsoft.VisualBasic.Compatibility.VB6





<System.Runtime.InteropServices.ComVisible(False)> Friend Class main
    Inherits System.Windows.Forms.Form

    Public dataBytes() As Byte

    Public sqlDT As New DataTable
    Public sqlDaTaSet As New DataSet
    Public sqlDTx As New DataTable
    Public openedFileStream As System.IO.Stream

    Public gSplitter As String = ";"


    Dim SQLDT2 As New DataTable













    Dim gdb As New ADODB.Connection
    Dim gConnect As String
    Dim xl As New Microsoft.Office.Interop.Excel.Application

    Dim xlsheet As Microsoft.Office.Interop.Excel.Worksheet

    Dim xlsheet3 As Microsoft.Office.Interop.Excel.Worksheet
    Dim xlwbook As Microsoft.Office.Interop.Excel.Workbook

    Dim ROW As Integer
    Dim COL As Integer


    Dim rowId As Integer = 7
    Dim rowIdINNER As Integer = 7



    Dim f_logPel As String '30-00-00-0000
    Dim Party_IDParty As String  '12344   ΚΩΔ ΣΥΝΑΛΛΑΣΟΜΕΝΟΥ
    Dim AM_DcTp_Dscr As String '=""Τιμολόγιο Παροχής Υπηρεσιών
    Dim Party_AFM As String ' =""999349996
    Dim Party_ADDRESS As String ' =""ΠΟΛΥΣΤΗΛΟ ΚΑΒΑΛΑΣ
    Dim AM_DcTp_cd As String ' =""#ΤΥΠ-0""
    Dim AMO_Srl_DSCR As String  '=""ΠΩΛΗΣΕΙΣ (ΜΗΧΑΝΟΓΡΑΦΗΜΕΝΗ)
    Dim Base_dt As String '=""2014-05-07""
    Dim Base_INVOICE As String ' =""#ΤΥΠ-0/Π000/1/Τιμολόγιο Παροχής Υπηρεσιών"
    Dim Party_SNAME As String '=""Θ. ΓΡΑΜΜΑΤΗΣ Κ.ΣΙΑ Ε.Ε""
    Dim KAU_AJIA As String
    Dim FPA As String

    Dim FL_Ledg_Dscr As String ' =""ΠΩΛΗΣΕΙΣ ΥΠΗΡΕΣΙΩΝ ΧΟΝΔΡΙΚΗΣ ΕΣΩΤΕΡΙΚΟΥ ΦΠΑ 23%
    Dim FL_Ledg_cd As String '=""73-00-00-0057
    Dim KAU_AJIA1 As String
    Dim FPA1 As String
    Dim IsHand As String
    Dim cdRetailIdentity As String
    Dim System_sys As String   ') 'SB =POLISEIS FR PISTVTIKA YPIRESIES FP= PLIROMES BP=AGORES
    Dim MVTP As String

    ' Dim Party_IDParty As String
    Dim tit_paras As String



    Dim kau13 As Single, f_70_13 As String, fpa13 As Single
    Dim kau23 As Single, f_70_23 As String, fpa23 As Single

    Dim kau24 As Single, f_70_24 As String, fpa24 As Single
    Dim kau17 As Single, f_70_17 As String, fpa17 As Single


    Dim kau16 As Single, f_70_16 As String, fpa16 As Single
    Dim kau9 As Single, f_70_9 As String, fpa9 As Single
    Dim kau0 As Single, f_70_0 As String ', fpa9 As Single

    Dim LOG13 As String, LOG23 As String, logarFpa23 As String
    Dim LOG16 As String, LOG9 As String, logarFpa13 As String
    Dim LOG0 As String

    Dim LOG24 As String, logarFpa24 As String
    Dim LOG17 As String, logarFpa17 As String

    Dim f_aitiologia As String


    Dim fnTimol, fnLian, fnPistTim, fnPistLian, fnTimAg, fnPistAg, fnPAR As Integer
    Dim fcTimol, fcLian, fcPistTim, fcPistLian, fcTimAg, fcPistAg, fcPAR As String






    Dim Metrhtaxond As Boolean, f_logTam As String
















    '46888  repaki

    '    Private Sub EXCELTOXML_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    '   End Sub

    '    Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

    '   End Sub
    '    Imports System.Text
    'Imports System.IO


    '    Public Class Tester
    '        Public Shared Sub Main()
    '            Dim myFileStream As FileStream
    '            Dim myStreamWriter As StreamWriter
    '            Dim strWrite As String

    '            Dim StreamEncoding As Encoding

    '            Try
    '                StreamEncoding = Encoding.Default
    '                'StreamEncoding = Encoding.Unicode
    '                'StreamEncoding = Encoding.UTF8
    '                'StreamEncoding = Encoding.UTF7

    '                myFileStream = New FileStream("test.txt", FileMode.OpenOrCreate, FileAccess.Write)
    '                myStreamWriter = New StreamWriter(myFileStream, StreamEncoding)
    '                strWrite = "asdf"

    '                myStreamWriter.Write(strWrite)
    '                myStreamWriter.Close()
    '                myFileStream.Close()

    '            Catch EX As IOException
    '                Console.WriteLine(EX.Message)
    '            End Try


    '        End Sub
    '    End Class

    '   Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)



    '  End Sub


    'Private Sub EXCELTOXML_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EXCELTOXML.Click
    '    Dim a As String
    '    Dim K As Short
    '    Dim C As String

    '    ' Write the string as utf-8.
    '    ' This also writes the 3-byte utf-8 preamble at the beginning of the file.
    '    Dim appendMode As Boolean = False ' This overwrites the entire file.
    '    Dim sw As New StreamWriter("C:\MERCVB\out_utf9.export", appendMode, System.Text.Encoding.UTF8)
    '    'sw.Write(TextBox1.Text)
    '    'sw.Close()


    '    Dim xlApp As Excel.Application
    '    Dim xlWorkBook As Excel.Workbook
    '    Dim xl As Excel.Worksheet

    '    xlApp = New Excel.ApplicationClass
    '    xlWorkBook = xlApp.Workbooks.Open(TextBox1.Text)
    '    xl = xlWorkBook.Worksheets(1) ' .Add

    '    'xlwbook = xl.Workbooks.Open(TextBox1.Text)
    '    'xlsheet = xlwbook.Sheets.Item(1)





    '    Dim rH As New ADODB.Recordset
    '    rH.Open("select * from EPSILON", gdb, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

    '    Dim rD As New ADODB.Recordset
    '    rD.Open("select * from EPSDETAIL", gdb, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

    '    'UPGRADE_NOTE: enter was upgraded to enter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    '    Dim enter_Renamed As String
    '    'UPGRADE_WARNING: Couldn't resolve default property of object enter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '    enter_Renamed = Chr(13)

    '    FileOpen(1, "C:\MERCVB\A778.XML", OpenMode.Output)
    '    ROW = 1

    '    a = "<?xml version=""1.0""?>"
    '    'UPGRADE_WARNING: Couldn't resolve default property of object enter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '    a = a + enter_Renamed + "<Data Style=""BRowse"" Name=""SX"">"
    '    Dim rowId As Integer = 7
    '    Dim rowIdINNER As Integer = 7
    '    '===============================================================================real onomatepvmymo 54100
    '    Do While True
    '        ROW = ROW + 1
    '        If IsDBNull(xl.Cells(ROW, 12).value) Then
    '            Exit Do
    '        End If

    '        If Len(xl.Cells(ROW, 11).ToString) < 2 Then
    '            Exit Do
    '        End If
    '        If xl.Cells(ROW, 11).value = Nothing Then
    '            Exit Do
    '        End If



    '        '' ΑΝ ΤΑ ΤΡΑΒΑΩ ΑΠΟ SQLSERVER
    '        'C = "" ' HEADER
    '        'For K = 0 To rH.Fields.Count - 1
    '        '    C = C & rH.Fields(K).Name & "=""" + Replace(rH.Fields(K).Value, "&", "") + """ "
    '        'Next

    '        ''''''''''''''''''''GEORGIADIS'''''''''''''''''tim1	AJ1-2	AJ2-3	AJ3-4	AJI-5	FPA1-6	FPA2-7	hmer-8	EPO-9	AFM-10	EPA-11	DIE-12	POL-13	PIST-14	tim-15


    '        '1	 2	    3	4	5	6	7	    8	    9	    10	    11	    12	13	14	15	16	17	    18	19
    '        'AJ1 AJ2	AJ3	AJ4	AJ5	AJI	FPA1	FPA2	FPA3	FPA4	ATIM	HME	EPO	AFM	KPE	DIE	XRVMA	EPA	POL
    '        Party_IDParty = xl.Cells(ROW, 14).value  ' As String  '12344   ΚΩΔ ΣΥΝΑΛΛΑΣΟΜΕΝΟΥ
    '        AM_DcTp_Dscr = "Τιμολόγιο"
    '        Party_AFM = xl.Cells(ROW, 14).value  'Dim Party_AFM As String ' =""999349996
    '        Party_ADDRESS = xl.Cells(ROW, 16).value 'ToString  ' "ΠΟΛΥΣΤΗΛΟ ΚΑΒΑΛΑΣ"
    '        AM_DcTp_cd = "#ΤΥΠ-0"
    '        AMO_Srl_DSCR = "ΠΩΛΗΣΕΙΣ"
    '        Base_dt = VB6.Format(xl.Cells(ROW, 12), "YYYY-mm-dd")
    '        Base_INVOICE = xl.Cells(ROW, 11).value  ' =""#ΤΥΠ-0/Π000/1/Τιμολόγιο Παροχής Υπηρεσιών"
    '        Party_SNAME = xl.Cells(ROW, 13).value  '=""Θ. ΓΡΑΜΜΑΤΗΣ Κ.ΣΙΑ Ε.Ε""

    '        KAU_AJIA = nVal(xl.Cells(ROW, 1).value) + nVal(xl.Cells(ROW, 2).value) + nval(xl.Cells(ROW, 3).value) + nval(xl.Cells(ROW, 4).value) + nval(xl.Cells(ROW, 5).value)
    '        FPA = nVal(xl.Cells(ROW, 7).value) + nVal(xl.Cells(ROW, 8).value) + nVal(xl.Cells(ROW, 9).value) + nVal(xl.Cells(ROW, 10).value)

    '        FL_Ledg_Dscr = "ΠΩΛΗΣΕΙΣ ΧΟΝΔΡΙΚΗΣ ΕΣ. ΦΠΑ23%"
    '        FL_Ledg_cd = "70-00-00-0057"
    '        KAU_AJIA1 = KAU_AJIA
    '        FPA1 = FPA

    '        C = "<row rowId=""" + LTrim(Str(rowId)) + """ mode=""3"" name=""Hd""><data><new "
    '        C = C + "System_Dscr_1=""Πωλήσεις"" Party_IDParty=""" + Party_IDParty + """ APA_VIES_v_Dscr=""EL"" GlbCff=""1"" ExpenditureKind=""0"" "
    '        C = C + "AM_DcTp_Dscr=""" + AM_DcTp_Dscr + """ Party_AFM=""" + Party_AFM + """ ConstrCost=""0"" "
    '        C = C + "Party_ISK_D_A_Dscr="""" dumm=""0"" AM_DcTp_cd=""" + AM_DcTp_cd + """ Party_ADDRESS=""" + Party_ADDRESS + """ "

    '        a = a & C
    '        sw.WriteLine(a)
    '        a = "" : C = ""



    '        C = C + "Party_CASTVAT=""1"" AMO_Srl_DSCR=""" + AMO_Srl_DSCR + """ Base_dt=""" + Base_dt + """ System_sys=""SB"" "
    '        C = C + "F_Sites_dscr=""ΚΕΝΤΡΙΚΟ"" Party_DOY="""" cdRetailIdentity="""" AMO_Srl_cd=""Π000"" "
    '        C = C + "Party_Sts=""1"" Base_INVOICE=""" + Base_INVOICE + """ F_Sites_cd=""001"" "
    '        C = C + "IsHand="""" Party_SNAME=""" + Party_SNAME + """ Party_CASTVAT_Dscr=""ΚΑΝΟΝΙΚΟ"" "
    '        C = C + "KepyoCatData_ISAGRYP=""0"" KepyoCatData_SUMKEPYOYP=""" + KAU_AJIA + """ KepyoCatData_SUMKEPYOVAT=""" + FPA + """ "
    '        C = C + "Ledger_Cust=""30-00-00-0000"""




    '        a = a & C & "/></data>"


    '        PrintLine(1, a)
    '        sw.WriteLine(a)
    '        ' sw.Write(a)
    '        a = ""



    '        a = a + enter_Renamed + "<detail><row rowId=""" + LTrim(Str(rowId)) + """ mode=""3"" name=""Mv""><data><new "


    '        'C = ""
    '        'For K = 0 To rD.Fields.Count - 1
    '        '    C = C & rD.Fields(K).Name & "=""" + rD.Fields(K).Value + """ "
    '        'Next



    '        'tim1	AJ1-2	AJ2-3	AJ3-4	AJI-5	FPA1-6	FPA2-7	hmer-8	EPO-9	AFM-10	EPA-11	DIE-12	POL-13	PIST-14	tim-15
    '        C = " FL_Ledg_Dscr=""" + FL_Ledg_Dscr + """ FL_Ledg_cd=""" + FL_Ledg_cd + """ "
    '        C = C + "VatVal=""" + FPA1 + """ NetVal=""" + KAU_AJIA1 + """ RegVal=""" + KAU_AJIA1 + """ MvTp=""1"" RegVatVal=""0.0000"""


    '        a = a + enter_Renamed + C + "/></data>"

    '        PrintLine(1, a)
    '        sw.WriteLine(a)
    '        'sw.Write(a)
    '        a = ""

    '        'σειρα με ολες τις βαθμιδες των λογαριασμών

    '        a = a + enter_Renamed + "<detail><row rowId=""" + LTrim(Str(rowIdINNER)) + """ mode=""3"" name=""Ledg""><data><new Active=""1"" dscrLedg=""Πωλήσεις 23%"" cdLedg=""70-00-00-0057"" Anali=""0"" CanMv=""1""/></data></row>"
    '        rowIdINNER = rowIdINNER + 11
    '        a = a + "<row rowId=""" + LTrim(Str(rowIdINNER)) + """ mode=""3"" name=""Ledg""><data><new Active=""1"" dscrLedg=""Πωλήσεις"" cdLedg=""70-00-00"" Anali=""0"" CanMv=""0""/></data></row>"
    '        rowIdINNER = rowIdINNER + 11
    '        a = a + "<row rowId=""" + LTrim(Str(rowIdINNER)) + """ mode=""3"" name=""Ledg""><data><new Active=""1"" dscrLedg=""Πωλήσεις"" cdLedg=""70-00"" Anali=""0"" CanMv=""0""/></data></row>"
    '        rowIdINNER = rowIdINNER + 11
    '        a = a + "<row rowId=""" + LTrim(Str(rowIdINNER)) + """ mode=""3"" name=""Ledg""><data><new Active=""1"" dscrLedg=""Πωλήσεις Εμπορευμάτων"" cdLedg=""70"" Anali=""0"" CanMv=""0""/></data></row>"

    '        a = a + "</detail>"
    '        a = a + enter_Renamed + "</row></detail></row>"


    '        PrintLine(1, a)
    '        sw.WriteLine(a)
    '        a = ""



    '        rowId = rowId + 11
    '        rowIdINNER = rowIdINNER + 11


    '    Loop

    '    a = a + "</Data>"


    '    PrintLine(1, a)
    '    FileClose(1)


    '    sw.Write(a)
    '    sw.Close()




    '    'gdb.EXECUTE "UPDATE EPSILON SET System_Dscr_1='Αγορές',Party_IDParty='60',APA_VIES_v_Dscr='EL',GlbCff='1',ExpenditureKind='0',AM_DcTp_Dscr='Τιμολόγιο,Αγοράς,',Party_AFM='82296964',ConstrCost='0',Party_ISK_D_A_Dscr='ΚΑΝΟΝΙΚΟΣ',dumm='0',AM_DcTp_cd='#ΤΑΓ-0',Party_ADDRESS='ΝΕΑ,ΜΠΑΦΡΑ,ΣΕΡΡΩΝ',Party_CASTVAT='1',AMO_Srl_DSCR='ΑΓΟΡΕΣ,(ΧΕΙΡΟΓΡΑΦΗ)',Base_dt='2014-04-03',System_sys='BP',Party_ISK_D_A_CD='0',F_Sites_dscr='ΚΕΝΤΡΙΚΟ',Party_DOY='5621',Party_JOB='ΕΜΠΟΡ,ΕΛΑΣΤΙΚ-ΓΕΩΡΓΙ',cdRetailIdentity='',Ledger_Supl='50-00-00-0000',AMO_Srl_cd='ΑΓ00',Party_Sts='1',Base_INVOICE='#ΤΑΓ-0/ΑΓ00/22319/',F_Sites_cd='001',IsHand='',Party_SNAME='Ι.ΘΕΟΔΩΡΙΔΗΣ,&,ΣΙΑ,Ο.Ε.',Party_CASTVAT_Dscr='ΚΑΝΟΝΙΚΟ',KepyoCatData_ISAGRYP='0',KepyoCatData_SUMKEPYOYP='23.1700',KepyoCatData_SUMKEPYOVAT='5.3300'"
    '    MsgBox("ok")


    '    'gdb.EXECUTE "UPDATE EPSDETAIL SET FL_Ledg_Dscr='ΤΑΜΕΙΟ',FL_Ledg_cd='38-00-00-0000',VatVal='0.0000',NetVal='28.5000',RegVal='0.0000',MvTp='8',RegVatVal='0.0000'"



    '    xlApp.Quit()
    'End Sub




    'Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    'End Sub
    Function nVal(ByVal c As String) As Single
        Dim n1 As Integer, n2 As Integer
        n1 = InStr(c, ",")
        n2 = InStr(c, ".")
        If n1 > n2 Then ' p.x. 1.000,99  ή  120,99
            c = Replace(c, ".", "")
            c = Replace(c, ",", ".")
            nVal = Val(c)
        Else  ' p.x. 12,000.87  12.87
            c = Replace(c, ",", "")  '1200.87   12.87
            nVal = Val(c)
        End If

    End Function

    '   Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)






    Private Sub createNode(ByVal pID As String, ByVal pName As String, ByVal pPrice As String, ByVal writer As XmlTextWriter)
        writer.WriteStartElement("row")
        writer.WriteAttributeString("rowid", "7")
        writer.WriteAttributeString("mode", "3")
        writer.WriteAttributeString("name", "Hd")

        writer.WriteStartElement("Product_id")
        '<row rowId="7" mode="3" name="Hd">

        writer.WriteString(pID)
        writer.WriteEndElement()

        writer.WriteStartElement("Product_name")
        writer.WriteString(pName)
        writer.WriteEndElement()

        writer.WriteStartElement("Product_price")
        writer.WriteString(pPrice)
        writer.WriteEndElement()
        writer.WriteEndElement()
    End Sub
    Private Shared m_Document As String = "c:\mercvb\sampledata.xml"

    Public Shared Sub write2()
        Dim writer As XmlWriter = Nothing

        Try

            Dim settings As XmlWriterSettings = New XmlWriterSettings()
            settings.Indent = True
            writer = XmlWriter.Create(m_Document, settings)

            writer.WriteComment("sample XML fragment")

            ' Write an element (this one is the root).
            writer.WriteStartElement("book")

            ' Write the namespace declaration.
            writer.WriteAttributeString("xmlns", "bk", Nothing, "urn:samples")

            ' Write the genre attribute.
            writer.WriteAttributeString("genre", "novel")
            '<book xmlns:bk="urn:samples" genre="novel">

            ' Write the title.
            writer.WriteStartElement("title")
            writer.WriteString("The Handmaid's Tale")
            writer.WriteEndElement()

            ' Write the price.
            writer.WriteElementString("price", "19.95")

            ' Lookup the prefix and write the ISBN element. 
            Dim prefix As String = writer.LookupPrefix("urn:samples")
            writer.WriteStartElement(prefix, "ISBN", "urn:samples")
            writer.WriteString("1-861003-78")
            writer.WriteEndElement()

            ' Write the style element (shows a different way to handle prefixes).
            writer.WriteElementString("style", "urn:samples", "hardcover")

            ' Write the close tag for the root element.
            writer.WriteEndElement()

            ' Write the XML to file and close the writer.
            writer.Flush()
            writer.Close()

        Finally
            If Not (writer Is Nothing) Then
                writer.Close()
            End If
        End Try

        '        <?xml version="1.0" encoding="utf-8"?>
        '<!--sample XML fragment-->
        '<book xmlns:bk="urn:samples" genre="novel">
        '  <title>The Handmaid's Tale</title>
        '  <price>19.95</price>
        '  <bk:ISBN>1-861003-78</bk:ISBN>
        '  <bk:style>hardcover</bk:style>
        '</book>





    End Sub 'Main 

    Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button
        Me.EXCELTOXML = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.CD1 = New System.Windows.Forms.OpenFileDialog
        Me.XMLTEXTWRITER = New System.Windows.Forms.Button
        Me.xmlFin = New System.Windows.Forms.Button
        Me.bres_file = New System.Windows.Forms.Button
        Me.pol23 = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.pol13 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.pel30 = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.EPIS13 = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.EPIS23 = New System.Windows.Forms.TextBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.episLian13 = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.episLian23 = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Lian13 = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Lian23 = New System.Windows.Forms.TextBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.Label10 = New System.Windows.Forms.Label
        Me.POL9 = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.POL16 = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.POL0 = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.LIAN0 = New System.Windows.Forms.TextBox
        Me.Button6 = New System.Windows.Forms.Button
        Me.Button8 = New System.Windows.Forms.Button
        Me.filexml = New System.Windows.Forms.TextBox
        Me.xmlG = New System.Windows.Forms.Button
        Me.Label14 = New System.Windows.Forms.Label
        Me.LOGFPA13 = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.LOGFPA23 = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.lianLOGFPA13 = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.lianLOGFPA23 = New System.Windows.Forms.TextBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label61 = New System.Windows.Forms.Label
        Me.arTam = New System.Windows.Forms.TextBox
        Me.Label52 = New System.Windows.Forms.Label
        Me.lianLOGFPA24 = New System.Windows.Forms.TextBox
        Me.Label53 = New System.Windows.Forms.Label
        Me.episLian24 = New System.Windows.Forms.TextBox
        Me.Label54 = New System.Windows.Forms.Label
        Me.lian24 = New System.Windows.Forms.TextBox
        Me.prom50 = New System.Windows.Forms.TextBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.Label49 = New System.Windows.Forms.Label
        Me.LOGFPA17 = New System.Windows.Forms.TextBox
        Me.Label50 = New System.Windows.Forms.Label
        Me.EPIS17 = New System.Windows.Forms.TextBox
        Me.POL17 = New System.Windows.Forms.TextBox
        Me.Label51 = New System.Windows.Forms.Label
        Me.Label46 = New System.Windows.Forms.Label
        Me.LOGFPA24 = New System.Windows.Forms.TextBox
        Me.Label47 = New System.Windows.Forms.Label
        Me.EPIS24 = New System.Windows.Forms.TextBox
        Me.POL24 = New System.Windows.Forms.TextBox
        Me.Label48 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.LOGFPA9 = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.LOGFPA16 = New System.Windows.Forms.TextBox
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.Label58 = New System.Windows.Forms.Label
        Me.agofpa24_6 = New System.Windows.Forms.TextBox
        Me.Label59 = New System.Windows.Forms.Label
        Me.agoepis24_6 = New System.Windows.Forms.TextBox
        Me.ago24_6 = New System.Windows.Forms.TextBox
        Me.Label60 = New System.Windows.Forms.Label
        Me.agoepis0 = New System.Windows.Forms.TextBox
        Me.Label42 = New System.Windows.Forms.Label
        Me.Label27 = New System.Windows.Forms.Label
        Me.agofpa9 = New System.Windows.Forms.TextBox
        Me.ago0 = New System.Windows.Forms.TextBox
        Me.Label28 = New System.Windows.Forms.Label
        Me.Label45 = New System.Windows.Forms.Label
        Me.agofpa16 = New System.Windows.Forms.TextBox
        Me.Label29 = New System.Windows.Forms.Label
        Me.agoepis9 = New System.Windows.Forms.TextBox
        Me.Label30 = New System.Windows.Forms.Label
        Me.agoepis16 = New System.Windows.Forms.TextBox
        Me.ago9 = New System.Windows.Forms.TextBox
        Me.ago16 = New System.Windows.Forms.TextBox
        Me.Label31 = New System.Windows.Forms.Label
        Me.Label32 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.agofpa13 = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.agofpa23 = New System.Windows.Forms.TextBox
        Me.Label23 = New System.Windows.Forms.Label
        Me.agoepis13 = New System.Windows.Forms.TextBox
        Me.Label24 = New System.Windows.Forms.Label
        Me.agoepis23 = New System.Windows.Forms.TextBox
        Me.ago13 = New System.Windows.Forms.TextBox
        Me.ago23 = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.Label26 = New System.Windows.Forms.Label
        Me.ApoSeira = New System.Windows.Forms.TextBox
        Me.Label33 = New System.Windows.Forms.Label
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.Label35 = New System.Windows.Forms.Label
        Me.nPistLian = New System.Windows.Forms.TextBox
        Me.Label36 = New System.Windows.Forms.Label
        Me.nPistTim = New System.Windows.Forms.TextBox
        Me.Label37 = New System.Windows.Forms.Label
        Me.nLian = New System.Windows.Forms.TextBox
        Me.Label38 = New System.Windows.Forms.Label
        Me.nTimol = New System.Windows.Forms.TextBox
        Me.Label39 = New System.Windows.Forms.Label
        Me.cPistLian = New System.Windows.Forms.TextBox
        Me.cPistTim = New System.Windows.Forms.TextBox
        Me.cLian = New System.Windows.Forms.TextBox
        Me.cTimol = New System.Windows.Forms.TextBox
        Me.Label40 = New System.Windows.Forms.Label
        Me.agoresB = New System.Windows.Forms.Button
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.mercury = New System.Windows.Forms.Button
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.apo = New System.Windows.Forms.DateTimePicker
        Me.eos = New System.Windows.Forms.DateTimePicker
        Me.Label34 = New System.Windows.Forms.Label
        Me.Label41 = New System.Windows.Forms.Label
        Me.cPistAg = New System.Windows.Forms.TextBox
        Me.cTimAg = New System.Windows.Forms.TextBox
        Me.Label43 = New System.Windows.Forms.Label
        Me.nPistAg = New System.Windows.Forms.TextBox
        Me.Label44 = New System.Windows.Forms.Label
        Me.nTimAg = New System.Windows.Forms.TextBox
        Me.Button7 = New System.Windows.Forms.Button
        Me.epan = New System.Windows.Forms.CheckBox
        Me.cParox = New System.Windows.Forms.TextBox
        Me.Label55 = New System.Windows.Forms.Label
        Me.nParox = New System.Windows.Forms.TextBox
        Me.Label56 = New System.Windows.Forms.Label
        Me.EPISPAR24 = New System.Windows.Forms.TextBox
        Me.PAR24 = New System.Windows.Forms.TextBox
        Me.Label57 = New System.Windows.Forms.Label
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(576, 14)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(65, 26)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'EXCELTOXML
        '
        Me.EXCELTOXML.Location = New System.Drawing.Point(577, 75)
        Me.EXCELTOXML.Name = "EXCELTOXML"
        Me.EXCELTOXML.Size = New System.Drawing.Size(133, 26)
        Me.EXCELTOXML.TabIndex = 1
        Me.EXCELTOXML.Text = "EXCELTOXML"
        Me.EXCELTOXML.UseVisualStyleBackColor = True
        Me.EXCELTOXML.Visible = False
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(15, 44)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(193, 22)
        Me.TextBox1.TabIndex = 3
        '
        'CD1
        '
        Me.CD1.FileName = "OpenFileDialog1"
        '
        'XMLTEXTWRITER
        '
        Me.XMLTEXTWRITER.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.XMLTEXTWRITER.Location = New System.Drawing.Point(402, 70)
        Me.XMLTEXTWRITER.Name = "XMLTEXTWRITER"
        Me.XMLTEXTWRITER.Size = New System.Drawing.Size(122, 36)
        Me.XMLTEXTWRITER.TabIndex = 4
        Me.XMLTEXTWRITER.Text = "KEF4 MERCURY ENTERSOFT"
        Me.XMLTEXTWRITER.UseVisualStyleBackColor = False
        '
        'xmlFin
        '
        Me.xmlFin.BackColor = System.Drawing.Color.Lime
        Me.xmlFin.Location = New System.Drawing.Point(920, 136)
        Me.xmlFin.Name = "xmlFin"
        Me.xmlFin.Size = New System.Drawing.Size(30, 14)
        Me.xmlFin.TabIndex = 5
        Me.xmlFin.Text = "Δημιουργία XML Β κατηγ"
        Me.xmlFin.UseVisualStyleBackColor = False
        Me.xmlFin.Visible = False
        '
        'bres_file
        '
        Me.bres_file.Location = New System.Drawing.Point(211, 43)
        Me.bres_file.Name = "bres_file"
        Me.bres_file.Size = New System.Drawing.Size(96, 20)
        Me.bres_file.TabIndex = 6
        Me.bres_file.Text = "Εύρεση Excel"
        Me.bres_file.UseVisualStyleBackColor = True
        '
        'pol23
        '
        Me.pol23.Location = New System.Drawing.Point(94, 31)
        Me.pol23.Name = "pol23"
        Me.pol23.Size = New System.Drawing.Size(101, 22)
        Me.pol23.TabIndex = 7
        Me.pol23.Text = "70-0057"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 36)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(82, 17)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Πωλήσ 23%"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 6)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(86, 17)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Πωλήσ  13%"
        '
        'pol13
        '
        Me.pol13.Location = New System.Drawing.Point(94, 3)
        Me.pol13.Name = "pol13"
        Me.pol13.Size = New System.Drawing.Size(101, 22)
        Me.pol13.TabIndex = 9
        Me.pol13.Text = "70-0036"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(199, 65)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(90, 17)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Λογ.Πελάτες"
        '
        'pel30
        '
        Me.pel30.Location = New System.Drawing.Point(296, 62)
        Me.pel30.Name = "pel30"
        Me.pel30.Size = New System.Drawing.Size(101, 22)
        Me.pel30.TabIndex = 11
        Me.pel30.Text = "30-0000"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(203, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(75, 17)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "Επιστ 13%"
        '
        'EPIS13
        '
        Me.EPIS13.Location = New System.Drawing.Point(293, 1)
        Me.EPIS13.Name = "EPIS13"
        Me.EPIS13.Size = New System.Drawing.Size(101, 22)
        Me.EPIS13.TabIndex = 15
        Me.EPIS13.Text = "70-0036"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(201, 34)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(75, 17)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "Επιστ 23%"
        '
        'EPIS23
        '
        Me.EPIS23.Location = New System.Drawing.Point(293, 29)
        Me.EPIS23.Name = "EPIS23"
        Me.EPIS23.Size = New System.Drawing.Size(101, 22)
        Me.EPIS23.TabIndex = 13
        Me.EPIS23.Text = "70-0057"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(647, 14)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(63, 27)
        Me.Button2.TabIndex = 17
        Me.Button2.Text = "DEMOXML"
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(201, 43)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(95, 17)
        Me.Label6.TabIndex = 25
        Me.Label6.Text = "Επισ.Λιαν13%"
        '
        'episLian13
        '
        Me.episLian13.Location = New System.Drawing.Point(296, 36)
        Me.episLian13.Name = "episLian13"
        Me.episLian13.Size = New System.Drawing.Size(101, 22)
        Me.episLian13.TabIndex = 24
        Me.episLian13.Text = "70-0036"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(201, 17)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(95, 17)
        Me.Label7.TabIndex = 23
        Me.Label7.Text = "Επισ.Λιαν23%"
        '
        'episLian23
        '
        Me.episLian23.Location = New System.Drawing.Point(296, 10)
        Me.episLian23.Name = "episLian23"
        Me.episLian23.Size = New System.Drawing.Size(101, 22)
        Me.episLian23.TabIndex = 22
        Me.episLian23.Text = "70-0057"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(2, 43)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(98, 17)
        Me.Label8.TabIndex = 21
        Me.Label8.Text = "Λιανικες  13%"
        '
        'Lian13
        '
        Me.Lian13.Location = New System.Drawing.Point(94, 36)
        Me.Lian13.Name = "Lian13"
        Me.Lian13.Size = New System.Drawing.Size(101, 22)
        Me.Lian13.TabIndex = 20
        Me.Lian13.Text = "70-0036"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(2, 17)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(94, 17)
        Me.Label9.TabIndex = 19
        Me.Label9.Text = "Λιανικές 23%"
        '
        'Lian23
        '
        Me.Lian23.Location = New System.Drawing.Point(94, 10)
        Me.Lian23.Name = "Lian23"
        Me.Lian23.Size = New System.Drawing.Size(101, 22)
        Me.Lian23.TabIndex = 18
        Me.Lian23.Text = "70-0057"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(577, 42)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(133, 28)
        Me.Button3.TabIndex = 26
        Me.Button3.Text = "LOADXML"
        Me.Button3.UseVisualStyleBackColor = True
        Me.Button3.Visible = False
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(668, 581)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(187, 44)
        Me.Button4.TabIndex = 27
        Me.Button4.Text = "Αποθήκευση  Αρχείου Λογ/σμών"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(668, 659)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(187, 34)
        Me.Button5.TabIndex = 28
        Me.Button5.Text = "Γραμμογράφηση Excel"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(7, 91)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(78, 17)
        Me.Label10.TabIndex = 33
        Me.Label10.Text = "Πωλήσ  9%"
        '
        'POL9
        '
        Me.POL9.Location = New System.Drawing.Point(94, 85)
        Me.POL9.Name = "POL9"
        Me.POL9.Size = New System.Drawing.Size(101, 22)
        Me.POL9.TabIndex = 32
        Me.POL9.Text = "70-0036"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(7, 66)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(86, 17)
        Me.Label11.TabIndex = 31
        Me.Label11.Text = "Πωλήσ  16%"
        '
        'POL16
        '
        Me.POL16.Location = New System.Drawing.Point(94, 59)
        Me.POL16.Name = "POL16"
        Me.POL16.Size = New System.Drawing.Size(101, 22)
        Me.POL16.TabIndex = 30
        Me.POL16.Text = "70-0057"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(7, 118)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(78, 17)
        Me.Label12.TabIndex = 35
        Me.Label12.Text = "Πωλήσ  0%"
        '
        'POL0
        '
        Me.POL0.Location = New System.Drawing.Point(94, 111)
        Me.POL0.Name = "POL0"
        Me.POL0.Size = New System.Drawing.Size(101, 22)
        Me.POL0.TabIndex = 34
        Me.POL0.Text = "70-0000"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(3, 65)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(88, 17)
        Me.Label13.TabIndex = 37
        Me.Label13.Text = "Λιανικες  0%"
        '
        'LIAN0
        '
        Me.LIAN0.Location = New System.Drawing.Point(94, 62)
        Me.LIAN0.Name = "LIAN0"
        Me.LIAN0.Size = New System.Drawing.Size(101, 22)
        Me.LIAN0.TabIndex = 36
        Me.LIAN0.Text = "70-0000"
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(578, 107)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(78, 41)
        Me.Button6.TabIndex = 38
        Me.Button6.Text = "ENTERSOFT"
        Me.Button6.UseVisualStyleBackColor = True
        Me.Button6.Visible = False
        '
        'Button8
        '
        Me.Button8.Location = New System.Drawing.Point(668, 549)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(187, 26)
        Me.Button8.TabIndex = 40
        Me.Button8.Text = "Ανάκληση Αρχείου Λογ/σμών"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'filexml
        '
        Me.filexml.Location = New System.Drawing.Point(668, 631)
        Me.filexml.Name = "filexml"
        Me.filexml.Size = New System.Drawing.Size(187, 22)
        Me.filexml.TabIndex = 41
        Me.filexml.Text = "GAT.XML"
        '
        'xmlG
        '
        Me.xmlG.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.xmlG.Location = New System.Drawing.Point(211, 119)
        Me.xmlG.Name = "xmlG"
        Me.xmlG.Size = New System.Drawing.Size(168, 31)
        Me.xmlG.TabIndex = 42
        Me.xmlG.Text = "Δημιουργία XML Γ κατηγ"
        Me.xmlG.UseVisualStyleBackColor = False
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(412, 3)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(98, 17)
        Me.Label14.TabIndex = 46
        Me.Label14.Text = "Λογ.ΦΠΑ 13%"
        '
        'LOGFPA13
        '
        Me.LOGFPA13.Location = New System.Drawing.Point(511, 3)
        Me.LOGFPA13.Name = "LOGFPA13"
        Me.LOGFPA13.Size = New System.Drawing.Size(101, 22)
        Me.LOGFPA13.TabIndex = 45
        Me.LOGFPA13.Text = "70-0036"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(412, 34)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(98, 17)
        Me.Label15.TabIndex = 44
        Me.Label15.Text = "Λογ.ΦΠΑ 23%"
        '
        'LOGFPA23
        '
        Me.LOGFPA23.Location = New System.Drawing.Point(511, 31)
        Me.LOGFPA23.Name = "LOGFPA23"
        Me.LOGFPA23.Size = New System.Drawing.Size(101, 22)
        Me.LOGFPA23.TabIndex = 43
        Me.LOGFPA23.Text = "70-0057"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(414, 41)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(102, 17)
        Me.Label16.TabIndex = 50
        Me.Label16.Text = "Λογ.ΦΠΑ  13%"
        '
        'lianLOGFPA13
        '
        Me.lianLOGFPA13.Location = New System.Drawing.Point(517, 36)
        Me.lianLOGFPA13.Name = "lianLOGFPA13"
        Me.lianLOGFPA13.Size = New System.Drawing.Size(101, 22)
        Me.lianLOGFPA13.TabIndex = 49
        Me.lianLOGFPA13.Text = "70-0036"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(415, 16)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(102, 17)
        Me.Label17.TabIndex = 48
        Me.Label17.Text = "Λογ.ΦΠΑ  23%"
        '
        'lianLOGFPA23
        '
        Me.lianLOGFPA23.Location = New System.Drawing.Point(517, 10)
        Me.lianLOGFPA23.Name = "lianLOGFPA23"
        Me.lianLOGFPA23.Size = New System.Drawing.Size(101, 22)
        Me.lianLOGFPA23.TabIndex = 47
        Me.lianLOGFPA23.Text = "70-0057"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Label61)
        Me.Panel1.Controls.Add(Me.arTam)
        Me.Panel1.Controls.Add(Me.lianLOGFPA23)
        Me.Panel1.Controls.Add(Me.lianLOGFPA13)
        Me.Panel1.Controls.Add(Me.Lian13)
        Me.Panel1.Controls.Add(Me.Lian23)
        Me.Panel1.Controls.Add(Me.Label52)
        Me.Panel1.Controls.Add(Me.lianLOGFPA24)
        Me.Panel1.Controls.Add(Me.Label53)
        Me.Panel1.Controls.Add(Me.episLian24)
        Me.Panel1.Controls.Add(Me.Label54)
        Me.Panel1.Controls.Add(Me.lian24)
        Me.Panel1.Controls.Add(Me.Label16)
        Me.Panel1.Controls.Add(Me.Label17)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.episLian13)
        Me.Panel1.Controls.Add(Me.LIAN0)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.episLian23)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.prom50)
        Me.Panel1.Controls.Add(Me.Label20)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.pel30)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Location = New System.Drawing.Point(17, 389)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(641, 156)
        Me.Panel1.TabIndex = 51
        '
        'Label61
        '
        Me.Label61.AutoSize = True
        Me.Label61.Location = New System.Drawing.Point(0, 126)
        Me.Label61.Name = "Label61"
        Me.Label61.Size = New System.Drawing.Size(95, 17)
        Me.Label61.TabIndex = 58
        Me.Label61.Text = "Αρ.Ταμειακής"
        '
        'arTam
        '
        Me.arTam.Location = New System.Drawing.Point(94, 123)
        Me.arTam.Name = "arTam"
        Me.arTam.Size = New System.Drawing.Size(101, 22)
        Me.arTam.TabIndex = 57
        '
        'Label52
        '
        Me.Label52.AutoSize = True
        Me.Label52.Location = New System.Drawing.Point(415, 93)
        Me.Label52.Name = "Label52"
        Me.Label52.Size = New System.Drawing.Size(102, 17)
        Me.Label52.TabIndex = 56
        Me.Label52.Text = "Λογ.ΦΠΑ  24%"
        '
        'lianLOGFPA24
        '
        Me.lianLOGFPA24.Location = New System.Drawing.Point(517, 93)
        Me.lianLOGFPA24.Name = "lianLOGFPA24"
        Me.lianLOGFPA24.Size = New System.Drawing.Size(101, 22)
        Me.lianLOGFPA24.TabIndex = 55
        Me.lianLOGFPA24.Text = "70-0087"
        '
        'Label53
        '
        Me.Label53.AutoSize = True
        Me.Label53.Location = New System.Drawing.Point(201, 93)
        Me.Label53.Name = "Label53"
        Me.Label53.Size = New System.Drawing.Size(95, 17)
        Me.Label53.TabIndex = 54
        Me.Label53.Text = "Επισ.Λιαν24%"
        '
        'episLian24
        '
        Me.episLian24.Location = New System.Drawing.Point(296, 93)
        Me.episLian24.Name = "episLian24"
        Me.episLian24.Size = New System.Drawing.Size(101, 22)
        Me.episLian24.TabIndex = 53
        Me.episLian24.Text = "70-0087"
        '
        'Label54
        '
        Me.Label54.AutoSize = True
        Me.Label54.Location = New System.Drawing.Point(0, 96)
        Me.Label54.Name = "Label54"
        Me.Label54.Size = New System.Drawing.Size(92, 17)
        Me.Label54.TabIndex = 52
        Me.Label54.Text = "Λιανικές 24%"
        '
        'lian24
        '
        Me.lian24.Location = New System.Drawing.Point(94, 93)
        Me.lian24.Name = "lian24"
        Me.lian24.Size = New System.Drawing.Size(101, 22)
        Me.lian24.TabIndex = 51
        Me.lian24.Text = "70-0087"
        '
        'prom50
        '
        Me.prom50.Location = New System.Drawing.Point(517, 65)
        Me.prom50.Name = "prom50"
        Me.prom50.Size = New System.Drawing.Size(101, 22)
        Me.prom50.TabIndex = 13
        Me.prom50.Text = "30-0000"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(427, 68)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(86, 17)
        Me.Label20.TabIndex = 14
        Me.Label20.Text = "Λογ.Προμηθ"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.Label49)
        Me.Panel3.Controls.Add(Me.LOGFPA17)
        Me.Panel3.Controls.Add(Me.Label50)
        Me.Panel3.Controls.Add(Me.EPIS17)
        Me.Panel3.Controls.Add(Me.POL17)
        Me.Panel3.Controls.Add(Me.Label51)
        Me.Panel3.Controls.Add(Me.Label46)
        Me.Panel3.Controls.Add(Me.LOGFPA24)
        Me.Panel3.Controls.Add(Me.Label47)
        Me.Panel3.Controls.Add(Me.EPIS24)
        Me.Panel3.Controls.Add(Me.POL24)
        Me.Panel3.Controls.Add(Me.Label48)
        Me.Panel3.Controls.Add(Me.Label14)
        Me.Panel3.Controls.Add(Me.Label18)
        Me.Panel3.Controls.Add(Me.LOGFPA13)
        Me.Panel3.Controls.Add(Me.LOGFPA9)
        Me.Panel3.Controls.Add(Me.Label15)
        Me.Panel3.Controls.Add(Me.Label19)
        Me.Panel3.Controls.Add(Me.LOGFPA23)
        Me.Panel3.Controls.Add(Me.LOGFPA16)
        Me.Panel3.Controls.Add(Me.Label4)
        Me.Panel3.Controls.Add(Me.POL0)
        Me.Panel3.Controls.Add(Me.EPIS13)
        Me.Panel3.Controls.Add(Me.Label5)
        Me.Panel3.Controls.Add(Me.POL16)
        Me.Panel3.Controls.Add(Me.EPIS23)
        Me.Panel3.Controls.Add(Me.Label11)
        Me.Panel3.Controls.Add(Me.POL9)
        Me.Panel3.Controls.Add(Me.pol13)
        Me.Panel3.Controls.Add(Me.pol23)
        Me.Panel3.Controls.Add(Me.Label10)
        Me.Panel3.Controls.Add(Me.Label1)
        Me.Panel3.Controls.Add(Me.Label12)
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Location = New System.Drawing.Point(24, 162)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(641, 193)
        Me.Panel3.TabIndex = 53
        '
        'Label49
        '
        Me.Label49.AutoSize = True
        Me.Label49.Location = New System.Drawing.Point(412, 163)
        Me.Label49.Name = "Label49"
        Me.Label49.Size = New System.Drawing.Size(98, 17)
        Me.Label49.TabIndex = 66
        Me.Label49.Text = "Λογ.ΦΠΑ 17%"
        '
        'LOGFPA17
        '
        Me.LOGFPA17.Location = New System.Drawing.Point(511, 163)
        Me.LOGFPA17.Name = "LOGFPA17"
        Me.LOGFPA17.Size = New System.Drawing.Size(101, 22)
        Me.LOGFPA17.TabIndex = 65
        Me.LOGFPA17.Text = "70-0036"
        '
        'Label50
        '
        Me.Label50.AutoSize = True
        Me.Label50.Location = New System.Drawing.Point(207, 163)
        Me.Label50.Name = "Label50"
        Me.Label50.Size = New System.Drawing.Size(75, 17)
        Me.Label50.TabIndex = 64
        Me.Label50.Text = "Επιστ 17%"
        '
        'EPIS17
        '
        Me.EPIS17.Location = New System.Drawing.Point(293, 163)
        Me.EPIS17.Name = "EPIS17"
        Me.EPIS17.Size = New System.Drawing.Size(101, 22)
        Me.EPIS17.TabIndex = 63
        Me.EPIS17.Text = "70-0036"
        '
        'POL17
        '
        Me.POL17.Location = New System.Drawing.Point(94, 163)
        Me.POL17.Name = "POL17"
        Me.POL17.Size = New System.Drawing.Size(101, 22)
        Me.POL17.TabIndex = 61
        Me.POL17.Text = "70-0036"
        '
        'Label51
        '
        Me.Label51.AutoSize = True
        Me.Label51.Location = New System.Drawing.Point(7, 163)
        Me.Label51.Name = "Label51"
        Me.Label51.Size = New System.Drawing.Size(86, 17)
        Me.Label51.TabIndex = 62
        Me.Label51.Text = "Πωλήσ  17%"
        '
        'Label46
        '
        Me.Label46.AutoSize = True
        Me.Label46.Location = New System.Drawing.Point(412, 140)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(98, 17)
        Me.Label46.TabIndex = 60
        Me.Label46.Text = "Λογ.ΦΠΑ 24%"
        '
        'LOGFPA24
        '
        Me.LOGFPA24.Location = New System.Drawing.Point(511, 137)
        Me.LOGFPA24.Name = "LOGFPA24"
        Me.LOGFPA24.Size = New System.Drawing.Size(101, 22)
        Me.LOGFPA24.TabIndex = 59
        Me.LOGFPA24.Text = "70-0087"
        '
        'Label47
        '
        Me.Label47.AutoSize = True
        Me.Label47.Location = New System.Drawing.Point(207, 140)
        Me.Label47.Name = "Label47"
        Me.Label47.Size = New System.Drawing.Size(75, 17)
        Me.Label47.TabIndex = 58
        Me.Label47.Text = "Επιστ 24%"
        '
        'EPIS24
        '
        Me.EPIS24.Location = New System.Drawing.Point(293, 137)
        Me.EPIS24.Name = "EPIS24"
        Me.EPIS24.Size = New System.Drawing.Size(101, 22)
        Me.EPIS24.TabIndex = 57
        Me.EPIS24.Text = "70-0087"
        '
        'POL24
        '
        Me.POL24.Location = New System.Drawing.Point(94, 137)
        Me.POL24.Name = "POL24"
        Me.POL24.Size = New System.Drawing.Size(101, 22)
        Me.POL24.TabIndex = 55
        Me.POL24.Text = "70-0087"
        '
        'Label48
        '
        Me.Label48.AutoSize = True
        Me.Label48.Location = New System.Drawing.Point(7, 140)
        Me.Label48.Name = "Label48"
        Me.Label48.Size = New System.Drawing.Size(86, 17)
        Me.Label48.TabIndex = 56
        Me.Label48.Text = "Πωλήσ  24%"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(415, 88)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(94, 17)
        Me.Label18.TabIndex = 54
        Me.Label18.Text = "Λογ.ΦΠΑ  9%"
        '
        'LOGFPA9
        '
        Me.LOGFPA9.Location = New System.Drawing.Point(511, 85)
        Me.LOGFPA9.Name = "LOGFPA9"
        Me.LOGFPA9.Size = New System.Drawing.Size(101, 22)
        Me.LOGFPA9.TabIndex = 53
        Me.LOGFPA9.Text = "70-0036"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(412, 62)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(98, 17)
        Me.Label19.TabIndex = 52
        Me.Label19.Text = "Λογ.ΦΠΑ 16%"
        '
        'LOGFPA16
        '
        Me.LOGFPA16.Location = New System.Drawing.Point(511, 59)
        Me.LOGFPA16.Name = "LOGFPA16"
        Me.LOGFPA16.Size = New System.Drawing.Size(101, 22)
        Me.LOGFPA16.TabIndex = 51
        Me.LOGFPA16.Text = "70-0057"
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel4.Controls.Add(Me.Label58)
        Me.Panel4.Controls.Add(Me.agofpa24_6)
        Me.Panel4.Controls.Add(Me.Label59)
        Me.Panel4.Controls.Add(Me.agoepis24_6)
        Me.Panel4.Controls.Add(Me.ago24_6)
        Me.Panel4.Controls.Add(Me.Label60)
        Me.Panel4.Controls.Add(Me.agoepis0)
        Me.Panel4.Controls.Add(Me.Label42)
        Me.Panel4.Controls.Add(Me.Label27)
        Me.Panel4.Controls.Add(Me.agofpa9)
        Me.Panel4.Controls.Add(Me.ago0)
        Me.Panel4.Controls.Add(Me.Label28)
        Me.Panel4.Controls.Add(Me.Label45)
        Me.Panel4.Controls.Add(Me.agofpa16)
        Me.Panel4.Controls.Add(Me.Label29)
        Me.Panel4.Controls.Add(Me.agoepis9)
        Me.Panel4.Controls.Add(Me.Label30)
        Me.Panel4.Controls.Add(Me.agoepis16)
        Me.Panel4.Controls.Add(Me.ago9)
        Me.Panel4.Controls.Add(Me.ago16)
        Me.Panel4.Controls.Add(Me.Label31)
        Me.Panel4.Controls.Add(Me.Label32)
        Me.Panel4.Controls.Add(Me.Label21)
        Me.Panel4.Controls.Add(Me.agofpa13)
        Me.Panel4.Controls.Add(Me.Label22)
        Me.Panel4.Controls.Add(Me.agofpa23)
        Me.Panel4.Controls.Add(Me.Label23)
        Me.Panel4.Controls.Add(Me.agoepis13)
        Me.Panel4.Controls.Add(Me.Label24)
        Me.Panel4.Controls.Add(Me.agoepis23)
        Me.Panel4.Controls.Add(Me.ago13)
        Me.Panel4.Controls.Add(Me.ago23)
        Me.Panel4.Controls.Add(Me.Label25)
        Me.Panel4.Controls.Add(Me.Label26)
        Me.Panel4.Location = New System.Drawing.Point(18, 540)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(640, 165)
        Me.Panel4.TabIndex = 54
        '
        'Label58
        '
        Me.Label58.AutoSize = True
        Me.Label58.Location = New System.Drawing.Point(405, 134)
        Me.Label58.Name = "Label58"
        Me.Label58.Size = New System.Drawing.Size(102, 17)
        Me.Label58.TabIndex = 105
        Me.Label58.Text = "Λογ.ΦΠΑ  24%"
        '
        'agofpa24_6
        '
        Me.agofpa24_6.Location = New System.Drawing.Point(521, 134)
        Me.agofpa24_6.Name = "agofpa24_6"
        Me.agofpa24_6.Size = New System.Drawing.Size(101, 22)
        Me.agofpa24_6.TabIndex = 104
        Me.agofpa24_6.Text = "70-0036"
        '
        'Label59
        '
        Me.Label59.AutoSize = True
        Me.Label59.Location = New System.Drawing.Point(209, 134)
        Me.Label59.Name = "Label59"
        Me.Label59.Size = New System.Drawing.Size(75, 17)
        Me.Label59.TabIndex = 103
        Me.Label59.Text = "Επιστ 24%"
        '
        'agoepis24_6
        '
        Me.agoepis24_6.Location = New System.Drawing.Point(299, 134)
        Me.agoepis24_6.Name = "agoepis24_6"
        Me.agoepis24_6.Size = New System.Drawing.Size(101, 22)
        Me.agoepis24_6.TabIndex = 102
        Me.agoepis24_6.Text = "70-0036"
        '
        'ago24_6
        '
        Me.ago24_6.Location = New System.Drawing.Point(96, 134)
        Me.ago24_6.Name = "ago24_6"
        Me.ago24_6.Size = New System.Drawing.Size(101, 22)
        Me.ago24_6.TabIndex = 100
        Me.ago24_6.Text = "70-0036"
        '
        'Label60
        '
        Me.Label60.AutoSize = True
        Me.Label60.Location = New System.Drawing.Point(8, 134)
        Me.Label60.Name = "Label60"
        Me.Label60.Size = New System.Drawing.Size(74, 17)
        Me.Label60.TabIndex = 101
        Me.Label60.Text = "Αγ.24%(6)"
        '
        'agoepis0
        '
        Me.agoepis0.Location = New System.Drawing.Point(298, 108)
        Me.agoepis0.Name = "agoepis0"
        Me.agoepis0.Size = New System.Drawing.Size(101, 22)
        Me.agoepis0.TabIndex = 98
        Me.agoepis0.Text = "20-3000"
        '
        'Label42
        '
        Me.Label42.AutoSize = True
        Me.Label42.Location = New System.Drawing.Point(206, 104)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(67, 17)
        Me.Label42.TabIndex = 99
        Me.Label42.Text = "Επιστ 9%"
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(411, 83)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(94, 17)
        Me.Label27.TabIndex = 70
        Me.Label27.Text = "Λογ.ΦΠΑ  9%"
        '
        'agofpa9
        '
        Me.agofpa9.Location = New System.Drawing.Point(521, 80)
        Me.agofpa9.Name = "agofpa9"
        Me.agofpa9.Size = New System.Drawing.Size(101, 22)
        Me.agofpa9.TabIndex = 69
        Me.agofpa9.Text = "70-0036"
        '
        'ago0
        '
        Me.ago0.Location = New System.Drawing.Point(96, 104)
        Me.ago0.Name = "ago0"
        Me.ago0.Size = New System.Drawing.Size(101, 22)
        Me.ago0.TabIndex = 96
        Me.ago0.Text = "20-3000"
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(405, 57)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(102, 17)
        Me.Label28.TabIndex = 68
        Me.Label28.Text = "Λογ.ΦΠΑ  16%"
        '
        'Label45
        '
        Me.Label45.AutoSize = True
        Me.Label45.Location = New System.Drawing.Point(5, 109)
        Me.Label45.Name = "Label45"
        Me.Label45.Size = New System.Drawing.Size(79, 17)
        Me.Label45.TabIndex = 97
        Me.Label45.Text = "Αγορές 0%"
        '
        'agofpa16
        '
        Me.agofpa16.Location = New System.Drawing.Point(521, 54)
        Me.agofpa16.Name = "agofpa16"
        Me.agofpa16.Size = New System.Drawing.Size(101, 22)
        Me.agofpa16.TabIndex = 67
        Me.agofpa16.Text = "70-0057"
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Location = New System.Drawing.Point(206, 83)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(67, 17)
        Me.Label29.TabIndex = 66
        Me.Label29.Text = "Επιστ 9%"
        '
        'agoepis9
        '
        Me.agoepis9.Location = New System.Drawing.Point(298, 80)
        Me.agoepis9.Name = "agoepis9"
        Me.agoepis9.Size = New System.Drawing.Size(101, 22)
        Me.agoepis9.TabIndex = 65
        Me.agoepis9.Text = "70-0036"
        '
        'Label30
        '
        Me.Label30.AutoSize = True
        Me.Label30.Location = New System.Drawing.Point(206, 59)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(75, 17)
        Me.Label30.TabIndex = 64
        Me.Label30.Text = "Επιστ 16%"
        '
        'agoepis16
        '
        Me.agoepis16.Location = New System.Drawing.Point(298, 54)
        Me.agoepis16.Name = "agoepis16"
        Me.agoepis16.Size = New System.Drawing.Size(101, 22)
        Me.agoepis16.TabIndex = 63
        Me.agoepis16.Text = "70-0057"
        '
        'ago9
        '
        Me.ago9.Location = New System.Drawing.Point(96, 80)
        Me.ago9.Name = "ago9"
        Me.ago9.Size = New System.Drawing.Size(101, 22)
        Me.ago9.TabIndex = 61
        Me.ago9.Text = "70-0036"
        '
        'ago16
        '
        Me.ago16.Location = New System.Drawing.Point(96, 54)
        Me.ago16.Name = "ago16"
        Me.ago16.Size = New System.Drawing.Size(101, 22)
        Me.ago16.TabIndex = 59
        Me.ago16.Text = "70-0057"
        '
        'Label31
        '
        Me.Label31.AutoSize = True
        Me.Label31.Location = New System.Drawing.Point(4, 59)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(91, 17)
        Me.Label31.TabIndex = 60
        Me.Label31.Text = "Αγορές  16%"
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Location = New System.Drawing.Point(3, 83)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(79, 17)
        Me.Label32.TabIndex = 62
        Me.Label32.Text = "Αγορές 9%"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(405, 34)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(102, 17)
        Me.Label21.TabIndex = 58
        Me.Label21.Text = "Λογ.ΦΠΑ  13%"
        '
        'agofpa13
        '
        Me.agofpa13.Location = New System.Drawing.Point(521, 29)
        Me.agofpa13.Name = "agofpa13"
        Me.agofpa13.Size = New System.Drawing.Size(101, 22)
        Me.agofpa13.TabIndex = 57
        Me.agofpa13.Text = "70-0036"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(405, 6)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(102, 17)
        Me.Label22.TabIndex = 56
        Me.Label22.Text = "Λογ.ΦΠΑ  23%"
        '
        'agofpa23
        '
        Me.agofpa23.Location = New System.Drawing.Point(521, 3)
        Me.agofpa23.Name = "agofpa23"
        Me.agofpa23.Size = New System.Drawing.Size(101, 22)
        Me.agofpa23.TabIndex = 55
        Me.agofpa23.Text = "70-0057"
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(206, 32)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(75, 17)
        Me.Label23.TabIndex = 54
        Me.Label23.Text = "Επιστ 13%"
        '
        'agoepis13
        '
        Me.agoepis13.Location = New System.Drawing.Point(298, 29)
        Me.agoepis13.Name = "agoepis13"
        Me.agoepis13.Size = New System.Drawing.Size(101, 22)
        Me.agoepis13.TabIndex = 53
        Me.agoepis13.Text = "70-0036"
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(206, 8)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(75, 17)
        Me.Label24.TabIndex = 52
        Me.Label24.Text = "Επιστ 23%"
        '
        'agoepis23
        '
        Me.agoepis23.Location = New System.Drawing.Point(298, 3)
        Me.agoepis23.Name = "agoepis23"
        Me.agoepis23.Size = New System.Drawing.Size(101, 22)
        Me.agoepis23.TabIndex = 51
        Me.agoepis23.Text = "70-0057"
        '
        'ago13
        '
        Me.ago13.Location = New System.Drawing.Point(96, 29)
        Me.ago13.Name = "ago13"
        Me.ago13.Size = New System.Drawing.Size(101, 22)
        Me.ago13.TabIndex = 49
        Me.ago13.Text = "70-0036"
        '
        'ago23
        '
        Me.ago23.Location = New System.Drawing.Point(96, 3)
        Me.ago23.Name = "ago23"
        Me.ago23.Size = New System.Drawing.Size(101, 22)
        Me.ago23.TabIndex = 47
        Me.ago23.Text = "70-0057"
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(4, 6)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(74, 17)
        Me.Label25.TabIndex = 48
        Me.Label25.Text = "Αγ.23%(2)"
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(5, 32)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(74, 17)
        Me.Label26.TabIndex = 50
        Me.Label26.Text = "Αγ.13%(1)"
        '
        'ApoSeira
        '
        Me.ApoSeira.Location = New System.Drawing.Point(115, 18)
        Me.ApoSeira.Name = "ApoSeira"
        Me.ApoSeira.Size = New System.Drawing.Size(90, 22)
        Me.ApoSeira.TabIndex = 55
        Me.ApoSeira.Text = "2"
        '
        'Label33
        '
        Me.Label33.AutoSize = True
        Me.Label33.Location = New System.Drawing.Point(12, 19)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(121, 17)
        Me.Label33.TabIndex = 56
        Me.Label33.Text = "Εναρξη από σειρά"
        '
        'ListBox1
        '
        Me.ListBox1.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(161, Byte))
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 17
        Me.ListBox1.Location = New System.Drawing.Point(668, 154)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(283, 157)
        Me.ListBox1.TabIndex = 57
        Me.ListBox1.UseWaitCursor = True
        '
        'Label35
        '
        Me.Label35.AutoSize = True
        Me.Label35.Location = New System.Drawing.Point(672, 457)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(100, 17)
        Me.Label35.TabIndex = 74
        Me.Label35.Text = "Επιστ.Λιανικης"
        '
        'nPistLian
        '
        Me.nPistLian.Location = New System.Drawing.Point(764, 454)
        Me.nPistLian.Name = "nPistLian"
        Me.nPistLian.Size = New System.Drawing.Size(30, 22)
        Me.nPistLian.TabIndex = 73
        Me.nPistLian.Text = "2"
        '
        'Label36
        '
        Me.Label36.AutoSize = True
        Me.Label36.Location = New System.Drawing.Point(672, 433)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(79, 17)
        Me.Label36.TabIndex = 72
        Me.Label36.Text = "Πιστωτ.Τιm"
        '
        'nPistTim
        '
        Me.nPistTim.Location = New System.Drawing.Point(764, 428)
        Me.nPistTim.Name = "nPistTim"
        Me.nPistTim.Size = New System.Drawing.Size(30, 22)
        Me.nPistTim.TabIndex = 71
        Me.nPistTim.Text = "2"
        '
        'Label37
        '
        Me.Label37.AutoSize = True
        Me.Label37.Location = New System.Drawing.Point(672, 406)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(80, 17)
        Me.Label37.TabIndex = 70
        Me.Label37.Text = "Δελτίο Λιαν"
        '
        'nLian
        '
        Me.nLian.Location = New System.Drawing.Point(764, 403)
        Me.nLian.Name = "nLian"
        Me.nLian.Size = New System.Drawing.Size(30, 22)
        Me.nLian.TabIndex = 69
        Me.nLian.Text = "2"
        '
        'Label38
        '
        Me.Label38.AutoSize = True
        Me.Label38.Location = New System.Drawing.Point(672, 382)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(69, 17)
        Me.Label38.TabIndex = 68
        Me.Label38.Text = "Τιμολογιο"
        '
        'nTimol
        '
        Me.nTimol.Location = New System.Drawing.Point(764, 377)
        Me.nTimol.Name = "nTimol"
        Me.nTimol.Size = New System.Drawing.Size(30, 22)
        Me.nTimol.TabIndex = 67
        Me.nTimol.Text = "2"
        '
        'Label39
        '
        Me.Label39.AutoSize = True
        Me.Label39.Location = New System.Drawing.Point(743, 334)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(76, 17)
        Me.Label39.TabIndex = 75
        Me.Label39.Text = "Αρ.Ψηφίων"
        '
        'cPistLian
        '
        Me.cPistLian.Location = New System.Drawing.Point(820, 454)
        Me.cPistLian.Name = "cPistLian"
        Me.cPistLian.Size = New System.Drawing.Size(30, 22)
        Me.cPistLian.TabIndex = 79
        Me.cPistLian.Text = "2"
        '
        'cPistTim
        '
        Me.cPistTim.Location = New System.Drawing.Point(820, 428)
        Me.cPistTim.Name = "cPistTim"
        Me.cPistTim.Size = New System.Drawing.Size(30, 22)
        Me.cPistTim.TabIndex = 78
        Me.cPistTim.Text = "2"
        '
        'cLian
        '
        Me.cLian.Location = New System.Drawing.Point(820, 403)
        Me.cLian.Name = "cLian"
        Me.cLian.Size = New System.Drawing.Size(30, 22)
        Me.cLian.TabIndex = 77
        Me.cLian.Text = "2"
        '
        'cTimol
        '
        Me.cTimol.Location = New System.Drawing.Point(820, 377)
        Me.cTimol.Name = "cTimol"
        Me.cTimol.Size = New System.Drawing.Size(30, 22)
        Me.cTimol.TabIndex = 76
        Me.cTimol.Text = "2"
        '
        'Label40
        '
        Me.Label40.AutoSize = True
        Me.Label40.Location = New System.Drawing.Point(808, 334)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(47, 17)
        Me.Label40.TabIndex = 80
        Me.Label40.Text = "Ψηφία"
        '
        'agoresB
        '
        Me.agoresB.BackColor = System.Drawing.Color.LightGreen
        Me.agoresB.Location = New System.Drawing.Point(211, 69)
        Me.agoresB.Name = "agoresB"
        Me.agoresB.Size = New System.Drawing.Size(168, 37)
        Me.agoresB.TabIndex = 81
        Me.agoresB.Text = "Δημιουργία XML Aγορ.Β κατηγ"
        Me.agoresB.UseVisualStyleBackColor = False
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(732, 12)
        Me.TextBox2.Multiline = True
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(218, 94)
        Me.TextBox2.TabIndex = 82
        Me.TextBox2.Text = "ΓΡΑΜΟΓΡΑΦΗΣΗ: AJ1=13%  AJ2=23%" & Global.Microsoft.VisualBasic.ChrW(9) & "AJ3=16%" & Global.Microsoft.VisualBasic.ChrW(9) & "AJ4=9%" & Global.Microsoft.VisualBasic.ChrW(9) & "AJ5=0%" & Global.Microsoft.VisualBasic.ChrW(9) & "AJI" & Global.Microsoft.VisualBasic.ChrW(9) & "FPA1" & Global.Microsoft.VisualBasic.ChrW(9) & "FPA2" & Global.Microsoft.VisualBasic.ChrW(9) & "FPA3" & Global.Microsoft.VisualBasic.ChrW(9) & "FPA4" & Global.Microsoft.VisualBasic.ChrW(9) & "ATIM" & _
            "" & Global.Microsoft.VisualBasic.ChrW(9) & "HME" & Global.Microsoft.VisualBasic.ChrW(9) & "EPO" & Global.Microsoft.VisualBasic.ChrW(9) & "AFM"""
        '
        'mercury
        '
        Me.mercury.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.mercury.Location = New System.Drawing.Point(17, 79)
        Me.mercury.Name = "mercury"
        Me.mercury.Size = New System.Drawing.Size(188, 25)
        Me.mercury.TabIndex = 83
        Me.mercury.Text = "Δημιουργία XML mercury B' κατηγ"
        Me.mercury.UseVisualStyleBackColor = False
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'apo
        '
        Me.apo.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.apo.Location = New System.Drawing.Point(100, 107)
        Me.apo.Name = "apo"
        Me.apo.Size = New System.Drawing.Size(105, 22)
        Me.apo.TabIndex = 84
        '
        'eos
        '
        Me.eos.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.eos.Location = New System.Drawing.Point(100, 134)
        Me.eos.Name = "eos"
        Me.eos.Size = New System.Drawing.Size(105, 22)
        Me.eos.TabIndex = 85
        '
        'Label34
        '
        Me.Label34.AutoSize = True
        Me.Label34.Location = New System.Drawing.Point(57, 113)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(32, 17)
        Me.Label34.TabIndex = 86
        Me.Label34.Text = "από"
        '
        'Label41
        '
        Me.Label41.AutoSize = True
        Me.Label41.Location = New System.Drawing.Point(56, 141)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(32, 17)
        Me.Label41.TabIndex = 87
        Me.Label41.Text = "έως"
        '
        'cPistAg
        '
        Me.cPistAg.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.cPistAg.Location = New System.Drawing.Point(820, 514)
        Me.cPistAg.Name = "cPistAg"
        Me.cPistAg.Size = New System.Drawing.Size(30, 22)
        Me.cPistAg.TabIndex = 94
        Me.cPistAg.Text = "D"
        '
        'cTimAg
        '
        Me.cTimAg.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.cTimAg.Location = New System.Drawing.Point(820, 488)
        Me.cTimAg.Name = "cTimAg"
        Me.cTimAg.Size = New System.Drawing.Size(30, 22)
        Me.cTimAg.TabIndex = 93
        Me.cTimAg.Text = "GgΞΠ"
        '
        'Label43
        '
        Me.Label43.AutoSize = True
        Me.Label43.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label43.Location = New System.Drawing.Point(672, 517)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(88, 17)
        Me.Label43.TabIndex = 91
        Me.Label43.Text = "Πιστ.Αγορών"
        '
        'nPistAg
        '
        Me.nPistAg.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.nPistAg.Location = New System.Drawing.Point(764, 514)
        Me.nPistAg.Name = "nPistAg"
        Me.nPistAg.Size = New System.Drawing.Size(30, 22)
        Me.nPistAg.TabIndex = 90
        Me.nPistAg.Text = "1"
        '
        'Label44
        '
        Me.Label44.AutoSize = True
        Me.Label44.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label44.Location = New System.Drawing.Point(672, 493)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(83, 17)
        Me.Label44.TabIndex = 89
        Me.Label44.Text = "Τιμολ. Αγορ"
        '
        'nTimAg
        '
        Me.nTimAg.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.nTimAg.Location = New System.Drawing.Point(764, 488)
        Me.nTimAg.Name = "nTimAg"
        Me.nTimAg.Size = New System.Drawing.Size(30, 22)
        Me.nTimAg.TabIndex = 88
        Me.nTimAg.Text = "1"
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(657, 107)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(53, 39)
        Me.Button7.TabIndex = 95
        Me.Button7.Text = "Εύρεση Excel"
        Me.Button7.UseVisualStyleBackColor = True
        Me.Button7.Visible = False
        '
        'epan
        '
        Me.epan.AutoSize = True
        Me.epan.Location = New System.Drawing.Point(16, 62)
        Me.epan.Name = "epan"
        Me.epan.Size = New System.Drawing.Size(132, 21)
        Me.epan.TabIndex = 96
        Me.epan.Text = "Επανενημέρωση"
        Me.epan.UseVisualStyleBackColor = True
        '
        'cParox
        '
        Me.cParox.Location = New System.Drawing.Point(820, 351)
        Me.cParox.Name = "cParox"
        Me.cParox.Size = New System.Drawing.Size(30, 22)
        Me.cParox.TabIndex = 99
        Me.cParox.Text = "2"
        '
        'Label55
        '
        Me.Label55.AutoSize = True
        Me.Label55.Location = New System.Drawing.Point(672, 351)
        Me.Label55.Name = "Label55"
        Me.Label55.Size = New System.Drawing.Size(85, 17)
        Me.Label55.TabIndex = 98
        Me.Label55.Text = "Παροχ.Υπηρ"
        '
        'nParox
        '
        Me.nParox.Location = New System.Drawing.Point(764, 351)
        Me.nParox.Name = "nParox"
        Me.nParox.Size = New System.Drawing.Size(30, 22)
        Me.nParox.TabIndex = 97
        Me.nParox.Text = "2"
        '
        'Label56
        '
        Me.Label56.AutoSize = True
        Me.Label56.Location = New System.Drawing.Point(227, 364)
        Me.Label56.Name = "Label56"
        Me.Label56.Size = New System.Drawing.Size(101, 17)
        Me.Label56.TabIndex = 103
        Me.Label56.Text = "Επ.Παροχ 24%"
        '
        'EPISPAR24
        '
        Me.EPISPAR24.Location = New System.Drawing.Point(319, 361)
        Me.EPISPAR24.Name = "EPISPAR24"
        Me.EPISPAR24.Size = New System.Drawing.Size(101, 22)
        Me.EPISPAR24.TabIndex = 102
        Me.EPISPAR24.Text = "70-0087"
        '
        'PAR24
        '
        Me.PAR24.Location = New System.Drawing.Point(120, 361)
        Me.PAR24.Name = "PAR24"
        Me.PAR24.Size = New System.Drawing.Size(101, 22)
        Me.PAR24.TabIndex = 100
        Me.PAR24.Text = "70-0087"
        '
        'Label57
        '
        Me.Label57.AutoSize = True
        Me.Label57.Location = New System.Drawing.Point(33, 364)
        Me.Label57.Name = "Label57"
        Me.Label57.Size = New System.Drawing.Size(84, 17)
        Me.Label57.TabIndex = 101
        Me.Label57.Text = "Παροχ  24%"
        '
        'main
        '
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(994, 765)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.mercury)
        Me.Controls.Add(Me.EPISPAR24)
        Me.Controls.Add(Me.Label56)
        Me.Controls.Add(Me.PAR24)
        Me.Controls.Add(Me.Label57)
        Me.Controls.Add(Me.cParox)
        Me.Controls.Add(Me.Label55)
        Me.Controls.Add(Me.nParox)
        Me.Controls.Add(Me.nPistLian)
        Me.Controls.Add(Me.epan)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.cPistAg)
        Me.Controls.Add(Me.cTimAg)
        Me.Controls.Add(Me.Label43)
        Me.Controls.Add(Me.nPistAg)
        Me.Controls.Add(Me.Label44)
        Me.Controls.Add(Me.nTimAg)
        Me.Controls.Add(Me.Label41)
        Me.Controls.Add(Me.Label34)
        Me.Controls.Add(Me.eos)
        Me.Controls.Add(Me.apo)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.agoresB)
        Me.Controls.Add(Me.Label40)
        Me.Controls.Add(Me.cPistLian)
        Me.Controls.Add(Me.cPistTim)
        Me.Controls.Add(Me.cLian)
        Me.Controls.Add(Me.cTimol)
        Me.Controls.Add(Me.Label39)
        Me.Controls.Add(Me.Label35)
        Me.Controls.Add(Me.Label36)
        Me.Controls.Add(Me.nPistTim)
        Me.Controls.Add(Me.Label37)
        Me.Controls.Add(Me.nLian)
        Me.Controls.Add(Me.Label38)
        Me.Controls.Add(Me.nTimol)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Label33)
        Me.Controls.Add(Me.ApoSeira)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.xmlG)
        Me.Controls.Add(Me.filexml)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.bres_file)
        Me.Controls.Add(Me.xmlFin)
        Me.Controls.Add(Me.XMLTEXTWRITER)
        Me.Controls.Add(Me.EXCELTOXML)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "main"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents EXCELTOXML As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents CD1 As System.Windows.Forms.OpenFileDialog


    Private Sub Form1_Load_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'χωρισ κωδικό
        'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MERCURY;Data Source=HP530\SQLEXPRESS
        'με κωδικό
        'gConnect = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;pwd=38983;Initial Catalog=MERCURY"
        'UPGRADE_WARNING: Couldn't resolve default property of object gConnect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        gConnect = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MERCURY;Data Source=HP530\SQLEXPRESS"
        'UPGRADE_WARNING: Couldn't resolve default property of object gConnect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ' gdb.Open(gConnect)
        'gdb.Execute("UPDATE EPSDETAIL SET FL_Ledg_Dscr='ΤΑΜΕΙΟ',FL_Ledg_cd='38-00-00-0000',VatVal='0.0000',NetVal='28.5000',RegVal='0.0000',MvTp='8',RegVatVal='0.0000'")
        'gdb.Execute("UPDATE EPSILON SET System_Dscr_1='Αγορές',Party_IDParty='60',APA_VIES_v_Dscr='EL',GlbCff='1',ExpenditureKind='0',AM_DcTp_Dscr='Τιμολόγιο,Αγοράς,',Party_AFM='82296964',ConstrCost='0',Party_ISK_D_A_Dscr='ΚΑΝΟΝΙΚΟΣ',dumm='0',AM_DcTp_cd='#ΤΑΓ-0',Party_ADDRESS='ΝΕΑ,ΜΠΑΦΡΑ,ΣΕΡΡΩΝ',Party_CASTVAT='1',AMO_Srl_DSCR='ΑΓΟΡΕΣ,(ΧΕΙΡΟΓΡΑΦΗ)',Base_dt='2014-04-03',System_sys='BP',Party_ISK_D_A_CD='0',F_Sites_dscr='ΚΕΝΤΡΙΚΟ',Party_DOY='5621',Party_JOB='ΕΜΠΟΡ,ΕΛΑΣΤΙΚ-ΓΕΩΡΓΙ',cdRetailIdentity='',Ledger_Supl='50-00-00-0000',AMO_Srl_cd='ΑΓ00',Party_Sts='1',Base_INVOICE='#ΤΑΓ-0/ΑΓ00/22319/',F_Sites_cd='001',IsHand='',Party_SNAME='Ι.ΘΕΟΔΩΡΙΔΗΣ,&,ΣΙΑ,Ο.Ε.',Party_CASTVAT_Dscr='ΚΑΝΟΝΙΚΟ',KepyoCatData_ISAGRYP='0',KepyoCatData_SUMKEPYOYP='23.1700',KepyoCatData_SUMKEPYOVAT='5.3300'")

        ' "GAT.XML"

        If File.Exists("C:\MERCVB\MERCURYRUN.TXT") Then
            mercury.Show()

            'File.Delete(Path)
            'LOAD_XML()
        End If
        If File.Exists(filexml.Text) Then
            'File.Delete(Path)
            LOAD_XML()
        End If



    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'WORKING FINE TEST MODE
        Dim a As String
        Dim K As Short
        Dim C As String

        ' Write the string as utf-8.
        ' This also writes the 3-byte utf-8 preamble at the beginning of the file.
        Dim appendMode As Boolean = False ' This overwrites the entire file.
        Dim sw As New StreamWriter("C:\MERCVB\out_utf8.XML", appendMode, System.Text.Encoding.UTF8)
        'sw.Write(TextBox1.Text)
        'sw.Close()









        Dim rH As New ADODB.Recordset
        rH.Open("select * from EPSILON", gdb, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        Dim rD As New ADODB.Recordset
        rD.Open("select * from EPSDETAIL", gdb, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        'UPGRADE_NOTE: enter was upgraded to enter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim enter_Renamed As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object enter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        enter_Renamed = ""

        FileOpen(1, "C:\MERCVB\A77.XML", OpenMode.Output)


        a = "<?xml version=""1.0""?>"
        'UPGRADE_WARNING: Couldn't resolve default property of object enter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        a = a + enter_Renamed + "<Data Style=""BRowse"" Name=""SX""><row rowId=""7"" mode=""3"" name=""Hd""><data><new "








        C = "" ' HEADER
        For K = 0 To rH.Fields.Count - 1
            C = C & rH.Fields(K).Name & "=""" + Replace(rH.Fields(K).Value, "&", "") + """ "
        Next
        a = a & C & "/></data>"


        PrintLine(1, a)
        'a = ""


        'UPGRADE_WARNING: Couldn't resolve default property of object enter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        a = a + enter_Renamed + "<detail><row rowId=""7"" mode=""3"" name=""Mv""><data><new "


        C = ""
        For K = 0 To rD.Fields.Count - 1
            C = C & rD.Fields(K).Name & "=""" + rD.Fields(K).Value + """ "
        Next

        'UPGRADE_WARNING: Couldn't resolve default property of object enter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        a = a + enter_Renamed + C + "/></data>"

        PrintLine(1, a)
        'a = ""


        'a = a + enter + "</data><detail><row rowId=""7"" mode=""3"" name=""Mv""><data><new"


        'σειρα με ολες τις βαθμιδες των λογαριασμών
        'UPGRADE_WARNING: Couldn't resolve default property of object enter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        a = a + enter_Renamed + "<detail><row rowId=""95"" mode=""3"" name=""Ledg""><data><new Active=""1"" dscrLedg=""ΤΑΜΕΙΟ"" cdLedg=""38-00-00-0000"" Anali=""0"" CanMv=""1""/></data></row><row rowId=""106"" mode=""3"" name=""Ledg""><data><new Active=""1"" dscrLedg=""ΤΑΜΕΙΟ"" cdLedg=""38-00-00"" Anali=""0"" CanMv=""0""/></data></row><row rowId=""117"" mode=""3"" name=""Ledg""><data><new Active=""1"" dscrLedg=""TAMEIO"" cdLedg=""38-00"" Anali=""0"" CanMv=""0""/></data></row><row rowId=""128"" mode=""3"" name=""Ledg""><data><new Active=""1"" dscrLedg=""ΧΡΗΜΑΤΙΚΑ ΔΙΑΘΕΣΙΜΑ"" cdLedg=""38"" Anali=""0"" CanMv=""0""/></data></row></detail>"

        'UPGRADE_WARNING: Couldn't resolve default property of object enter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        a = a + enter_Renamed + "</row></detail></row></Data>"





        '  Loop





        PrintLine(1, a)
        FileClose(1)


        sw.Write(a)
        sw.Close()




        'gdb.EXECUTE "UPDATE EPSILON SET System_Dscr_1='Αγορές',Party_IDParty='60',APA_VIES_v_Dscr='EL',GlbCff='1',ExpenditureKind='0',AM_DcTp_Dscr='Τιμολόγιο,Αγοράς,',Party_AFM='82296964',ConstrCost='0',Party_ISK_D_A_Dscr='ΚΑΝΟΝΙΚΟΣ',dumm='0',AM_DcTp_cd='#ΤΑΓ-0',Party_ADDRESS='ΝΕΑ,ΜΠΑΦΡΑ,ΣΕΡΡΩΝ',Party_CASTVAT='1',AMO_Srl_DSCR='ΑΓΟΡΕΣ,(ΧΕΙΡΟΓΡΑΦΗ)',Base_dt='2014-04-03',System_sys='BP',Party_ISK_D_A_CD='0',F_Sites_dscr='ΚΕΝΤΡΙΚΟ',Party_DOY='5621',Party_JOB='ΕΜΠΟΡ,ΕΛΑΣΤΙΚ-ΓΕΩΡΓΙ',cdRetailIdentity='',Ledger_Supl='50-00-00-0000',AMO_Srl_cd='ΑΓ00',Party_Sts='1',Base_INVOICE='#ΤΑΓ-0/ΑΓ00/22319/',F_Sites_cd='001',IsHand='',Party_SNAME='Ι.ΘΕΟΔΩΡΙΔΗΣ,&,ΣΙΑ,Ο.Ε.',Party_CASTVAT_Dscr='ΚΑΝΟΝΙΚΟ',KepyoCatData_ISAGRYP='0',KepyoCatData_SUMKEPYOYP='23.1700',KepyoCatData_SUMKEPYOVAT='5.3300'"
        MsgBox("ok")


        'gdb.EXECUTE "UPDATE EPSDETAIL SET FL_Ledg_Dscr='ΤΑΜΕΙΟ',FL_Ledg_cd='38-00-00-0000',VatVal='0.0000',NetVal='28.5000',RegVal='0.0000',MvTp='8',RegVatVal='0.0000'"


    End Sub


    Private Sub EXCELTOXML_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EXCELTOXML.Click
        Dim a As String
        'Dim K As Short
        Dim C As String

        ' Write the string as utf-8.
        ' This also writes the 3-byte utf-8 preamble at the beginning of the file.
        Dim appendMode As Boolean = False ' This overwrites the entire file.
        Dim sw As New StreamWriter("C:\MERCVB\out_utf9.export", appendMode, System.Text.Encoding.UTF8)
        'sw.Write(TextBox1.Text)
        'sw.Close()


        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xl As Excel.Worksheet

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open(TextBox1.Text)
        xl = xlWorkBook.Worksheets(1) ' .Add

        'xlwbook = xl.Workbooks.Open(TextBox1.Text)
        'xlsheet = xlwbook.Sheets.Item(1)





        Dim rH As New ADODB.Recordset
        rH.Open("select * from EPSILON", gdb, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        Dim rD As New ADODB.Recordset
        rD.Open("select * from EPSDETAIL", gdb, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        'UPGRADE_NOTE: enter was upgraded to enter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim enter_Renamed As String
        'UPGRADE_WARNING: Couldn't resolve default property of object enter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        enter_Renamed = Chr(13)

        FileOpen(1, "C:\MERCVB\A778.XML", OpenMode.Output)
        ROW = 1

        a = "<?xml version=""1.0""?>"
        'UPGRADE_WARNING: Couldn't resolve default property of object enter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        a = a + enter_Renamed + "<Data Style=""BRowse"" Name=""SX"">"

        '===============================================================================real onomatepvmymo 54100
        Do While True
            ROW = ROW + 1
            If IsDBNull(xl.Cells(ROW, 12).value) Then
                Exit Do
            End If

            If Len(xl.Cells(ROW, 11).ToString) < 2 Then
                Exit Do
            End If
            If xl.Cells(ROW, 11).value = Nothing Then
                Exit Do
            End If



            '' ΑΝ ΤΑ ΤΡΑΒΑΩ ΑΠΟ SQLSERVER
            'C = "" ' HEADER
            'For K = 0 To rH.Fields.Count - 1
            '    C = C & rH.Fields(K).Name & "=""" + Replace(rH.Fields(K).Value, "&", "") + """ "
            'Next

            ''''''''''''''''''''GEORGIADIS'''''''''''''''''tim1	AJ1-2	AJ2-3	AJ3-4	AJI-5	FPA1-6	FPA2-7	hmer-8	EPO-9	AFM-10	EPA-11	DIE-12	POL-13	PIST-14	tim-15


            '1	 2	    3	4	5	6	7	    8	    9	    10	    11	    12	13	14	15	16	17	    18	19
            'AJ1 AJ2	AJ3	AJ4	AJ5	AJI	FPA1	FPA2	FPA3	FPA4	ATIM	HME	EPO	AFM	KPE	DIE	XRVMA	EPA	POL
            Party_IDParty = xl.Cells(ROW, 14).value  ' As String  '12344   ΚΩΔ ΣΥΝΑΛΛΑΣΟΜΕΝΟΥ
            AM_DcTp_Dscr = "Τιμολόγιο"
            Party_AFM = Trim(xl.Cells(ROW, 14).value)  'Dim Party_AFM As String ' =""999349996
            Party_ADDRESS = xl.Cells(ROW, 16).value 'ToString  ' "ΠΟΛΥΣΤΗΛΟ ΚΑΒΑΛΑΣ"
            AM_DcTp_cd = "#ΤΥΠ-0"
            AMO_Srl_DSCR = "ΠΩΛΗΣΕΙΣ"
            Base_dt = VB6.Format(xl.Cells(ROW, 12), "YYYY-mm-dd")
            Base_INVOICE = xl.Cells(ROW, 11).value  ' =""#ΤΥΠ-0/Π000/1/Τιμολόγιο Παροχής Υπηρεσιών"
            Party_SNAME = xl.Cells(ROW, 13).value  '=""Θ. ΓΡΑΜΜΑΤΗΣ Κ.ΣΙΑ Ε.Ε""

            KAU_AJIA = nVal(xl.Cells(ROW, 1).value) + nVal(xl.Cells(ROW, 2).value) + nVal(xl.Cells(ROW, 3).value) + nVal(xl.Cells(ROW, 4).value) + nVal(xl.Cells(ROW, 5).value)
            FPA = nVal(xl.Cells(ROW, 7).value) + nVal(xl.Cells(ROW, 8).value) + nVal(xl.Cells(ROW, 9).value) + nVal(xl.Cells(ROW, 10).value)

            FL_Ledg_Dscr = "ΠΩΛΗΣΕΙΣ ΧΟΝΔΡΙΚΗΣ ΕΣ. ΦΠΑ23%"
            FL_Ledg_cd = "70-00-00-0057"
            KAU_AJIA1 = KAU_AJIA
            FPA1 = FPA





            C = "<row rowId=""" + LTrim(Str(rowId)) + """ mode=""3"" name=""Hd""><data><new "
            C = C + "System_Dscr_1=""Πωλήσεις"" Party_IDParty=""" + Party_IDParty + """ APA_VIES_v_Dscr=""EL"" GlbCff=""1"" ExpenditureKind=""0"" "
            C = C + "AM_DcTp_Dscr=""" + AM_DcTp_Dscr + """ Party_AFM=""" + Trim(Party_AFM) + """ ConstrCost=""0"" "
            C = C + "Party_ISK_D_A_Dscr="""" dumm=""0"" AM_DcTp_cd=""" + AM_DcTp_cd + """ Party_ADDRESS=""" + Party_ADDRESS + """ "
            a = a & C : sw.WriteLine(a) : a = "" : C = ""
            C = C + "Party_CASTVAT=""1"" AMO_Srl_DSCR=""" + AMO_Srl_DSCR + """ Base_dt=""" + Base_dt + """ System_sys=""SB"" "
            C = C + "F_Sites_dscr=""ΚΕΝΤΡΙΚΟ"" Party_DOY="""" cdRetailIdentity="""" AMO_Srl_cd=""Π000"" "
            C = C + "Party_Sts=""1"" Base_INVOICE=""" + Base_INVOICE + """ F_Sites_cd=""001"" "
            a = a & C : sw.WriteLine(a) : a = "" : C = ""
            C = C + "IsHand="""" Party_SNAME=""" + Party_SNAME + """ Party_CASTVAT_Dscr=""ΚΑΝΟΝΙΚΟ"" "
            C = C + "KepyoCatData_ISAGRYP=""0"" KepyoCatData_SUMKEPYOYP=""" + KAU_AJIA + """ KepyoCatData_SUMKEPYOVAT=""" + FPA + """ "
            C = C + "Ledger_Cust=""30-00-00-0000"""




            a = a & C & "/></data>"


            PrintLine(1, a)
            sw.WriteLine(a)
            ' sw.Write(a)
            a = ""



            a = a + enter_Renamed + "<detail><row rowId=""" + LTrim(Str(rowId)) + """ mode=""3"" name=""Mv""><data><new "


            'C = ""
            'For K = 0 To rD.Fields.Count - 1
            '    C = C & rD.Fields(K).Name & "=""" + rD.Fields(K).Value + """ "
            'Next



            'tim1	AJ1-2	AJ2-3	AJ3-4	AJI-5	FPA1-6	FPA2-7	hmer-8	EPO-9	AFM-10	EPA-11	DIE-12	POL-13	PIST-14	tim-15
            C = " FL_Ledg_Dscr=""" + FL_Ledg_Dscr + """ FL_Ledg_cd=""" + FL_Ledg_cd + """ "
            C = C + "VatVal=""" + FPA1 + """ NetVal=""" + KAU_AJIA1 + """ RegVal=""" + KAU_AJIA1 + """ MvTp=""1"" RegVatVal=""0.0000"""


            a = a + enter_Renamed + C + "/></data>"

            PrintLine(1, a)
            sw.WriteLine(a)
            'sw.Write(a)
            a = ""

            'σειρα με ολες τις βαθμιδες των λογαριασμών

            a = a + enter_Renamed + "<detail><row rowId=""" + LTrim(Str(rowIdINNER)) + """ mode=""3"" name=""Ledg""><data><new Active=""1"" dscrLedg=""Πωλήσεις 23%"" cdLedg=""70-00-00-0057"" Anali=""0"" CanMv=""1""/></data></row>"
            rowIdINNER = rowIdINNER + 11
            sw.WriteLine(a) : a = ""
            a = a + "<row rowId=""" + LTrim(Str(rowIdINNER)) + """ mode=""3"" name=""Ledg""><data><new Active=""1"" dscrLedg=""Πωλήσεις"" cdLedg=""70-00-00"" Anali=""0"" CanMv=""0""/></data></row>"
            rowIdINNER = rowIdINNER + 11
            sw.WriteLine(a) : a = ""
            a = a + "<row rowId=""" + LTrim(Str(rowIdINNER)) + """ mode=""3"" name=""Ledg""><data><new Active=""1"" dscrLedg=""Πωλήσεις"" cdLedg=""70-00"" Anali=""0"" CanMv=""0""/></data></row>"
            rowIdINNER = rowIdINNER + 11
            sw.WriteLine(a) : a = ""
            a = a + "<row rowId=""" + LTrim(Str(rowIdINNER)) + """ mode=""3"" name=""Ledg""><data><new Active=""1"" dscrLedg=""Πωλήσεις Εμπορευμάτων"" cdLedg=""70"" Anali=""0"" CanMv=""0""/></data></row>"
            sw.WriteLine(a) : a = ""
            a = a + "</detail>"
            a = a + enter_Renamed + "</row></detail></row>"
            sw.WriteLine(a) : a = ""


            rowId = rowId + 11
            rowIdINNER = rowIdINNER + 11


        Loop

        a = a + "</Data>"


        PrintLine(1, a)
        FileClose(1)


        sw.Write(a)
        sw.Close()




        'gdb.EXECUTE "UPDATE EPSILON SET System_Dscr_1='Αγορές',Party_IDParty='60',APA_VIES_v_Dscr='EL',GlbCff='1',ExpenditureKind='0',AM_DcTp_Dscr='Τιμολόγιο,Αγοράς,',Party_AFM='82296964',ConstrCost='0',Party_ISK_D_A_Dscr='ΚΑΝΟΝΙΚΟΣ',dumm='0',AM_DcTp_cd='#ΤΑΓ-0',Party_ADDRESS='ΝΕΑ,ΜΠΑΦΡΑ,ΣΕΡΡΩΝ',Party_CASTVAT='1',AMO_Srl_DSCR='ΑΓΟΡΕΣ,(ΧΕΙΡΟΓΡΑΦΗ)',Base_dt='2014-04-03',System_sys='BP',Party_ISK_D_A_CD='0',F_Sites_dscr='ΚΕΝΤΡΙΚΟ',Party_DOY='5621',Party_JOB='ΕΜΠΟΡ,ΕΛΑΣΤΙΚ-ΓΕΩΡΓΙ',cdRetailIdentity='',Ledger_Supl='50-00-00-0000',AMO_Srl_cd='ΑΓ00',Party_Sts='1',Base_INVOICE='#ΤΑΓ-0/ΑΓ00/22319/',F_Sites_cd='001',IsHand='',Party_SNAME='Ι.ΘΕΟΔΩΡΙΔΗΣ,&,ΣΙΑ,Ο.Ε.',Party_CASTVAT_Dscr='ΚΑΝΟΝΙΚΟ',KepyoCatData_ISAGRYP='0',KepyoCatData_SUMKEPYOYP='23.1700',KepyoCatData_SUMKEPYOVAT='5.3300'"
        MsgBox("ok")


        'gdb.EXECUTE "UPDATE EPSDETAIL SET FL_Ledg_Dscr='ΤΑΜΕΙΟ',FL_Ledg_cd='38-00-00-0000',VatVal='0.0000',NetVal='28.5000',RegVal='0.0000',MvTp='8',RegVatVal='0.0000'"



        xlApp.Quit()




    End Sub



    Private Sub XMLTEXTWRITER_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XMLTEXTWRITER.Click
        '       Imports System.Xml
        'Public Class Form1
        FORM3.Show()



        'Kill("c:\mercvb\product2.xml")


        ''  Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim writer As New XmlTextWriter("c:\mercvb\product2.xml", System.Text.Encoding.UTF8)
        'writer.WriteStartDocument(True)
        'writer.Formatting = Formatting.Indented
        'writer.Indentation = 2
        '' Create a Continent element and set its value to
        '' that of the New Continent dialog box
        ''writer.WriteAttributeString("Table", , "sadasasd sdsd")

        ''<Data Name="SX" Style="BRowse"><row name="Hd" mode="3" rowId="7">
        'writer.WriteStartElement("packages")
        'writer.WriteStartElement("package")
        'writer.WriteAttributeString("actor_afm", "SX")
        'writer.WriteAttributeString("month", "9")
        'writer.WriteAttributeString("year", "2014")


        'writer.WriteStartElement("groupedRevenues")
        'writer.WriteAttributeString("action", "replace")

        'writer.WriteStartElement("revenue")

        'writer.WriteStartElement("afm") : writer.WriteString("028783755") : writer.WriteEndElement()  'AFM
        'writer.WriteStartElement("amount") : writer.WriteString("028783755") : writer.WriteEndElement()  'AFM
        'writer.WriteStartElement("tax") : writer.WriteString("028783755") : writer.WriteEndElement()  'AFM
        'writer.WriteStartElement("invoices") : writer.WriteString("028783755") : writer.WriteEndElement()  'AFM

        'writer.WriteStartElement("note") : writer.WriteString("028783755") : writer.WriteEndElement()  'AFM
        'writer.WriteStartElement("date") : writer.WriteString("028783755") : writer.WriteEndElement()  'AFM





        'writer.WriteEndElement()  'REVENUE

        'writer.WriteEndElement()  'GROUPEDREVENUES




        ' writer.WriteEndElement()  'PACKAGE
        'writer.WriteEndElement()  'PACKAGES


        '     <groupedRevenues action="replace">
        '<revenue>
        '	<afm>090909099</afm>
        '	<amount>0</amount>
        '	<tax>0</tax>
        '	<invoices>1</invoices>
        '	<note>normal</note>
        '	<date>2014-01-01</date>
        '</revenue>







        Dim kke As Integer

        'For kke = 1 To 3
        '    writer.WriteStartElement("data")
        '    writer.WriteStartElement("new")

        '    writer.WriteAttributeString("CanMv", "canmove")
        '    writer.WriteAttributeString("Anali", "0")
        '    writer.WriteAttributeString("cdLedg", "log")
        '    writer.WriteAttributeString("dscrLedg", "πωλήσεις")
        '    writer.WriteAttributeString("Active", "1")





        '    writer.WriteEndElement()  'NEW
        '    writer.WriteEndElement()  'data



        '    '            write_row(writer)
        'Next

        '  write_row(writer)



        'createNode(1, "Product 1", "1000", writer)
        'createNode(2, "Product 2", "2000", writer)
        'createNode(3, "Product 3", "3000", writer)
        'createNode(4, "Product 4", "4000", writer)

        'writer.WriteEndDocument()
        'writer.Close()
        'write2()
        'MsgBox("ok")

    End Sub









    Private Sub xmlFin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles xmlFin.Click
        Dim a As String
        Dim K As Short
        Dim C As String

        ' CO TO DIAXORISTIKO DEKADIKON ARITMON
        Dim CO As String = String.Format(1.1).Substring(1, 1)


        MsgBox("ΠΡΟΣΟΧΗ ΔΙΑΒΑΖΕΙ ΑΠΟ ΤΗΝ 2η ΣΕΙΡΑ ΜΕ ΓΡΑΜΟΓΡΑΦΗΣΗ:" + Chr(13) + "AJ1 AJ2	AJ3	AJ4	AJ5	AJI	FPA1	FPA2	FPA3	FPA4	ATIM	HME	EPO	AFM")


        ' Write the string as utf-8.
        ' This also writes the 3-byte utf-8 preamble at the beginning of the file.
        Dim appendMode As Boolean = False ' This overwrites the entire file.
        ' Dim sw As New StreamWriter("C:\MERCVB\out_utf9.export", appendMode, System.Text.Encoding.UTF8)
        'sw.Write(TextBox1.Text)
        'sw.Close()


        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xl As Excel.Worksheet

        xlApp = New Excel.ApplicationClass


        If Len(TextBox1.Text) < 2 Then
            Exit Sub
        End If

        xlWorkBook = xlApp.Workbooks.Open(TextBox1.Text)
        xl = xlWorkBook.Worksheets(1) ' .Add




        '====================================================================================
        Dim ff As String = "c:\mercvb\m" + VB6.Format(Now, "YYYYddmmHHMM") + ".export" ' "\\Logisthrio\333\pr.export" '

        Dim writer As New XmlTextWriter(ff, System.Text.Encoding.UTF8)
        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2
        ' Create a Continent element and set its value to
        ' that of the New Continent dialog box
        'writer.WriteAttributeString("Table", , "sadasasd sdsd")

        '<Data Name="SX" Style="BRowse"><row name="Hd" mode="3" rowId="7">
        writer.WriteStartElement("Data")
        writer.WriteAttributeString("Name", "SX")
        writer.WriteAttributeString("Style", "Browse")
        '====================================================================================


        Dim enter_Renamed As String
        enter_Renamed = Chr(13)

        'FileOpen(1, "C:\MERCVB\A778.XML", OpenMode.Output)
        ROW = 1

        Dim hand As Integer = 0

        fnTimol = Val(nTimol.Text)
        fnLian = Val(nLian.Text)
        fnPistTim = Val(nPistTim.Text)
        fnPistLian = Val(nPistLian.Text)
        ' As Integer
        fcTimol = cTimol.Text
        fcLian = cLian.Text
        fcPistTim = cPistTim.Text
        fcPistLian = cPistLian.Text


        '===============================================================================real onomatepvmymo 54100
        Do While True
            ROW = ROW + 1

            Me.Text = ROW
            'system.doevents

            If IsDBNull(xl.Cells(ROW, 12).value) Then
                Exit Do
            End If

            If Len(xl.Cells(ROW, 11).ToString) < 2 Then
                Exit Do
            End If
            If xl.Cells(ROW, 11).value = Nothing Then
                Exit Do
            End If

            '1	 2	    3	4	5	6	7	    8	    9	    10	    11	    12	13	14	15	16	17	    18	19
            'AJ1 AJ2	AJ3	AJ4	AJ5	AJI	FPA1	FPA2	FPA3	FPA4	ATIM	HME	EPO	AFM	KPE	DIE	XRVMA	EPA	POL
            Party_IDParty = xl.Cells(ROW, 14).value  ' As String  '12344   ΚΩΔ ΣΥΝΑΛΛΑΣΟΜΕΝΟΥ
            AM_DcTp_Dscr = "Τιμολόγιο"
            Party_AFM = Trim(xl.Cells(ROW, 14).value)  'Dim Party_AFM As String ' =""999349996
            If Len(Trim(Party_AFM)) <= 4 Then
                Party_AFM = "000000000"
            End If

            Party_ADDRESS = xl.Cells(ROW, 16).value 'ToString  ' "ΠΟΛΥΣΤΗΛΟ ΚΑΒΑΛΑΣ"
            AM_DcTp_cd = "#ΤΥΠ-0"
            AMO_Srl_DSCR = "Πωλήσεις" '"ΠΩΛΗΣΕΙΣ"
            Base_dt = VB6.Format(xl.Cells(ROW, 12), "YYYY-mm-dd")
            Base_INVOICE = xl.Cells(ROW, 11).value  ' =""#ΤΥΠ-0/Π000/1/Τιμολόγιο Παροχής Υπηρεσιών"
            Party_SNAME = xl.Cells(ROW, 13).value  '=""Θ. ΓΡΑΜΜΑΤΗΣ Κ.ΣΙΑ Ε.Ε""
            f_logPel = pel30.Text ' "30-00-00-0000"

            KAU_AJIA = nVal(xl.Cells(ROW, 1).value) + nVal(xl.Cells(ROW, 2).value) + nVal(xl.Cells(ROW, 3).value) + nVal(xl.Cells(ROW, 4).value) + nVal(xl.Cells(ROW, 5).value)
            FPA = nVal(xl.Cells(ROW, 7).value) + nVal(xl.Cells(ROW, 8).value) + nVal(xl.Cells(ROW, 9).value) + nVal(xl.Cells(ROW, 10).value)


            kau13 = nVal(xl.Cells(ROW, 1).value)
            kau23 = nVal(xl.Cells(ROW, 2).value)
            kau16 = nVal(xl.Cells(ROW, 3).value)
            kau9 = nVal(xl.Cells(ROW, 4).value)
            kau0 = nVal(xl.Cells(ROW, 5).value)

            fpa13 = nVal(xl.Cells(ROW, 7).value)
            fpa23 = nVal(xl.Cells(ROW, 8).value)
            fpa16 = nVal(xl.Cells(ROW, 9).value)


            LOG13 = pol13.Text : LOG23 = pol23.Text
            LOG16 = POL16.Text : LOG9 = POL9.Text
            LOG0 = POL0.Text



            FL_Ledg_Dscr = "ΠΩΛΗΣΕΙΣ ΧΟΝΔΡΙΚΗΣ ΕΣ. ΦΠΑ23%"
            FL_Ledg_cd = pol23.Text ' "70-00-00-0057"

            MVTP = "1"  ' 2=agores  6=επιστροφεσ πολισεον 3=εισπαρ  4=πλιρομεσ
            System_sys = "SB" '      'SB =POLISEIS FR



            If InStr(fcLian, Mid(Base_INVOICE, 1, fnLian)) > 0 Then
                IsHand = "1" 'LTrim(Str(hand))
                cdRetailIdentity = ""
                LOG13 = Lian13.Text : LOG23 = Lian23.Text
                LOG0 = LIAN0.Text
                If InStr("yπ", Mid(Base_INVOICE, 1, 1)) > 0 Then 'yphresies
                    LOG23 = "73-4057"
                End If
                If InStr("πφ", Mid(Base_INVOICE, 1, 1)) > 0 Then 'yphresies
                    cdRetailIdentity = "ΣΥ09002067"
                    LOG23 = "73-4057"
                    IsHand = "0" 'LTrim(Str(hand))
                End If




            ElseIf InStr("GgΞμ", Mid(Base_INVOICE, 1, 1)) > 0 Then
                'θελει ψαξιμο.................

                IsHand = "" 'LTrim(Str(hand))  BP=AGORES
                cdRetailIdentity = ""
                LOG13 = ago13.Text : LOG23 = ago23.Text
                'LOG0 = LIAN0.Text
                System_sys = "BP"

                'System_Dscr_1
                AMO_Srl_DSCR = "Αγορές"
                AMO_Srl_DSCR = "ΑΓΟΡΕΣ (ΧΕΙΡΟΓΡΑΦΗ)"
                System_sys = "BP"



            ElseIf InStr(fcTimol, Mid(Base_INVOICE, 1, fnTimol)) > 0 Then 'τιμολογια -πιστωτικά
                cdRetailIdentity = ""
                IsHand = ""
                'LOG13 = pol13.Text : LOG23 = pol23.Text
                'LOG16 = POL16.Text : LOG9 = POL9.Text
                'LOG0 = POL0.Text

            ElseIf InStr("Y", Mid(Base_INVOICE, 1, 1)) > 0 Then 'yphresies
                LOG23 = "73-0057"
                'End If




                'ElseIf InStr("P", Mid(Base_INVOICE, 1, 1)) > 0 Then 'Mid(Base_INVOICE, 1, 1) = "P" Then
                '    LOG13 = EPIS13.Text : LOG23 = EPIS23.Text
                '    MVTP = 6
                'End If
            ElseIf InStr(fcPistLian, Mid(Base_INVOICE, 1, fnPistLian)) > 0 Then  'επιστροφη λιανικης
                LOG13 = episLian13.Text : LOG23 = episLian23.Text
                MVTP = 6
                IsHand = "1" 'LTrim(Str(hand))
                kau13 = -kau13
                kau23 = -kau23
                fpa13 = -fpa13
                fpa23 = -fpa23
                KAU_AJIA = -KAU_AJIA
                FPA = -FPA

                System_sys = "FR"

                'End If
            ElseIf InStr(fcPistTim, Mid(Base_INVOICE, 1, fnPistTim)) > 0 Then  'pistvtiko timologio
                kau13 = -kau13
                kau23 = -kau23
                fpa13 = -fpa13
                fpa23 = -fpa23
                KAU_AJIA = -KAU_AJIA
                FPA = -FPA
                LOG13 = EPIS13.Text : LOG23 = EPIS23.Text
                MVTP = 6
                'End If


            End If
            KAU_AJIA1 = KAU_AJIA
            FPA1 = FPA
            write_row(writer)
            rowId = rowId + 11
        Loop





        writer.WriteEndDocument()
        writer.Close()





        MsgBox("Δημιουργήθηκε στο " + ff)
        xlApp.Quit()

    End Sub

    '    End Sub
    Sub write_row(ByVal w As XmlTextWriter)
        'DDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDD

        'big row
        w.WriteStartElement("row") : w.WriteAttributeString("rowid", LTrim(Str(rowId))) : w.WriteAttributeString("mode", "3") : w.WriteAttributeString("name", "Hd")


        '========================================================
        w.WriteStartElement("data")
        w.WriteStartElement("new")

        w.WriteAttributeString("IsHand", IsHand)
        w.WriteAttributeString("F_Sites_cd", "001")
        w.WriteAttributeString("ConstrCost", "0")
        w.WriteAttributeString("Party_ISK_D_A_Dscr", "ΚΑΝΟΝΙΚΟΣ")
        w.WriteAttributeString("AM_DcTp_Dscr", AM_DcTp_Dscr)
        w.WriteAttributeString("GlbCff", "1")
        w.WriteAttributeString("Party_Sts", "1") ' neo
        w.WriteAttributeString("Party_IDParty", Mid(Party_AFM, 5, 5)) ' κωδικος πελατη 13 
        w.WriteAttributeString("AMO_Srl_cd", "πλ00")
        w.WriteAttributeString("KepyoCatData_SUMKEPYOVAT", Replace(Str(FPA1), ",", "."))  'FPA1)
        w.WriteAttributeString("Party_PHONE2", "")
        w.WriteAttributeString("Party_CASTVAT_Dscr", "ΚΑΝΟΝΙΚΟ")
        w.WriteAttributeString("cdRetailIdentity", cdRetailIdentity)
        w.WriteAttributeString("Ledger_Cust", pel30.Text)  'neo   
        w.WriteAttributeString("dumm", "0")
        w.WriteAttributeString("Party_CASTVAT", "1")
        w.WriteAttributeString("KepyoCatData_ISAGRYP", "0")
        w.WriteAttributeString("Party_Zip", "")  '66100 
        w.WriteAttributeString("Party_SNAME", Party_SNAME)
        w.WriteAttributeString("Party_DOY", "")
        w.WriteAttributeString("AM_DcTp_cd", "#ΠΛ04")
        w.WriteAttributeString("fldinvoice", "0") 'NEO neo
        w.WriteAttributeString("Party_ISK_D_A_CD", "0")
        w.WriteAttributeString("APA_VIES_v_Dscr", "EL")
        w.WriteAttributeString("KepyoCatData_SUMKEPYOYP", Replace(Str(KAU_AJIA1), ",", ".")) 'KAU_AJIA1)
        w.WriteAttributeString("F_Sites_dscr", "ΚΕΝΤΡΙΚΟ")
        w.WriteAttributeString("Base_INVOICE", Base_INVOICE)
        w.WriteAttributeString("Party_City", "δραμα")
        w.WriteAttributeString("Party_PHONE1", "")
        w.WriteAttributeString("ExpenditureKind", "0")
        w.WriteAttributeString("Party_AFM", Trim(Party_AFM))
        w.WriteAttributeString("System_sys", System_sys) 'SB =POLISEIS FR PISTVTIKA YPIRESIES FP= PLIROMES BP=AGORES
        w.WriteAttributeString("AMO_Srl_DSCR", "ΠΛΗΡΩΜΩΝ - ΠΡΟΜΗΘΕΥΤΩΝ")
        w.WriteAttributeString("System_Dscr_1", AM_DcTp_Dscr)
        w.WriteAttributeString("Party_ADDRESS", Party_ADDRESS)
        w.WriteAttributeString("Party_JOB", "εμπορια")
        w.WriteAttributeString("Base_dt", Base_dt)
        w.WriteEndElement() ' new />
        w.WriteEndElement() ' /data
        '========================================================


        '========================================================
        w.WriteStartElement("detail")  'big
        If kau23 <> 0 Then
            row_detail(LOG23, kau23, fpa23, w)
        End If
        If kau13 <> 0 Then
            row_detail(LOG13, kau13, fpa13, w)
        End If

        If kau16 <> 0 Then
            row_detail(LOG16, kau16, fpa16, w)
        End If
        '  Exit Sub
        If kau9 <> 0 Then
            row_detail(LOG9, kau9, fpa9, w)
        End If

        If kau0 <> 0 Then
            row_detail(LOG0, kau0, 0, w)
        End If




        w.WriteEndElement()  'detail  big
        '========================================================


        w.WriteEndElement() 'row hd
        'DDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDD

    End Sub


    Private Sub row_detail(ByVal mlog As String, ByVal mKau As Single, ByVal mFpa As Single, ByVal w As XmlTextWriter)


        '--------- Mv -------------------------------------
        w.WriteStartElement("row") : w.WriteAttributeString("rowid", LTrim(Str(rowId))) : w.WriteAttributeString("mode", "3") : w.WriteAttributeString("name", "Mv")



        FL_Ledg_cd = mlog ' pol13.Text
        w.WriteStartElement("data") '''''''''''''''''''''''''''''


        w.WriteStartElement("new")
        w.WriteAttributeString("FL_Ledg_Dscr", FL_Ledg_Dscr)
        w.WriteAttributeString("FL_Ledg_cd", FL_Ledg_cd)
        w.WriteAttributeString("VatVal", Replace(Str(mFpa), ",", "."))
        w.WriteAttributeString("NetVal", Replace(Str(mKau), ",", "."))
        w.WriteAttributeString("RegVal", Replace(Str(mKau), ",", "."))
        w.WriteAttributeString("MvTp", MVTP)
        w.WriteAttributeString("RegVatVal", "0.000")
        w.WriteEndElement()  'new



        w.WriteStartElement("detail")  'detail
        row_ledg("1", FL_Ledg_cd, w) '4b
        rowIdINNER = rowIdINNER + 11
        If Len(FL_Ledg_cd) = 7 Then ' 70-0057
            row_ledg("0", Mid(FL_Ledg_cd, 1, 2), w) '3b
            rowIdINNER = rowIdINNER + 11
        Else  '70-00-00-0057
            row_ledg("0", Mid(FL_Ledg_cd, 1, 8), w) '3b
            rowIdINNER = rowIdINNER + 11
            row_ledg("0", Mid(FL_Ledg_cd, 1, 5), w) '2b
            rowIdINNER = rowIdINNER + 11
            row_ledg("0", Mid(FL_Ledg_cd, 1, 2), w) '1b
            rowIdINNER = rowIdINNER + 11
        End If
        w.WriteEndElement()  'detail   


        w.WriteEndElement()  'data   '''''''''''''''''''''''''''''



        w.WriteEndElement()  'row mv
        '---------------------------------------------------

    End Sub


    Sub row_ledg(ByVal canmove As String, ByVal log As String, ByVal w As XmlTextWriter)
        '////////////////////////////
        w.WriteStartElement("row") : w.WriteAttributeString("rowid", LTrim(Str(rowIdINNER))) : w.WriteAttributeString("mode", "3") : w.WriteAttributeString("name", "Ledg")

        '''''''''''''''''''''''''''''
        w.WriteStartElement("data")
        w.WriteStartElement("new")

        w.WriteAttributeString("CanMv", canmove)
        w.WriteAttributeString("Anali", "0")
        w.WriteAttributeString("cdLedg", log)
        w.WriteAttributeString("dscrLedg", "πωλήσεις")
        w.WriteAttributeString("Active", "1")

        w.WriteEndElement()  'NEW
        w.WriteEndElement()  'data
        '''''''''''''''''''''''''''''

        w.WriteEndElement()  'row  Ledg
        '////////////////////////////

    End Sub



    Private Sub bres_file_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bres_file.Click

        If Len(Trim(TextBox1.Text)) = 0 Then
            CD1.ShowDialog()
            TextBox1.Text = CD1.FileName
        Else
            If Len(Dir(LTrim(TextBox1.Text), FileAttribute.Normal)) < 2 Then
                MsgBox("δεν υπάρχει το αρχείο " & TextBox1.Text)
                Exit Sub
            End If
        End If
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '        //**************************************
        '// Name: Create simple a Xml file in .Net
        '// Description:Create your XML file by using the included XML classes in Visual Studio. 
        '// By: Timo Boehme
        '//
        '//This code is copyrighted and has// limited warranties.Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=7002&lngWId=10//for details.//**************************************

        Dim XmlDoc As New XmlDocument
        'Write down the XML declaration
        Dim XmlDeclaration As XmlDeclaration = XmlDoc.CreateXmlDeclaration("1.0", "UTF-8", Nothing)
        'Create the root element
        Dim RootNode As XmlElement = XmlDoc.CreateElement("RootNode")
        XmlDoc.InsertBefore(XmlDeclaration, XmlDoc.DocumentElement)
        XmlDoc.AppendChild(RootNode)
        'Create a new <Category> element and add it to the root node
        Dim ParentNode As XmlElement = XmlDoc.CreateElement("Parent")
        'Set attribute name and value!
        ParentNode.SetAttribute("AttributName", "AttributWert")
        XmlDoc.DocumentElement.PrependChild(ParentNode)
        'Create the required nodes
        Dim FirstElement As XmlElement = XmlDoc.CreateElement("FirstElement")
        Dim SecondElement As XmlElement = XmlDoc.CreateElement("SecondElement")
        Dim ThirdElement As XmlElement = XmlDoc.CreateElement("ThirdElement")
        'retrieve the text
        Dim FirstTextElement As XmlText = XmlDoc.CreateTextNode("This is the text from the first element")
        Dim SecondTextElement As XmlText = XmlDoc.CreateTextNode("This is the text from the second element")
        Dim ThirdTextElement As XmlText = XmlDoc.CreateTextNode("This is the text from the third element")


        'append the nodes to the parentNode without the value
        ParentNode.AppendChild(FirstElement)
        ParentNode.AppendChild(SecondElement)
        ParentNode.AppendChild(ThirdElement)

        'save the value of the fields into the nodes
        FirstElement.AppendChild(FirstTextElement)
        SecondElement.AppendChild(SecondTextElement)
        ThirdElement.AppendChild(ThirdTextElement)
        'Save to the XML file
        XmlDoc.Save("demo.xml")

        '        <?xml version="1.0" encoding="UTF-8"?>
        '<RootNode>
        '  <Parent AttributName="AttributWert">
        '    <FirstElement>This is the text from the first element</FirstElement>
        '    <SecondElement>This is the text from the second element</SecondElement>
        '    <ThirdElement>This is the text from the third element</ThirdElement>
        '  </Parent>
        '</RootNode>




    End Sub

    Private Sub LOAD_XML()
        Dim xmlDoc As New XmlDocument()


        If InStr(UCase(filexml.Text), "XML") = 0 Then
            MsgBox("ΔΕΝ ΕΙΝΑΙ ΑΡΧΕΙΟ XML")
            Exit Sub
        End If

        xmlDoc.Load(filexml.Text) '"GAT.xml")

        Dim nodes As XmlNodeList = xmlDoc.DocumentElement.SelectNodes("/Table/Product")
        Dim pID As String = "", pName As String = "", pPrice As String = ""


        For Each node As XmlNode In nodes
            pol13.Text = node.Attributes("POL13").Value  ' ΔΙΑΒΑΖΩ ΤΟ ATTRIBUTE
            pol23.Text = node.Attributes("POL23").Value
            POL16.Text = node.Attributes("POL16").Value  ' ΔΙΑΒΑΖΩ ΤΟ ATTRIBUTE
            POL9.Text = node.Attributes("POL9").Value
            POL0.Text = node.Attributes("POL0").Value

            EPIS13.Text = node.Attributes("EPIS13").Value
            EPIS23.Text = node.Attributes("EPIS23").Value
            '==============================================================
            Lian13.Text = node.Attributes("lian13").Value  ' ΔΙΑΒΑΖΩ ΤΟ ATTRIBUTE
            Lian23.Text = node.Attributes("lian23").Value

            episLian13.Text = node.Attributes("epislian13").Value  ' ΔΙΑΒΑΖΩ ΤΟ ATTRIBUTE
            episLian23.Text = node.Attributes("epislian23").Value

            Dim a As String

            Try
                lian24.Text = node.Attributes("lian24").Value
                LOGFPA13.Text = node.Attributes("logfpa13").Value  ' ΔΙΑΒΑΖΩ ΤΟ ATTRIBUTE
                LOGFPA23.Text = node.Attributes("logfpa23").Value
                episLian24.Text = node.Attributes("epislian24").Value



                lianLOGFPA13.Text = node.Attributes("lianlogfpa13").Value  ' ΔΙΑΒΑΖΩ ΤΟ ATTRIBUTE
                lianLOGFPA23.Text = node.Attributes("lianlogfpa23").Value

                lianLOGFPA24.Text = node.Attributes("lianlogfpa24").Value



                LOGFPA16.Text = node.Attributes("LOGfpa16").Value  ' ΔΙΑΒΑΖΩ ΤΟ ATTRIBUTE
                LOGFPA9.Text = node.Attributes("logfpa9").Value




                pel30.Text = node.Attributes("pel30").Value
                prom50.Text = node.Attributes("prom50").Value

                ago13.Text = node.Attributes("ago13").Value
                ago23.Text = node.Attributes("ago23").Value
                ago16.Text = node.Attributes("ago16").Value
                ago9.Text = node.Attributes("ago9").Value

                agoepis13.Text = node.Attributes("agoepis13").Value
                agoepis23.Text = node.Attributes("agoepis23").Value
                agoepis16.Text = node.Attributes("agoepis16").Value
                agoepis9.Text = node.Attributes("agoepis9").Value

                agofpa13.Text = node.Attributes("agofpa13").Value
                agofpa23.Text = node.Attributes("agofpa23").Value
                agofpa16.Text = node.Attributes("agofpa16").Value
                agofpa9.Text = node.Attributes("agofpa9").Value


                ago24_6.Text = node.Attributes("ago24_6").Value
                agoepis24_6.Text = node.Attributes("agoepis24_6").Value
                agofpa24_6.Text = node.Attributes("agofpa24_6").Value



            Catch ex As Exception
                MsgBox("δεν διαβαστηκαν οι παραμετροι λογ/σμων")
            End Try




            Try

                nTimol.Text = node.Attributes("nTimol").Value
                cTimol.Text = node.Attributes("cTimol").Value

                nLian.Text = node.Attributes("nLian").Value
                cLian.Text = node.Attributes("cLian").Value


                nPistTim.Text = node.Attributes("nPistTim").Value
                cPistTim.Text = node.Attributes("cPistTim").Value

                nPistLian.Text = node.Attributes("nPistLian").Value
                cPistLian.Text = node.Attributes("cPistLian").Value


                nParox.Text = node.Attributes("nParox").Value
                cParox.Text = node.Attributes("cParox").Value




            Catch ex As Exception
                MsgBox("δεν διαβαστηκαν οι παραμετροι αναγνωρισης παραστατικων")
            End Try



            Try
                POL24.Text = node.Attributes("POL24").Value
                EPIS24.Text = node.Attributes("EPIS24").Value
                LOGFPA24.Text = node.Attributes("LOGFPA24").Value
                
                POL17.Text = node.Attributes("POL17").Value
                EPIS17.Text = node.Attributes("EPIS17").Value
                LOGFPA17.Text = node.Attributes("LOGFPA17").Value

                PAR24.Text = node.Attributes("PAR24").Value
                EPISPAR24.Text = node.Attributes("EPISPAR24").Value

            Catch ex As Exception
                MsgBox("δεν διαβαστηκαν οι παραμετροι λογ/σμων")
            End Try
            

           







        Next






        ' ΠΑΡΑΔΕΙΓΜΑ INNER TEXT

        'Dim xmlDoc As New XmlDocument()
        'xmlDoc.Load("GAT.xml")
        'Dim nodes As XmlNodeList = xmlDoc.DocumentElement.SelectNodes("/Table/Product")
        'Dim pID As String = "", pName As String = "", pPrice As String = ""
        'For Each node As XmlNode In nodes
        '    PID = node.SelectSingleNode("POL13").InnerText' DIABAZO TO INNER TEXT
        '    pPrice = node.SelectSingleNode("POL23").InnerText
        '    MessageBox.Show(pID & " " & pName & " " & pPrice)
        'Next
        '-----------------------------------------------------------------------------
        '        <Table>             
        '<Product>
        '       <POL13>70-0036</POL13>
        '       <POL23>70-0057</POL23>
        '</Product>

        '</Table>




        ' ΠΑΡΑΔΕΙΓΜΑ ΜΕ ATTRIBUTES
        '        <?xml version="1.0" encoding="utf-8" standalone="yes"?>
        '<Table>
        '  <Product POL13="70-0036" POL23="70-0057" />
        '</Table

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '====================================================================================
        Dim ff As String = filexml.Text ' "GAT.XML" ' "\\Logisthrio\333\pr.export" '

        Dim writer As New XmlTextWriter(ff, System.Text.Encoding.UTF8)
        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2

        writer.WriteStartElement("Table")

        writer.WriteStartElement("Product")

        writer.WriteAttributeString("POL13", pol13.Text)
        writer.WriteAttributeString("POL23", pol23.Text)
        writer.WriteAttributeString("POL16", POL16.Text)
        writer.WriteAttributeString("POL9", POL9.Text)
        writer.WriteAttributeString("POL0", POL0.Text)
        writer.WriteAttributeString("EPIS13", EPIS13.Text)
        writer.WriteAttributeString("EPIS23", EPIS23.Text)
        '-----------------------------------
        writer.WriteAttributeString("lian13", Lian13.Text)
        writer.WriteAttributeString("lian23", Lian23.Text)
        writer.WriteAttributeString("lian24", Lian23.Text)


        writer.WriteAttributeString("epislian13", episLian13.Text)
        writer.WriteAttributeString("epislian23", episLian23.Text)
        writer.WriteAttributeString("epislian24", episLian23.Text)


        writer.WriteAttributeString("logfpa13", LOGFPA13.Text)
        writer.WriteAttributeString("logfpa23", LOGFPA23.Text)

        writer.WriteAttributeString("lianlogfpa13", lianLOGFPA13.Text)
        writer.WriteAttributeString("lianlogfpa23", lianLOGFPA23.Text)
        writer.WriteAttributeString("lianlogfpa24", lianLOGFPA13.Text)

        writer.WriteAttributeString("LOGfpa16", LOGFPA16.Text)
        writer.WriteAttributeString("logfpa9", LOGFPA9.Text)

        writer.WriteAttributeString("pel30", pel30.Text)
        writer.WriteAttributeString("prom50", prom50.Text)


        '----------------------------------------------------------------

        writer.WriteAttributeString("ago13", ago13.Text)
        writer.WriteAttributeString("ago23", ago23.Text)
        writer.WriteAttributeString("ago16", ago16.Text)
        writer.WriteAttributeString("ago9", ago9.Text)

        writer.WriteAttributeString("agoepis13", agoepis13.Text)
        writer.WriteAttributeString("agoepis23", agoepis23.Text)
        writer.WriteAttributeString("agoepis16", agoepis16.Text)
        writer.WriteAttributeString("agoepis9", agoepis9.Text)

        writer.WriteAttributeString("agofpa13", agofpa13.Text)
        writer.WriteAttributeString("agofpa23", agofpa23.Text)
        writer.WriteAttributeString("agofpa16", agofpa16.Text)
        writer.WriteAttributeString("agofpa9", agofpa9.Text)

        writer.WriteAttributeString("nTimol", nTimol.Text)
        writer.WriteAttributeString("cTimol", cTimol.Text)

        writer.WriteAttributeString("nParox", nParox.Text)
        writer.WriteAttributeString("cParox", cParox.Text)


        writer.WriteAttributeString("nLian", nLian.Text)
        writer.WriteAttributeString("cLian", cLian.Text)

        writer.WriteAttributeString("nPistTim", nPistTim.Text)
        writer.WriteAttributeString("cPistTim", cPistTim.Text)

        writer.WriteAttributeString("nPistLian", nPistLian.Text)
        writer.WriteAttributeString("cPistLian", cPistLian.Text)

        'POL24.Text = node.Attributes("POL24").Value
        'EPIS24.Text = node.Attributes("EPIS24").Value
        'LOGFPA24.Text = node.Attributes("LOGFPA24").Value

        'POL17.Text = node.Attributes("POL17").Value
        'EPIS17.Text = node.Attributes("EPIS17").Value
        'LOGFPA17.Text = node.Attributes("LOGFPA17").Value

        writer.WriteAttributeString("POL24", POL24.Text)
        writer.WriteAttributeString("EPIS24", EPIS24.Text)
        writer.WriteAttributeString("LOGFPA24", LOGFPA24.Text)

        writer.WriteAttributeString("POL17", POL17.Text)
        writer.WriteAttributeString("EPIS17", EPIS17.Text)
        writer.WriteAttributeString("LOGFPA17", LOGFPA17.Text)

        writer.WriteAttributeString("PAR24", PAR24.Text)
        writer.WriteAttributeString("EPISPAR24", EPISPAR24.Text)


        writer.WriteAttributeString("ago24_6", ago24_6.Text)
        writer.WriteAttributeString("agoepis24_6", agoepis24_6.Text)
        writer.WriteAttributeString("agofpa24_6", agofpa24_6.Text)

        ' ago24_6.Text = node.Attributes("ago24_6").Value
        ' agoepis24_6.Text = node.Attributes("agoepis24_6").Value
        ' agofpa24_6.Text = node.Attributes("agofpa24_6").Value




        writer.WriteEndElement() ' PRODUCT
        writer.WriteEndElement() ' TABLE



        writer.WriteEndDocument()
        writer.Close()


    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        LOAD_XML()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        'query.Visible = True
        MsgBox("SELECT  AJ1,AJ2,AJ3,AJ4,AJ5,AJI,FPA1,FPA2,FPA3,FPA4,ATIM,CONVERT(CHAR(10),HME,3) AS HMEP,PEL.EPO,PEL.AFM,KPE,PEL.DIE,PEL.XRVMA,PEL.EPA,PEL.POL FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD WHERE LEFT(ATIM,1) IN ('L','l','T','t','Y')   and HME>=@x1 AND HME<=@X2 order by HME")
        MsgBox("1=καθ13% 2=καθ23  3=6.5%	4=συνολο	5=φπα13	6=φπα23	7=φπα6.6	8=ATIM	9=HME	10=EPO	11=AFM	12=KPE	13=DIE	14=τκ 15=EPA	17=POL" & Chr(13) & "")
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        'αμβροσιαδης

        'ΣΤΟ ΦΥΛΛΟ 2 ΕΧΩ ΤΟΥΣ ΠΕΛΑΤΕΣ ΜΕ ΑΦΜ ΚΑΙ ΣΤΟ ΦΥΛΛΟ1 ΤΑ ΤΙΜΟΛΟΓΙΑ ΜΕ ΤΑ ΠΟΣΑ
        'μεταφέρει το ΑΦΜ ΣΤΟ ΦΥΛΛΟ1(στηλη 14)  ΑΠΟ ΤΟ ΦΥΛΛΟ2

        ' pel(ROW, 2)  πινακας που φορτώνει ολους τους πελατες απο το φυλλο 2
        ' 


        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xl As Excel.Worksheet
        Dim xlPEL As Excel.Worksheet
        Dim xlok As Excel.Worksheet

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open(TextBox1.Text)
        xlWorkBook.Worksheets.Add()


        xl = xlWorkBook.Worksheets(2) ' .Add

        xlPEL = xlWorkBook.Worksheets(3)

        xlok = xlWorkBook.Worksheets(1)

        'metafora me σωστη γραμμογραφηση στο 3
        '=========================================
        '===============================================================================real onomatepvmymo 54100
        ROW = 1
        Do While True
            ROW = ROW + 1
            If xl.Cells(ROW, 1).value = Nothing Then
                Exit Do
            End If
            xlok.Cells(ROW, 1) = xl.Cells(ROW, 5) '13% kauarh
            xlok.Cells(ROW, 2) = xl.Cells(ROW, 4) ' 23%
            xlok.Cells(ROW, 5) = xl.Cells(ROW, 10) '0%

            xlok.Cells(ROW, 6) = xl.Cells(ROW, 14) 'συνολικη αξια

            'fpa
            xlok.Cells(ROW, 7) = xl.Cells(ROW, 15).value  'fpa 13
            xlok.Cells(ROW, 8) = xl.Cells(ROW, 8).value  '23%




            'xlok.Cells(ROW, 7) = xl.Cells(ROW, 12).value - xl.Cells(ROW, 8).value  '13%
            '11 apa   12 hme   13 epo  14 afm



            xlok.Cells(ROW, 11) = xl.Cells(ROW, 2).ToString   'apa
            xlok.Cells(ROW, 12) = xl.Cells(ROW, 1).value  'hmeromhnia

            xlok.Cells(ROW, 13) = xl.Cells(ROW, 3).ToString   'epvnymia

            xlok.Cells(ROW, 14) = xl.Cells(ROW, 3).ToString   'epvnymia





            Me.Text = ROW



        Loop
        MsgBox("ok")

        xlWorkBook.Save()
        xlApp.Quit()

        Exit Sub


        '==========================================

        Dim pel(2000, 2) As String

        ROW = 1

        Dim hand As Integer = 0

        '===============================================================================real onomatepvmymo 54100
        Do While True
            ROW = ROW + 1
            If xlPEL.Cells(ROW, 1).value = Nothing Then
                Exit Do
            End If
            pel(ROW, 1) = xlPEL.Cells(ROW, 1).value.ToString

            If IsDBNull(xlPEL.Cells(ROW, 2).value) Then
                pel(ROW, 2) = ""
            Else
                If xlPEL.Cells(ROW, 2).value = Nothing Then
                    pel(ROW, 2) = ""
                Else
                    pel(ROW, 2) = xlPEL.Cells(ROW, 2).value.ToString
                End If
            End If
            Me.Text = Str(ROW)



        Loop


        'βαζω τα ΑΦΜ  ΣΤΟ ΦΥΛΛΟ1

        ROW = 1
        Dim K As Integer, N As Integer
        Dim C As String

        ''===============================================================================real onomatepvmymo 54100
        'Do While True
        '    ROW = ROW + 1

        '    If xl.Cells(ROW, 1).value = Nothing Then
        '        Exit Do
        '    End If

        '    'N = InStr(xl.Cells(ROW, 3).VALUE.ToString, "-")

        '    ''ΑΝ ΕΧΕΙ 2Η ΠΑΥΛΑ  ΠΑΡΕ ΤΗΝ ΤΕΛΕΥΤΑΙΑ
        '    'If N < InStrRev(xl.Cells(ROW, 3).VALUE.ToString, "-") Then
        '    '    N = InStrRev(xl.Cells(ROW, 3).VALUE.ToString, "-")
        '    'End If



        '    'If N <= 1 Then
        '    '    C = ""
        '    'Else
        '    '    C = Mid(xl.Cells(ROW, 3).VALUE.ToString, 1, N - 1)
        '    'End If
        '    C = xl.Cells(ROW, 13).VALUE.ToString
        '    xl.Cells(ROW, 14) = SCAN_PEL(C, pel)
        '    Me.Text = Str(ROW)

        'Loop
        'MsgBox("ok ")


        '===============================================================================real onomatepvmymo 54100
        Do While True
            ROW = ROW + 1

            If xl.Cells(ROW, 13).value = Nothing Then
                Exit Do
            End If

            'N = InStr(xl.Cells(ROW, 3).VALUE.ToString, "-")

            ''ΑΝ ΕΧΕΙ 2Η ΠΑΥΛΑ  ΠΑΡΕ ΤΗΝ ΤΕΛΕΥΤΑΙΑ
            'If N < InStrRev(xl.Cells(ROW, 3).VALUE.ToString, "-") Then
            '    N = InStrRev(xl.Cells(ROW, 3).VALUE.ToString, "-")
            'End If



            'If N <= 1 Then
            '    C = ""
            'Else
            '    C = Mid(xl.Cells(ROW, 3).VALUE.ToString, 1, N - 1)
            'End If
            C = xl.Cells(ROW, 13).VALUE.ToString
            xl.Cells(ROW, 14) = SCAN_PEL(C, pel)
            Me.Text = Str(ROW) + xl.Cells(ROW, 14).ToString

        Loop


















        '    xlWorkBook.Save()
        '        xlApp.Quit()

        xlWorkBook.Save()
        xlApp.Quit()




    End Sub


    Function SCAN_PEL(ByVal X As String, ByRef PEL(,) As String) As String
        Dim K As Integer
        SCAN_PEL = ""
        Dim L As Integer
        For K = 1 To 2000
            L = Len(PEL(K, 1)) 'ΜΗΚΟΣ ΚΩΔΙΚΟΥ
            If L > 1 Then
                If Mid(X, 1, L) = PEL(K, 1) Then
                    SCAN_PEL = PEL(K, 2)
                    Exit For
                End If
            End If

        Next
        'For K = 1 To 2000
        '    If X = PEL(K, 1) Then
        '        SCAN_PEL = PEL(K, 2)
        '        Exit For
        '    End If

        'Next

    End Function




    Function getControlFromName(ByRef containerObj As Object, _
                         ByVal name As String) As Control
        Try
            Dim tempCtrl As Control
            For Each tempCtrl In containerObj.Controls
                If tempCtrl.Name.ToUpper.Trim = name.ToUpper.Trim Then
                    Return tempCtrl
                End If
            Next tempCtrl
        Catch ex As Exception
        End Try

    End Function







    Private Sub set_TextBOX(ByVal onoma As String, ByVal timh As String)
        Dim tempText As TextBox = _
           CType(getControlFromName(Me, onoma), TextBox)
        tempText.Text = timh
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Use To use it, enclose it in a CType function to give you a useful reference to the control. Like this..
        'Hide   Copy Code
        Dim tempText As TextBox = _
           CType(getControlFromName(Me, "pol13"), TextBox)
        Me.Text = tempText.Text


        'Dim tempCtrl As Button = _
        '           CType(getControlFromName(Me, "Button6"), Button)
        'Me.Text = tempCtrl.Text


    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        CD1.ShowDialog()
        filexml.Text = CD1.FileName
        LOAD_XML()


    End Sub


    Private Sub xmlG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles xmlG.Click
        Dim a As String
        Dim K As Short
        Dim C As String

        ' CO TO DIAXORISTIKO DEKADIKON ARITMON
        Dim CO As String = String.Format(1.1).Substring(1, 1)


        MsgBox("ΠΡΟΣΟΧΗ ΔΙΑΒΑΖΕΙ ΑΠΟ ΤΗΝ " + ApoSeira.Text + "η ΣΕΙΡΑ ΜΕ ΓΡΑΜΟΓΡΑΦΗΣΗ:" + Chr(13) + "AJ1 AJ2	AJ3	AJ4	AJ5	AJI	FPA1	FPA2	FPA3	FPA4	ATIM	HME	EPO	AFM")


        ' Write the string as utf-8.
        ' This also writes the 3-byte utf-8 preamble at the beginning of the file.
        Dim appendMode As Boolean = False ' This overwrites the entire file.
        ' Dim sw As New StreamWriter("C:\MERCVB\out_utf9.export", appendMode, System.Text.Encoding.UTF8)
        'sw.Write(TextBox1.Text)
        'sw.Close()


        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xl As Excel.Worksheet

        xlApp = New Excel.ApplicationClass
        Try
            xlWorkBook = xlApp.Workbooks.Open(TextBox1.Text)
            xl = xlWorkBook.Worksheets(1) ' .Add
        Catch
            MsgBox("Δεν ανοιγει το αρχείο excel")
            Exit Sub

        End Try


        '====================================================================================
        Dim ff As String = "c:\mercvb\m" + VB6.Format(Now, "YYYYddmmHHMM") + ".export" ' "\\Logisthrio\333\pr.export" '

        Dim writer As New XmlTextWriter(ff, System.Text.Encoding.UTF8)
        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2
        ' Create a Continent element and set its value to
        ' that of the New Continent dialog box
        'writer.WriteAttributeString("Table", , "sadasasd sdsd")

        '<Data Name="SX" Style="BRowse"><row name="Hd" mode="3" rowId="7">
        writer.WriteStartElement("Data")
        writer.WriteAttributeString("Name", "GL")
        writer.WriteAttributeString("Style", "Browse")
        '====================================================================================


        Dim enter_Renamed As String
        enter_Renamed = Chr(13)

        'FileOpen(1, "C:\MERCVB\A778.XML", OpenMode.Output)
        ROW = Val(ApoSeira.Text) - 1

        Dim hand As Integer = 0
        Dim suma As Single = 0
        Dim SXETIKO As String
        '===============================================================================real onomatepvmymo 54100
        Do While True
            ROW = ROW + 1

            Me.Text = ROW
            'system.doevents

            If IsDBNull(xl.Cells(ROW, 12).value) Then
                Exit Do
            End If

            If Len(xl.Cells(ROW, 11).ToString) < 2 Then
                Exit Do
            End If
            If xl.Cells(ROW, 11).value = Nothing Then
                Exit Do
            End If

            '1	 2	    3	4	5	6	7	    8	    9	    10	    11	    12	13	14	15	16	17	    18	19
            'AJ1 AJ2	AJ3	AJ4	AJ5	AJI	FPA1	FPA2	FPA3	FPA4	ATIM	HME	EPO	AFM	KPE	DIE	XRVMA	EPA	POL
            Party_IDParty = xl.Cells(ROW, 14).value  ' As String  '12344   ΚΩΔ ΣΥΝΑΛΛΑΣΟΜΕΝΟΥ
            AM_DcTp_Dscr = "Τιμολόγιο"
            Party_AFM = Trim(xl.Cells(ROW, 14).value)  'Dim Party_AFM As String ' =""999349996
            If Len(Trim(Party_AFM)) <= 4 Then
                Party_AFM = "000000000"
            End If

            Party_ADDRESS = xl.Cells(ROW, 16).value 'ToString  ' "ΠΟΛΥΣΤΗΛΟ ΚΑΒΑΛΑΣ"
            AM_DcTp_cd = "#ΤΥΠ-0"
            AMO_Srl_DSCR = "Πωλήσεις" '"ΠΩΛΗΣΕΙΣ"
            Base_dt = VB6.Format(xl.Cells(ROW, 12), "YYYY-mm-dd")

            SXETIKO = Mid(xl.Cells(ROW, 22).ToString, 9, 7)   'Σχ.Παρ. Τ000123
            Base_INVOICE = xl.Cells(ROW, 11).value
            ' =""#ΤΥΠ-0/Π000/1/Τιμολόγιο Παροχής Υπηρεσιών"
            Party_SNAME = xl.Cells(ROW, 13).value  '=""Θ. ΓΡΑΜΜΑΤΗΣ Κ.ΣΙΑ Ε.Ε""
            f_logPel = pel30.Text ' "30-00-00-0000"

            ListBox1.Items.Add(Str(ROW) + ". " + Party_SNAME)


            KAU_AJIA = nVal(xl.Cells(ROW, 1).value) + nVal(xl.Cells(ROW, 2).value) + nVal(xl.Cells(ROW, 3).value) + nVal(xl.Cells(ROW, 4).value) + nVal(xl.Cells(ROW, 5).value)
            FPA = nVal(xl.Cells(ROW, 7).value) + nVal(xl.Cells(ROW, 8).value) + nVal(xl.Cells(ROW, 9).value) + nVal(xl.Cells(ROW, 10).value)
            suma = suma + KAU_AJIA
            kau13 = nVal(xl.Cells(ROW, 1).value)
            kau23 = nVal(xl.Cells(ROW, 2).value)
            kau16 = nVal(xl.Cells(ROW, 3).value)
            kau9 = nVal(xl.Cells(ROW, 4).value)
            kau0 = nVal(xl.Cells(ROW, 5).value)

            fpa13 = nVal(xl.Cells(ROW, 7).value)
            fpa23 = nVal(xl.Cells(ROW, 8).value)
            fpa16 = nVal(xl.Cells(ROW, 9).value)

            FL_Ledg_Dscr = "ΠΩΛΗΣΕΙΣ ΧΟΝΔΡΙΚΗΣ ΕΣ. ΦΠΑ23%"
            FL_Ledg_cd = pol23.Text ' "70-00-00-0057"

            MVTP = "1"
            System_sys = "SB" '                     'SB =POLISEIS FR   η ειναι ακυρωτικο 
            If InStr("Lρ", Mid(Base_INVOICE, 1, 1)) > 0 Or (Mid(Base_INVOICE, 1, 1) = "κ" And InStr("Lρ", Mid(SXETIKO, 1, 1)) > 0) Then
                IsHand = "1" 'LTrim(Str(hand))
                cdRetailIdentity = ""
                LOG13 = Lian13.Text : LOG23 = Lian23.Text
                logarFpa23 = lianLOGFPA23.Text : logarFpa13 = lianLOGFPA13.Text
                LOG0 = LIAN0.Text
                f_logPel = "38-00-00-0000"
                Party_AFM = "000000000"
                f_aitiologia = "ΛΙΑΝΙΚΕΣ ΠΩΛΗΣΕΙΣ"
                Party_IDParty = ""
                tit_paras = "ΑΠΛ"

                'αν ειναι ακυρωτικό λιανικής
                If Mid(Base_INVOICE, 1, 1) = "κ" And InStr("Lρ", Mid(SXETIKO, 1, 1)) > 0 Then
                    f_aitiologia = "ΑΚΥΡΩΤΙΚΟ ΛΙΑΝΙΚΩΝ ΠΩΛΗΣΕΩΝ"
                    Party_IDParty = ""
                    tit_paras = "ΑΚΥΡ"
                End If
                Metrhtaxond = False


                'ElseIf
            Else
                Metrhtaxond = False
                IsHand = ""
                Party_IDParty = Mid(Party_AFM, 5, 5)

                If InStr("Tt", Mid(Base_INVOICE, 1, 1)) > 0 Then

                    LOG13 = pol13.Text : LOG23 = pol23.Text
                    LOG16 = POL16.Text : LOG9 = POL9.Text
                    LOG0 = POL0.Text
                    logarFpa23 = LOGFPA23.Text : logarFpa13 = LOGFPA13.Text
                    tit_paras = "ΤΠ"
                    f_aitiologia = "ΧΟΝΔΡΙΚΕΣ ΠΩΛΗΣΕΙΣ"
                    If xl.Cells(ROW, 1).ToString = "ΜΕ" Then
                        Metrhtaxond = True
                    End If
                End If




                'αν ειναι ακυρωτικό ΧΟΝΔΡΙΚΗς
                If Mid(Base_INVOICE, 1, 1) = "κ" And InStr("Tt", Mid(SXETIKO, 1, 1)) > 0 Then

                    LOG13 = pol13.Text : LOG23 = pol23.Text
                    LOG16 = POL16.Text : LOG9 = POL9.Text
                    LOG0 = POL0.Text
                    logarFpa23 = LOGFPA23.Text : logarFpa13 = LOGFPA13.Text
                    f_aitiologia = "ΑΚΥΡΩΤΙΚΟ ΧΟΝΔΡΙΚΩΝ ΠΩΛΗΣΕΩΝ"
                    Party_IDParty = ""
                    tit_paras = "ΑΚΥΡ"
                End If





                If Mid(Base_INVOICE, 1, 1) = "P" Or (Mid(Base_INVOICE, 1, 1) = "κ" And InStr("P", Mid(SXETIKO, 1, 1)) > 0) Then
                    LOG13 = EPIS13.Text : LOG23 = EPIS23.Text
                    MVTP = 6
                    f_aitiologia = "ΕΠΙΣΤΡΟΦΕΣ ΠΩΛΗΣΕΩΝ"
                    tit_paras = "ΠΤ"
                    kau13 = -kau13
                    kau23 = -kau23
                    fpa13 = -fpa13
                    fpa23 = -fpa23
                    KAU_AJIA = -KAU_AJIA
                    FPA = -FPA
                    'αν ειναι ακυρωτικό pistotikoy ΧΟΝΔΡΙΚΗς
                    If Mid(Base_INVOICE, 1, 1) = "κ" And InStr("P", Mid(SXETIKO, 1, 1)) > 0 Then
                        f_aitiologia = "ΑΚΥΡΩΤΙΚΟ ΠΙΣΤΩΤΙΚΩΝ ΠΩΛΗΣΕΩΝ"
                        Party_IDParty = ""
                        tit_paras = "ΑΚΥΡ"
                    End If





                End If




                If Mid(Base_INVOICE, 1, 1) = "p" Or Mid(Base_INVOICE, 1, 1) = "κ" And InStr("p", Mid(SXETIKO, 1, 1)) > 0 Then
                    Party_IDParty = ""
                    LOG13 = episLian13.Text : LOG23 = episLian23.Text
                    logarFpa23 = lianLOGFPA23.Text : logarFpa13 = lianLOGFPA13.Text
                    f_logPel = "38-00-00-0000"
                    Party_AFM = "000000000"
                    MVTP = 6
                    IsHand = "1"
                    kau13 = -kau13
                    kau23 = -kau23
                    fpa13 = -fpa13
                    fpa23 = -fpa23
                    KAU_AJIA = -KAU_AJIA
                    FPA = -FPA
                    f_aitiologia = "ΕΠΙΣΤΡΟΦΕΣ ΛΙΑΝΙΚΩΝ ΠΩΛΗΣΕΩΝ"
                    tit_paras = "ΔΕΠ"

                    'αν ειναι ακυρωτικό pistotikoy ΧΟΝΔΡΙΚΗς
                    If Mid(Base_INVOICE, 1, 1) = "κ" And InStr("p", Mid(SXETIKO, 1, 1)) > 0 Then
                        f_aitiologia = "ΑΚΥΡΩΤΙΚΟ επιστροφων λιανικων"
                        Party_IDParty = ""
                        tit_paras = "ΑΚΥΡ"
                    End If

                End If
                If Mid(Base_INVOICE, 1, 2) = "ΠΤ" Then
                    kau13 = -kau13
                    kau23 = -kau23
                    fpa13 = -fpa13
                    fpa23 = -fpa23
                    KAU_AJIA = -KAU_AJIA
                    FPA = -FPA
                    LOG13 = EPIS13.Text : LOG23 = EPIS23.Text
                    MVTP = 6
                    f_aitiologia = "ΕΠΙΣΤΡΟΦΕΣ ΠΩΛΗΣΕΩΝ"
                    Party_IDParty = Mid(Party_AFM, 5, 5)
                    tit_paras = "ΠΤ"
                End If
                cdRetailIdentity = ""

            End If
            KAU_AJIA1 = KAU_AJIA
            FPA1 = FPA
            writeG_row(writer)
            rowId = rowId + 11
        Loop





        writer.WriteEndDocument()
        writer.Close()





        MsgBox("Δημιουργήθηκε στο " + ff)
        xlApp.Quit()
        Me.Text = "ΣΥΝΟΛΟ ΚΑΘ.ΑΞΙΑΣ " + VB6.Format(suma, "#####,###,###.00")
    End Sub





    Sub writeG_row(ByVal w As XmlTextWriter)
        'DDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDD

        Dim mBase_INVOICE As String = LTrim(Str(Val(Mid(Base_INVOICE, 2, 6))))



        'big row
        w.WriteStartElement("row") : w.WriteAttributeString("rowid", LTrim(Str(rowId))) : w.WriteAttributeString("mode", "3") : w.WriteAttributeString("name", "Hd")


        '========================================================
        w.WriteStartElement("data")
        w.WriteStartElement("new")

        If Mid(Base_INVOICE, 1, 1) = "L" Or Mid(Base_INVOICE, 1, 1) = "p" Then
            'ok για ΆΠΟΔΕΙΞΕΙΣ ΛΙΑΝΙΚΗΣ ΚΑΙ ΔΕΛΤΙΑ ΕΠΙΣΤΡΟΦΗΣ----------------------
            'w.WriteAttributeString("IsHand", IsHand)
            'w.WriteAttributeString("F_Sites_cd", "001")
            'w.WriteAttributeString("ConstrCost", "0")
            'w.WriteAttributeString("Party_ISK_D_A_Dscr", "ΚΑΝΟΝΙΚΟΣ")
            'w.WriteAttributeString("AM_DcTp_Dscr", AM_DcTp_Dscr)
            'w.WriteAttributeString("GlbCff", "1")
            'w.WriteAttributeString("Party_Sts", "1") ' neo
            'w.WriteAttributeString("Party_IDParty", Mid(Party_AFM, 5, 5)) ' κωδικος πελατη 13 
            ''w.WriteAttributeString("AMO_Srl_cd", "")   'Base_INVOICE)
            'w.WriteAttributeString("KepyoCatData_SUMKEPYOVAT", FPA1)
            'w.WriteAttributeString("Party_PHONE2", "")
            'w.WriteAttributeString("Party_CASTVAT_Dscr", "ΚΑΝΟΝΙΚΟ")
            'w.WriteAttributeString("cdRetailIdentity", cdRetailIdentity)
            'w.WriteAttributeString("Ledger_Cust", pel30.Text)  'neo   
            'w.WriteAttributeString("dumm", "0")
            'w.WriteAttributeString("Party_CASTVAT", "1")
            'w.WriteAttributeString("KepyoCatData_ISAGRYP", "0")
            'w.WriteAttributeString("Party_Zip", "")  '66100 
            'w.WriteAttributeString("Party_SNAME", "ΛΙΑΝΙΚΕΣ ΠΩΛΗΣΕΙΣ") ' Party_SNAME
            'w.WriteAttributeString("Party_DOY", "")
            'If Mid(Base_INVOICE, 1, 1) = "p" Then
            '    w.WriteAttributeString("AM_DcTp_cd", "Δ.Ε")
            'Else
            '    w.WriteAttributeString("AM_DcTp_cd", "ΑΛΠ")
            'End If
            'w.WriteAttributeString("fldinvoice", "0") 'NEO neo
            'w.WriteAttributeString("Party_ISK_D_A_CD", "0")
            'w.WriteAttributeString("APA_VIES_v_Dscr", "EL")
            'w.WriteAttributeString("KepyoCatData_SUMKEPYOYP", KAU_AJIA1)
            'w.WriteAttributeString("F_Sites_dscr", "ΚΕΝΤΡΙΚΟ")
            'w.WriteAttributeString("Base_INVOICE", Base_INVOICE)
            'w.WriteAttributeString("Cnt", Base_INVOICE)
            'w.WriteAttributeString("JrnCnt", "0")
            'w.WriteAttributeString("Party_City", "δραμα")
            'w.WriteAttributeString("Party_PHONE1", "")
            'w.WriteAttributeString("KEPYO_BClass", "0")
            'w.WriteAttributeString("Ap_Party_cd", "18-015") ' sthn tyxh
            'w.WriteAttributeString("ExpenditureKind", "0")
            'w.WriteAttributeString("Party_AFM", Party_AFM)
            'w.WriteAttributeString("System_sys", System_sys) 'SB =POLISEIS FR PISTVTIKA YPIRESIES FP= PLIROMES BP=AGORES
            'w.WriteAttributeString("AMO_Srl_DSCR", f_aitiologia)
            'w.WriteAttributeString("System_Dscr_1", AM_DcTp_Dscr)
            'w.WriteAttributeString("Party_ADDRESS", Party_ADDRESS)
            'w.WriteAttributeString("Party_JOB", "εμπορια")
            'w.WriteAttributeString("Base_dt", Base_dt)
            'w.WriteAttributeString("dt", Base_dt)

            w.WriteAttributeString("System_sys", "SB")
            w.WriteAttributeString("XU_Usr_cd", "inner")
            w.WriteAttributeString("KEPYO_BClass", "0")
            w.WriteAttributeString("AMO_Srl_cd", "") ' "Π000")
            w.WriteAttributeString("ConstrCost", "0")
            ' w.WriteAttributeString("AM_DcTp_Dscr", "Απόδειξη Λιανικής Πώλησης")
            'w.WriteAttributeString("AM_DcTp_cd", "ΑΛΠ")

            If Mid(Base_INVOICE, 1, 1) = "p" Then
                w.WriteAttributeString("AM_DcTp_cd", "Δ.Ε")
                w.WriteAttributeString("AM_DcTp_Dscr", "Απόδειξη Eπιστροφης Λιαν.Πώλησης")
            Else
                w.WriteAttributeString("AM_DcTp_cd", "ΑΛΠ")
                w.WriteAttributeString("AM_DcTp_Dscr", "Απόδειξη Λιανικής Πώλησης")
            End If





            w.WriteAttributeString("Party_ISK_D_A_Dscr", "")
            w.WriteAttributeString("Reserved", "0")
            w.WriteAttributeString("AMO_Srl_DSCR", "ΠΩΛΗΣΕΙΣ (ΜΗΧΑΝΟΓΡΑΦΗΜΕΝΗ)")
            w.WriteAttributeString("KepyoCatData_ISAGRYP", "0")
            w.WriteAttributeString("System_Dscr_1", "Πωλήσεις")
            w.WriteAttributeString("F_Coin_Dscr", "ΕURO")
            w.WriteAttributeString("Party_CASTVAT_Dscr", "ΚΑΝΟΝΙΚΟ")
            w.WriteAttributeString("Party_SNAME", "ΠΕΛΑΤΗΣ ΛΙΑΝΙΚΗΣ*")
            w.WriteAttributeString("Party_ADDRESS", "")
            w.WriteAttributeString("FL_Dgrs_Dscr", "ΓΕΝΙΚΗ ΛΟΓΙΣΤΙΚΗ / ΕΣΟΔΑ - ΕΞΟΔΑ")
            w.WriteAttributeString("cdRetailIdentity", "ΑΝΕΥ")
            w.WriteAttributeString("JrnCnt", "0")
            w.WriteAttributeString("KepyoCatData_SUMKEPYOVAT", FPA1)
            w.WriteAttributeString("Party_CASTVAT", "1")
            w.WriteAttributeString("CDt", Base_dt)
            w.WriteAttributeString("F_Coin_ShCut", "€")
            w.WriteAttributeString("F_Sites_cd", "001")
            w.WriteAttributeString("authDt", Base_dt)
            w.WriteAttributeString("GlbCff", "1")
            w.WriteAttributeString("Dgrs", "1")
            w.WriteAttributeString("Party_Sts", "1")
            w.WriteAttributeString("AP_Party_Dscr", "ΛΙΑΝΙΚΕΣ ΠΩΛΗΣΕΙΣ")
            w.WriteAttributeString("APA_VIES_v_Dscr", "EL")
            w.WriteAttributeString("fldinvoice", "0")
            w.WriteAttributeString("Party_AFM", "000000000")
            w.WriteAttributeString("Cnt", mBase_INVOICE)
            w.WriteAttributeString("Party_IDParty", "1")
            w.WriteAttributeString("KepyoCatData_SUMKEPYOYP", KAU_AJIA1)
            w.WriteAttributeString("IsHand", "1")
            w.WriteAttributeString("AP_Party_cd", "01")
            w.WriteAttributeString("dt", Base_dt)
            w.WriteAttributeString("XU_Usr_dscr", "inner")
            w.WriteAttributeString("F_Sites_dscr", "ΚΕΝΤΡΙΚΟ")
            w.WriteAttributeString("ExpenditureKind", "0")

            'w.WriteAttributeString("System_sys", "SB")
            'w.WriteAttributeString("XU_Usr_cd", "inner")
            'w.WriteAttributeString("KEPYO_BClass", "0")
            'w.WriteAttributeString("AMO_Srl_cd", "Π000")
            'w.WriteAttributeString("ConstrCost", "0")
            'w.WriteAttributeString("AM_DcTp_Dscr", "Απόδειξη Λιανικής Πώλησης")
            'w.WriteAttributeString("AM_DcTp_cd", "#ΑΛΠ-0")
            'w.WriteAttributeString("Party_ISK_D_A_Dscr", "")
            'w.WriteAttributeString("Reserved", "0")
            'w.WriteAttributeString("AMO_Srl_DSCR", "ΠΩΛΗΣΕΙΣ (ΜΗΧΑΝΟΓΡΑΦΗΜΕΝΗ)")
            'w.WriteAttributeString("KepyoCatData_ISAGRYP", "0")
            'w.WriteAttributeString("System_Dscr_1", "Πωλήσεις")
            'w.WriteAttributeString("F_Coin_Dscr", "ΕURO")
            'w.WriteAttributeString("Party_CASTVAT_Dscr", "ΚΑΝΟΝΙΚΟ")
            'w.WriteAttributeString("Party_SNAME", "ΠΕΛΑΤΗΣ ΛΙΑΝΙΚΗΣ*")
            'w.WriteAttributeString("Party_ADDRESS", "")
            'w.WriteAttributeString("FL_Dgrs_Dscr", "ΓΕΝΙΚΗ ΛΟΓΙΣΤΙΚΗ / ΕΣΟΔΑ - ΕΞΟΔΑ")
            'w.WriteAttributeString("cdRetailIdentity", "ΑΝΕΥ")
            'w.WriteAttributeString("JrnCnt", "0")
            'w.WriteAttributeString("KepyoCatData_SUMKEPYOVAT", "19.2600")
            'w.WriteAttributeString("Party_CASTVAT", "1")
            'w.WriteAttributeString("CDt", "2015-04-29T06:02:56.6670000+03:00")
            'w.WriteAttributeString("F_Coin_ShCut", "€")
            'w.WriteAttributeString("F_Sites_cd", "001")
            'w.WriteAttributeString("authDt", "2015-04-29T06:03:11.8830000+03:00")
            'w.WriteAttributeString("GlbCff", "1")
            'w.WriteAttributeString("Dgrs", "1")
            'w.WriteAttributeString("Party_Sts", "1")
            'w.WriteAttributeString("AP_Party_Dscr", "ΠΕΛΑΤΗΣ ΛΙΑΝΙΚΗΣ*")
            'w.WriteAttributeString("APA_VIES_v_Dscr", "EL")
            'w.WriteAttributeString("fldinvoice", "0")
            'w.WriteAttributeString("Party_AFM", "000000000")
            'w.WriteAttributeString("Cnt", "1")
            'w.WriteAttributeString("Party_IDParty", "1")
            'w.WriteAttributeString("KepyoCatData_SUMKEPYOYP", "83.74")
            'w.WriteAttributeString("IsHand", "1")
            'w.WriteAttributeString("AP_Party_cd", "01")
            'w.WriteAttributeString("dt", "2015-04-29")
            'w.WriteAttributeString("XU_Usr_dscr", "inner")
            'w.WriteAttributeString("F_Sites_dscr", "ΚΕΝΤΡΙΚΟ")
            'w.WriteAttributeString("ExpenditureKind", "0")







        Else



            '<new 
            w.WriteAttributeString("GlbCff", "1")
            w.WriteAttributeString("authDt", Base_dt)
            w.WriteAttributeString("Reserved", "0")
            w.WriteAttributeString("CDt", Base_dt)
            w.WriteAttributeString("Cnt", mBase_INVOICE)
            w.WriteAttributeString("dscr", "")
            w.WriteAttributeString("dt", Base_dt)
            w.WriteAttributeString("JrnCnt", "0")
            w.WriteAttributeString("Party_ADDRESS", Party_SNAME) '==========
            w.WriteAttributeString("Party_JOB", "")
            w.WriteAttributeString("Party_DOY", "5111")
            w.WriteAttributeString("Party_CASTVAT_Dscr", "")
            w.WriteAttributeString("Party_ISK_D_A_Dscr", "")
            w.WriteAttributeString("Party_AFM", Trim(Party_AFM))
            w.WriteAttributeString("Party_SNAME", "-" + Party_SNAME) '==========
            w.WriteAttributeString("Party_IDParty", Party_IDParty) '  "1881")
            w.WriteAttributeString("KEPYO_BClass", "0")
            w.WriteAttributeString("KEPYO_Val", KAU_AJIA1)
            w.WriteAttributeString("AP_Party_Dscr", Party_SNAME) '==========
            w.WriteAttributeString("AP_Party_cd", "18-015")
            w.WriteAttributeString("F_Coin_Dscr", "ΕURO")
            w.WriteAttributeString("F_Coin_ShCut", "€")
            w.WriteAttributeString("FL_Dgrs_Dscr", "ΓΕΝΙΚΗ ΛΟΓΙΣΤΙΚΗ / ΕΣΟΔΑ - ΕΞΟΔΑ")
            w.WriteAttributeString("XU_Usr_dscr", "inner")
            w.WriteAttributeString("XU_Usr_cd", "inner")
            w.WriteAttributeString("AM_DcTp_Dscr", AM_DcTp_Dscr)
            w.WriteAttributeString("AM_DcTp_cd", tit_paras)
            w.WriteAttributeString("AMO_Srl_DSCR", f_aitiologia)
            w.WriteAttributeString("AMO_Srl_cd", "")
            w.WriteAttributeString("KepyoCatData_SUMKEPYOVAT", FPA1)
            w.WriteAttributeString("KepyoCatData_SUMKEPYOYP", KAU_AJIA1)
            w.WriteAttributeString("Party_CASTVAT", "1")
            w.WriteAttributeString("System_sys", "SB")


            w.WriteAttributeString("IsHand", IsHand)


        End If



        w.WriteEndElement() ' new />  w.WriteAttributeString("KepyoCatData_SUMKEPYOVAT", FPA1)
        w.WriteEndElement() ' /data
        '========================================================
        '<row name="Mv" mode="3" rowId="40" status="AllChildsQueried">
        '  <data>
        '    <new Comment="ΔΙΧΑΛΑ ΕΛΕΝΗ" Credit="697.2000" debit="0" Dscr="ΠΩΛΗΣ.ΕΜΠΟΡ.ΕΣΩΤΕΡ.ΧΟΝΔΡΙΚΑ ΦΠΑ 19%" cd="70-00-00-0077"/>
        '  </data>
        '</row>


        '========================================================
        w.WriteStartElement("detail")  'big
        If kau23 <> 0 Then
            rowG_detail(f_aitiologia, LOG23, kau23, 0, w)
            rowG_detail(f_aitiologia, LOGFPA23.Text, fpa23, 0, w)
            ' old  rowG_detail(f_aitiologia, lianLOGFPA23.Text, fpa23, 0, w)
        End If
        If kau13 <> 0 Then
            rowG_detail(f_aitiologia, LOG13, kau13, fpa13, w)
            rowG_detail(f_aitiologia, LOGFPA13.Text, fpa23, 0, w)

        End If
        If kau16 <> 0 Then
            rowG_detail(f_aitiologia, LOG16, kau16, 0, w)
            rowG_detail(f_aitiologia, LOGFPA16.Text, fpa16, 0, w)
        End If
        If kau9 <> 0 Then
            rowG_detail(f_aitiologia, LOG9, kau9, 0, w)
            rowG_detail(f_aitiologia, LOGFPA9.Text, fpa9, 0, w)
        End If
        If kau0 <> 0 Then
            rowG_detail("", LOG0, kau0, 0, w)
        End If
        rowG_detail(f_aitiologia, f_logPel, 0, kau23 + kau13 + kau16 + kau9 + fpa23 + fpa13 + fpa16 + fpa9, w)

        If Metrhtaxond = True Then  'metrhta

            rowG_detail(f_aitiologia, f_logPel, 0, kau23 + kau13 + kau16 + kau9 + fpa23 + fpa13 + fpa16 + fpa9, w)
            rowG_detail(f_aitiologia, f_logTam, kau23 + kau13 + kau16 + kau9 + fpa23 + fpa13 + fpa16 + fpa9, 0, w)

        End If




        w.WriteEndElement()  'detail  big
        '========================================================
        w.WriteEndElement() 'row hd
        'DDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDD

    End Sub


    Private Sub rowG_detail(ByVal mComment As String, ByVal mlog As String, ByVal mCredit As Single, ByVal mDebit As Single, ByVal w As XmlTextWriter)

        '<row name="Mv" mode="3" rowId="40" status="AllChildsQueried">
        '  <data>
        '    <new Comment="ΔΙΧΑΛΑ ΕΛΕΝΗ" 
        '    Credit="697.2000" 
        '    debit="0" '
        '    Dscr = "ΠΩΛΗΣ.ΕΜΠΟΡ.ΕΣΩΤΕΡ.ΧΟΝΔΡΙΚΑ ΦΠΑ 19%"
        '    cd="70-00-00-0077"/>
        '  </data>
        '</row>

        '--------- Mv -------------------------------------
        w.WriteStartElement("row")
        : w.WriteAttributeString("rowid", LTrim(Str(rowId)))
        : w.WriteAttributeString("mode", "3")
        : w.WriteAttributeString("name", "Mv")
        : w.WriteAttributeString("status", "AllChildsQueried")
        FL_Ledg_cd = mlog ' pol13.Text

        w.WriteStartElement("data") '''''''''''''''''''''''''''''
        w.WriteStartElement("new")
        w.WriteAttributeString("Comment", mComment)
        w.WriteAttributeString("Credit", mCredit)
        w.WriteAttributeString("debit", mDebit)
        w.WriteAttributeString("Dscr", FL_Ledg_Dscr)
        w.WriteAttributeString("cd", mlog)
        w.WriteEndElement()  'new
        w.WriteEndElement()  'data   '''''''''''''''''''''''''''''

        w.WriteEndElement()  'row mv
        '---------------------------------------------------


    End Sub




    '    Parsing XML files has always been time consuming and sometimes tricky. .NET framework provides powerful new ways of parsing XML. The various techniques know to parse xml files with .NET framework are using XmlTextReader, XmlDocument, XmlSerializer, DataSet and XpathDocument. I will explore the XmlTextReader and XmlDocument approach here.

    'The Xml File
    'Figure 1 outlines the xml file that will be parsed.

    'Hide   Copy Code

    '<?xml version="1.0" encoding="UTF-8"?>
    '<family>
    '  <name gender="Male">
    '    <firstname>Tom</firstname>
    '    <lastname>Smith</lastname>
    '  </name>
    '  <name gender="Female">
    '    <firstname>Dale</firstname>
    '    <lastname>Smith</lastname>
    '  </name>
    '</family>
    'Figure1: Xml file



    'Parsing XML with XMLTextReader
    'Using XmlTextReader is appropriate when the structure of the XML file is relatively simple. Parsing with XmlTextReader gives you a pre .net feel as you sequentially walk through the file using Read() and get data using GetAttribute() andReadElementString() methods. Thus while using XmlTextReader it is up to the developer to keep track where he is in the Xml file and Read() correctly. Figure 2 below outlines parsing of xml file with XmlTextReader

    'Hide   Shrink    Copy Code
    'Imports System.IO
    'Imports System.Xml
    'Module ParsingUsingXmlTextReader
    '        Sub Main()
    '            Dim m_xmlr As XmlTextReader
    '            'Create the XML Reader
    '            m_xmlr = New XmlTextReader("C:\Personal\family.xml")
    '            'Disable whitespace so that you don't have to read over whitespaces
    '            m_xmlr.WhiteSpaceHandling = WhiteSpaceHandling.NONE
    '            'read the xml declaration and advance to family tag
    '            m_xmlr.Read()
    '            'read the family tag
    '            m_xmlr.Read()
    '            'Load the Loop
    '            While Not m_xmlr.EOF
    '                'Go to the name tag
    '                m_xmlr.Read()
    '                'if not start element exit while loop
    '                If Not m_xmlr.IsStartElement() Then
    '                    Exit While
    '                End If
    '                'Get the Gender Attribute Value
    '                Dim genderAttribute = m_xmlr.GetAttribute("gender")
    '                'Read elements firstname and lastname
    '                m_xmlr.Read()
    '                'Get the firstName Element Value
    '                Dim firstNameValue = m_xmlr.ReadElementString("firstname")
    '                'Get the lastName Element Value
    '                Dim lastNameValue = m_xmlr.ReadElementString("lastname")
    '                'Write Result to the Console
    '                Console.WriteLine("Gender: " & genderAttribute _
    '                  & " FirstName: " & firstNameValue & " LastName: " _
    '                  & lastNameValue)
    '                Console.Write(vbCrLf)
    '            End While
    '            'close the reader
    '            m_xmlr.Close()
    '        End Sub
    '    End Module
    'Figure 2: Xml Parsing with XmlTextReader

    'Parsing XML with XmlDocument
    'The XmlDocument class is modeled based on Document Object Model. 
    'XmlDocument class is appropriate if you need to extract data in a non-sequential manner. 
    'Figure 3 below outlines parsing of xml file with XmlDocument

    'Hide   Shrink    Copy Code
    'Imports System.IO
    'Imports System.Xml
    'Module ParsingUsingXmlDocument
    '        Sub Main()
    '            Try
    '                Dim m_xmld As XmlDocument
    '                Dim m_nodelist As XmlNodeList
    '                Dim m_node As XmlNode
    '                'Create the XML Document
    '                m_xmld = New XmlDocument()
    '                'Load the Xml file
    '                m_xmld.Load("C:\CMS\Personal\family.xml")
    '                'Get the list of name nodes 
    '                m_nodelist = m_xmld.SelectNodes("/family/name")
    '                'Loop through the nodes
    '                For Each m_node In m_nodelist
    '                    'Get the Gender Attribute Value
    '                    Dim genderAttribute = m_node.Attributes.GetNamedItem("gender").Value
    '                    'Get the firstName Element Value
    '                    Dim firstNameValue = m_node.ChildNodes.Item(0).InnerText
    '                    'Get the lastName Element Value
    '                    Dim lastNameValue = m_node.ChildNodes.Item(1).InnerText
    '                    'Write Result to the Console
    '                    Console.Write("Gender: " & genderAttribute _
    '                      & " FirstName: " & firstNameValue & " LastName: " _
    '                      & lastNameValue)
    '                    Console.Write(vbCrLf)
    '                Next
    '            Catch errorVariable As Exception
    '                'Error trapping
    '                Console.Write(errorVariable.ToString())
    '            End Try
    '        End Sub










    'Private Sub aaaa() 'Button7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles agoresB.Click, agoresB.Click

    Private Sub agoresB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles agoresB.Click
        'AGORES 

        Dim a As String
        Dim K As Short
        Dim C As String

        ' CO TO DIAXORISTIKO DEKADIKON ARITMON
        Dim CO As String = String.Format(1.1).Substring(1, 1)


        ' MsgBox("ΠΡΟΣΟΧΗ ΔΙΑΒΑΖΕΙ ΑΠΟ ΤΗΝ 2η ΣΕΙΡΑ ΜΕ ΓΡΑΜΟΓΡΑΦΗΣΗ:" + Chr(13) + "AJ1 AJ2	AJ3	AJ4	AJ5	AJI	FPA1	FPA2	FPA3	FPA4	ATIM	HME	EPO	AFM")


        ' Write the string as utf-8.
        ' This also writes the 3-byte utf-8 preamble at the beginning of the file.
        Dim appendMode As Boolean = False ' This overwrites the entire file.
        ' Dim sw As New StreamWriter("C:\MERCVB\out_utf9.export", appendMode, System.Text.Encoding.UTF8)
        'sw.Write(TextBox1.Text)
        'sw.Close()


        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xl As Excel.Worksheet

        xlApp = New Excel.ApplicationClass


        If Len(TextBox1.Text) < 2 Then
            Exit Sub
        End If

        xlWorkBook = xlApp.Workbooks.Open(TextBox1.Text)
        xl = xlWorkBook.Worksheets(1) ' .Add




        '====================================================================================
        Dim ff As String = "c:\mercvb\m" + VB6.Format(Now, "YYYYddmmHHMM") + ".export" ' "\\Logisthrio\333\pr.export" '

        Dim writer As New XmlTextWriter(ff, System.Text.Encoding.UTF8)
        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2
        ' Create a Continent element and set its value to
        ' that of the New Continent dialog box
        'writer.WriteAttributeString("Table", , "sadasasd sdsd")

        '<Data Name="SX" Style="BRowse"><row name="Hd" mode="3" rowId="7">
        writer.WriteStartElement("Data")
        writer.WriteAttributeString("Name", "SX")
        writer.WriteAttributeString("Style", "Browse")
        '====================================================================================


        Dim enter_Renamed As String
        enter_Renamed = Chr(13)

        'FileOpen(1, "C:\MERCVB\A778.XML", OpenMode.Output)
        ROW = Val(ApoSeira.Text) - 1

        Dim hand As Integer = 0

        fnTimol = Val(nTimol.Text)
        fnLian = Val(nLian.Text)
        fnPistTim = Val(nPistTim.Text)
        fnPistLian = Val(nPistLian.Text)
        ' As Integer
        fcTimol = cTimol.Text
        fcLian = cLian.Text
        fcPistTim = cPistTim.Text
        fcPistLian = cPistLian.Text


        '===============================================================================real onomatepvmymo 54100
        Do While True
            ROW = ROW + 1

            Me.Text = ROW
            'system.doevents

            If IsDBNull(xl.Cells(ROW, 12).value) Then
                Exit Do
            End If

            If Len(xl.Cells(ROW, 11).ToString) < 2 Then
                Exit Do
            End If
            If xl.Cells(ROW, 11).value = Nothing Then
                Exit Do
            End If

            '1	 2	    3	4	5	6	7	    8	    9	    10	    11	    12	13	14	15	16	17	    18	19
            'AJ1 AJ2	AJ3	AJ4	AJ5	AJI	FPA1	FPA2	FPA3	FPA4	ATIM	HME	EPO	AFM	KPE	DIE	XRVMA	EPA	POL
            Party_IDParty = xl.Cells(ROW, 14).value  ' As String  '12344   ΚΩΔ ΣΥΝΑΛΛΑΣΟΜΕΝΟΥ
            AM_DcTp_Dscr = "Τιμολόγιο"
            Party_AFM = Trim(xl.Cells(ROW, 14).value)  'Dim Party_AFM As String ' =""999349996
            If Len(Trim(Party_AFM)) <= 4 Then
                Party_AFM = "000000000"
            End If

            Party_ADDRESS = xl.Cells(ROW, 16).value 'ToString  ' "ΠΟΛΥΣΤΗΛΟ ΚΑΒΑΛΑΣ"
            AM_DcTp_cd = "#ΤΥΠ-0"
            AMO_Srl_DSCR = "Πωλήσεις" '"ΠΩΛΗΣΕΙΣ"
            Base_dt = VB6.Format(xl.Cells(ROW, 12), "YYYY-mm-dd")
            Base_INVOICE = xl.Cells(ROW, 11).value  ' =""#ΤΥΠ-0/Π000/1/Τιμολόγιο Παροχής Υπηρεσιών"
            Party_SNAME = xl.Cells(ROW, 13).value  '=""Θ. ΓΡΑΜΜΑΤΗΣ Κ.ΣΙΑ Ε.Ε""
            f_logPel = pel30.Text ' "30-00-00-0000"

            KAU_AJIA = nVal(xl.Cells(ROW, 1).value) + nVal(xl.Cells(ROW, 2).value) + nVal(xl.Cells(ROW, 3).value) + nVal(xl.Cells(ROW, 4).value) + nVal(xl.Cells(ROW, 5).value)
            FPA = nVal(xl.Cells(ROW, 7).value) + nVal(xl.Cells(ROW, 8).value) + nVal(xl.Cells(ROW, 9).value) + nVal(xl.Cells(ROW, 10).value)


            kau13 = nVal(xl.Cells(ROW, 1).value)
            kau23 = nVal(xl.Cells(ROW, 2).value)
            kau16 = nVal(xl.Cells(ROW, 3).value)
            kau9 = nVal(xl.Cells(ROW, 4).value)
            kau0 = nVal(xl.Cells(ROW, 5).value)

            fpa13 = nVal(xl.Cells(ROW, 7).value)
            fpa23 = nVal(xl.Cells(ROW, 8).value)
            fpa16 = nVal(xl.Cells(ROW, 9).value)
            fpa9 = nVal(xl.Cells(ROW, 10).value)

            LOG13 = pol13.Text : LOG23 = pol23.Text
            LOG16 = POL16.Text : LOG9 = POL9.Text
            LOG0 = POL0.Text



            FL_Ledg_Dscr = "ΠΩΛΗΣΕΙΣ ΧΟΝΔΡΙΚΗΣ ΕΣ. ΦΠΑ23%"
            FL_Ledg_cd = pol23.Text ' "70-00-00-0057"

            MVTP = "1" '6=πιστωτικα 2=αγορες  7=πιστωτικα αγορών
            System_sys = "SB" '      'SB =POLISEIS FR



            If InStr(fcLian, Mid(Base_INVOICE, 1, fnLian)) > 0 Then
                IsHand = "1" 'LTrim(Str(hand))
                cdRetailIdentity = ""
                LOG13 = Lian13.Text : LOG23 = Lian23.Text
                LOG0 = LIAN0.Text
                If InStr("yπ", Mid(Base_INVOICE, 1, 1)) > 0 Then 'yphresies
                    LOG23 = "73-4057"
                End If
                If InStr("πφ", Mid(Base_INVOICE, 1, 1)) > 0 Then 'yphresies
                    cdRetailIdentity = "ΣΥ09002067"
                    LOG23 = "73-4057"
                    IsHand = "0" 'LTrim(Str(hand))
                End If
                MVTP = "1" '6=πιστωτικα 2=αγορες  7=πιστωτικα αγορών



            ElseIf InStr("GΓgΞμ", Mid(Base_INVOICE, 1, 1)) > 0 Then
                'θελει ψαξιμο.................
                MVTP = 2
                IsHand = "" 'LTrim(Str(hand))  BP=AGORES
                cdRetailIdentity = ""
                LOG13 = ago13.Text : LOG23 = ago23.Text
                LOG16 = ago16.Text : LOG9 = ago9.Text
                'LOG0 = LIAN0.Text
                'System_sys = "BP"

                'System_Dscr_1
                AMO_Srl_DSCR = "Αγορές"
                AMO_Srl_DSCR = "ΑΓΟΡΕΣ (ΧΕΙΡΟΓΡΑΦΗ)"
                System_sys = "BP"

            ElseIf InStr("D", Mid(Base_INVOICE, 1, 1)) > 0 Then
                'θελει ψαξιμο.................
                MVTP = 7
                IsHand = "" 'LTrim(Str(hand))  BP=AGORES
                cdRetailIdentity = ""
                LOG13 = ago13.Text : LOG23 = ago23.Text
                LOG16 = ago16.Text : LOG9 = ago9.Text
                'LOG0 = LIAN0.Text
                System_sys = "BP"

                'System_Dscr_1
                AMO_Srl_DSCR = "Αγορές"
                AMO_Srl_DSCR = "ΑΓΟΡΕΣ (ΧΕΙΡΟΓΡΑΦΗ)"
                System_sys = "BP"

            ElseIf InStr(fcTimol, Mid(Base_INVOICE, 1, fnTimol)) > 0 Then 'τιμολογια -πιστωτικά
                cdRetailIdentity = ""
                IsHand = ""
                'LOG13 = pol13.Text : LOG23 = pol23.Text
                'LOG16 = POL16.Text : LOG9 = POL9.Text
                'LOG0 = POL0.Text
                MVTP = "1" '6=πιστωτικα 2=αγορες  7=πιστωτικα αγορών



                'ElseIf InStr("Y", Mid(Base_INVOICE, 1, 1)) > 0 Then 'yphresies
                '   LOG23 = "73-0057"
                '  MVTP = "1" '6=πιστωτικα 2=αγορες  7=πιστωτικα αγορών
            ElseIf InStr(fcPistLian, Mid(Base_INVOICE, 1, fnPistLian)) > 0 Then  'επιστροφη λιανικης
                LOG13 = episLian13.Text : LOG23 = episLian23.Text
                MVTP = 6
                IsHand = "1" 'LTrim(Str(hand))
                'kau13 = -kau13
                'kau23 = -kau23
                'fpa13 = -fpa13
                'fpa23 = -fpa23
                'KAU_AJIA = -KAU_AJIA
                'FPA = -FPA
                System_sys = "SB" '      'SB =POLISEIS FR
                'System_sys = "FR"

                'End If
            ElseIf InStr(fcPistTim, Mid(Base_INVOICE, 1, fnPistTim)) > 0 Then  'pistvtiko timologio
                'kau13 = -kau13
                'kau23 = -kau23
                'fpa13 = -fpa13
                'fpa23 = -fpa23
                'KAU_AJIA = -KAU_AJIA
                'FPA = -FPA
                LOG13 = EPIS13.Text : LOG23 = EPIS23.Text
                MVTP = 6
                'End If
                System_sys = "SB" '      'SB =POLISEIS FR

            End If
            KAU_AJIA1 = KAU_AJIA
            FPA1 = FPA
            writeBAgor_row(writer)
            rowId = rowId + 11
        Loop





        writer.WriteEndDocument()
        writer.Close()





        MsgBox("Δημιουργήθηκε στο " + ff)
        xlApp.Quit()

    End Sub
    ''======================================================================================
    Sub writeBAgor_row(ByVal w As XmlTextWriter)
        'DDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDD


        '     <row name="Hd" mode="3" rowId="7">
        '<data>
        '<new 
        'Party_ADDRESS = ""
        'IsHand = ""
        'Party_ISK_D_A_CD = "0"
        'GlbCff = "1"
        'APA_VIES_v_Dscr = "EL"
        'Party_AFM = "028783755"
        'KepyoCatData_SUMKEPYOVAT = "92.0000"
        'KepyoCatData_ISAGRYP = "0"
        'Party_ISK_D_A_Dscr = "ΚΑΝΟΝΙΚΟΣ"
        'ExpenditureKind = "0"
        'AMO_Srl_cd = "ΑΓ00"
        'System_Dscr_1 = "Αγορές"
        'AM_DcTp_cd = "#ΤΑΓ-0"
        'KepyoCatData_SUMKEPYOYP = "400.0000"
        'cdRetailIdentity = ""
        'Ledger_Supl = "50-00-00-0000"
        'Party_Sts = "1"
        'Party_IDParty = "3"
        'Party_DOY = "1104"
        'Base_dt = "2015-07-31"
        'Party_CASTVAT_Dscr = "ΚΑΝΟΝΙΚΟ"
        'DocCd = "12345"

        'F_Sites_dscr = "ΚΕΝΤΡΙΚΟ"
        'Party_SNAME = "lagakis dokimastikos"
        'System_sys = "BP"
        'dumm = "0"
        'AMO_Srl_DSCR = "ΑΓΟΡΕΣ (ΧΕΙΡΟΓΡΑΦΗ)"
        'ConstrCost = "0"
        'fldinvoice = "0"
        'Base_INVOICE = "#ΤΑΓ-0/ΑΓ00/120/Τιμολόγιο Αγοράς - Δελτίο Αποστολής"
        'AM_DcTp_Dscr = "Τιμολόγιο Αγοράς - Δελτίο Αποστολής"
        'Party_CASTVAT = "1"
        'F_Sites_cd="001" />

        '</data><detail><row name="Mv" mode="3" rowId="7"><data><new RegVatVal="0.0000" NetVal="400.0000" RegVal="400.0000" VatVal="92.0000" MvTp="2" FL_Ledg_cd="20-00-00-0057" FL_Ledg_Dscr="ΑΓΟΡΕΣ ΕΜΠΟΡΕΥΜΑΤΩΝ ΕΣΩΤΕΡΙΚΟΥ 23%" /></data><detail><row name="Ledg" mode="3" rowId="7"><data><new Anali="0" Active="1" CanMv="1" cdLedg="20-00-00-0057" dscrLedg="ΑΓΟΡΕΣ ΕΜΠΟΡΕΥΜΑΤΩΝ ΕΣΩΤΕΡΙΚΟΥ 23%" /></data></row><row name="Ledg" mode="3" rowId="18"><data><new Anali="0" Active="1" CanMv="0" cdLedg="20-00-00" dscrLedg="ΑΓΟΡΕΣ ΧΡΗΣΕΩΣ ΕΣΩΤΕΡΙΚΟΥ" /></data></row><row name="Ledg" mode="3" rowId="29">
        '<data><new Anali="0" Active="1" CanMv="0" cdLedg="20-00" dscrLedg="ΕΜΠΟΡΕΥΜΑΤΑ" /></data></row><row name="Ledg" mode="3" rowId="40"><data><new Anali="0" Active="1" CanMv="0" cdLedg="20" dscrLedg="ΕΜΠΟΡΕΥΜΑΤΑ" /></data></row></detail></row></detail></row>
























        'big row
        w.WriteStartElement("row") : w.WriteAttributeString("rowid", LTrim(Str(rowId))) : w.WriteAttributeString("mode", "3") : w.WriteAttributeString("name", "Hd")


        '========================================================
        w.WriteStartElement("data")
        w.WriteStartElement("new")
        w.WriteAttributeString("Party_ADDRESS", IsHand)
        w.WriteAttributeString("IsHand", IsHand)
        w.WriteAttributeString("Party_ISK_D_A_CD", "0")
        w.WriteAttributeString("GlbCff", "1")
        w.WriteAttributeString("APA_VIES_v_Dscr", "EL")
        w.WriteAttributeString("Party_AFM", Trim(Party_AFM))
        w.WriteAttributeString("KepyoCatData_SUMKEPYOVAT", Replace(Str(FPA1), ",", "."))  'FPA1)
        w.WriteAttributeString("KepyoCatData_ISAGRYP", "0")
        w.WriteAttributeString("Party_ISK_D_A_Dscr", "ΚΑΝΟΝΙΚΟΣ")
        w.WriteAttributeString("ExpenditureKind", "0")
        w.WriteAttributeString("AMO_Srl_cd", "ΑΓ00")

        w.WriteAttributeString("System_Dscr_1", "Αγορές") 'AM_DcTp_Dscr)
        w.WriteAttributeString("AM_DcTp_cd", "#ΤΑΓ-0")
        ' ="400.0000" cdRetailIdentity="" 
        w.WriteAttributeString("KepyoCatData_SUMKEPYOYP", Replace(Str(KAU_AJIA1), ",", ".")) 'KAU_AJIA1)
        w.WriteAttributeString("cdRetailIdentity", cdRetailIdentity)

        w.WriteAttributeString("Ledger_Supl", pel30.Text)  'neo 
        w.WriteAttributeString("Party_Sts", "1")
        w.WriteAttributeString("Party_IDParty", Mid(Party_AFM, 5, 5)) ' κωδικος πελατη 13 
        w.WriteAttributeString("Party_DOY", "")
        w.WriteAttributeString("Base_dt", Base_dt)
        w.WriteAttributeString("Party_CASTVAT_Dscr", "ΚΑΝΟΝΙΚΟ")
        w.WriteAttributeString("DocCd", "12345")


        w.WriteAttributeString("F_Sites_dscr", "ΚΕΝΤΡΙΚΟ")
        w.WriteAttributeString("Party_SNAME", Party_SNAME)

        w.WriteAttributeString("System_sys", "BP")

        w.WriteAttributeString("dumm", "0")
        w.WriteAttributeString("AMO_Srl_DSCR", "ΠΛΗΡΩΜΩΝ - ΠΡΟΜΗΘΕΥΤΩΝ")
        w.WriteAttributeString("ConstrCost", "0")
        w.WriteAttributeString("fldinvoice", "0") 'NEO neo
        w.WriteAttributeString("Base_INVOICE", Base_INVOICE)
        w.WriteAttributeString("AM_DcTp_Dscr", AM_DcTp_Dscr)
        w.WriteAttributeString("Party_CASTVAT", "1")
        w.WriteAttributeString("F_Sites_cd", "001")




        'w.WriteAttributeString("Party_Sts", "1") ' neo
        'w.WriteAttributeString("Party_PHONE2", "")
        'w.WriteAttributeString("Party_Zip", "")  '66100 
        'w.WriteAttributeString("KepyoCatData_SUMKEPYOYP", Replace(Str(KAU_AJIA1), ",", ".")) 'KAU_AJIA1)
        'w.WriteAttributeString("Party_City", "δραμα")
        'w.WriteAttributeString("Party_PHONE1", "")
        'w.WriteAttributeString("System_sys", System_sys) 'SB =POLISEIS FR PISTVTIKA YPIRESIES FP= PLIROMES BP=AGORES
        'w.WriteAttributeString("Party_ADDRESS", Party_ADDRESS)
        'w.WriteAttributeString("Party_JOB", "εμπορια")





        w.WriteEndElement() ' new />
        w.WriteEndElement() ' /data
        '========================================================
        '        <detail>
        '<row name="Mv" mode="3" rowId="7">
        '<data>
        ' <new RegVatVal="0.0000" NetVal="400.0000" RegVal="400.0000" VatVal="92.0000" MvTp="2" FL_Ledg_cd="20-00-00-0057" FL_Ledg_Dscr="ΑΓΟΡΕΣ ΕΜΠΟΡΕΥΜΑΤΩΝ ΕΣΩΤΕΡΙΚΟΥ 23%"/>
        ' </data>
        '<detail>
        '<row name="Ledg" mode="3" rowId="7">
        '<data>
        ' <new Anali="0" Active="1" CanMv="1" cdLedg="20-00-00-0057" dscrLedg="ΑΓΟΡΕΣ ΕΜΠΟΡΕΥΜΑΤΩΝ ΕΣΩΤΕΡΙΚΟΥ 23%"/>
        ' </data>
        ' </row>
        '<row name="Ledg" mode="3" rowId="18">
        '<data>
        ' <new Anali="0" Active="1" CanMv="0" cdLedg="20-00-00" dscrLedg="ΑΓΟΡΕΣ ΧΡΗΣΕΩΣ ΕΣΩΤΕΡΙΚΟΥ"/>
        ' </data>
        ' </row>
        '<row name="Ledg" mode="3" rowId="29">
        '<data>
        ' <new Anali="0" Active="1" CanMv="0" cdLedg="20-00" dscrLedg="ΕΜΠΟΡΕΥΜΑΤΑ"/>
        ' </data>
        ' </row>
        '<row name="Ledg" mode="3" rowId="40">
        '<data>
        ' <new Anali="0" Active="1" CanMv="0" cdLedg="20" dscrLedg="ΕΜΠΟΡΕΥΜΑΤΑ"/>
        ' </data>
        ' </row>
        ' </detail>
        ' </row>
        ' </detail>


        '========================================================
        w.WriteStartElement("detail")  'big
        If kau23 <> 0 Then
            rowBAgor_detail(LOG23, kau23, fpa23, w)
        End If
        If kau13 <> 0 Then
            rowBAgor_detail(LOG13, kau13, fpa13, w)
        End If

        If kau24 <> 0 Then
            rowBAgor_detail(LOG24, kau24, fpa24, w)
        End If
        If kau17 <> 0 Then
            rowBAgor_detail(LOG17, kau17, fpa17, w)
        End If





        If kau16 <> 0 Then
            rowBAgor_detail(LOG16, kau16, fpa16, w)
        End If
        '  Exit Sub
        If kau9 <> 0 Then
            rowBAgor_detail(LOG9, kau9, fpa9, w)
        End If

        If kau0 <> 0 Then
            rowBAgor_detail(LOG0, kau0, 0, w)
        End If




        w.WriteEndElement()  'detail  big
        '========================================================


        w.WriteEndElement() 'row hd
        'DDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDD

    End Sub


    Private Sub rowBAgor_detail(ByVal mlog As String, ByVal mKau As Single, ByVal mFpa As Single, ByVal w As XmlTextWriter)

        FileOpen(11, "LOG.TXT", OpenMode.Append)

        'Type Visual Basic 6 code here...
        WriteLine(11, mlog + ";" + Str(mKau) + ";" + Str(mFpa))
        FileClose(11)





        '<row name="Mv" mode="3" rowId="7">
        '<data>
        ' <new RegVatVal="0.0000" 
        'NetVal="400.0000" 
        'RegVal = "400.0000"
        'VatVal="92.0000" 
        'MVTP = "2"
        'FL_Ledg_cd="20-00-00-0057" 
        'FL_Ledg_Dscr="ΑΓΟΡΕΣ ΕΜΠΟΡΕΥΜΑΤΩΝ ΕΣΩΤΕΡΙΚΟΥ 23%"/>
        ' </data>


        '--------- Mv -------------------------------------
        w.WriteStartElement("row") : w.WriteAttributeString("name", "Mv") : w.WriteAttributeString("mode", "3") : w.WriteAttributeString("rowid", LTrim(Str(rowId)))



        FL_Ledg_cd = mlog ' pol13.Text
        w.WriteStartElement("data") '''''''''''''''''''''''''''''


        w.WriteStartElement("new")
        w.WriteAttributeString("RegVatVal", "0.000")
        w.WriteAttributeString("NetVal", Replace(Str(mKau), ",", "."))
        w.WriteAttributeString("RegVal", Replace(Str(mKau), ",", "."))
        w.WriteAttributeString("VatVal", Replace(Str(mFpa), ",", "."))
        w.WriteAttributeString("MvTp", MVTP)
        w.WriteAttributeString("FL_Ledg_cd", FL_Ledg_cd)
        w.WriteAttributeString("FL_Ledg_Dscr", FL_Ledg_Dscr)
        w.WriteEndElement()  '/data




        'w.WriteStartElement("new")
        'w.WriteAttributeString("FL_Ledg_Dscr", FL_Ledg_Dscr)
        'w.WriteAttributeString("FL_Ledg_cd", FL_Ledg_cd)
        'w.WriteAttributeString("VatVal", Replace(Str(mFpa), ",", "."))
        'w.WriteAttributeString("NetVal", Replace(Str(mKau), ",", "."))
        'w.WriteAttributeString("RegVal", Replace(Str(mKau), ",", "."))
        'w.WriteAttributeString("MvTp", MVTP)
        'w.WriteAttributeString("RegVatVal", "0.000")
        'w.WriteEndElement()  'new

        w.WriteStartElement("detail")  'detail
        row_ledg("1", FL_Ledg_cd, w) '4b
        rowIdINNER = rowIdINNER + 11
        If Len(FL_Ledg_cd) = 7 Then ' 70-0057
            row_ledg("0", Mid(FL_Ledg_cd, 1, 2), w) '3b
            rowIdINNER = rowIdINNER + 11
        Else  '70-00-00-0057
            row_ledg("0", Mid(FL_Ledg_cd, 1, 8), w) '3b
            rowIdINNER = rowIdINNER + 11
            row_ledg("0", Mid(FL_Ledg_cd, 1, 5), w) '2b
            rowIdINNER = rowIdINNER + 11
            row_ledg("0", Mid(FL_Ledg_cd, 1, 2), w) '1b
            rowIdINNER = rowIdINNER + 11
        End If
        w.WriteEndElement()  'detail   


        w.WriteEndElement()  'data   '''''''''''''''''''''''''''''



        w.WriteEndElement()  'row mv
        '---------------------------------------------------

    End Sub




    '    End Sub


    Private Sub mercury_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mercury.Click
        '======================================================================================================
        ' mercury

        Dim a As String
        Dim K As Short
        Dim C As String


        Dim sb13 As Single = 0
        Dim sb23 As Single = 0
        Dim sb24 As Single = 0


        Dim sb17 As Single = 0
        Dim sb16 As Single = 0
        Dim sb9 As Single = 0



        Dim esb13 As Single = 0
        Dim esb23 As Single = 0

        Dim esb24 As Single = 0
        Dim esb17 As Single = 0


        Dim bp13 As Single = 0
        Dim bp23 As Single = 0
        Dim bp0 As Single = 0


        Dim sb0 As Single = 0
        Dim esb0 As Single = 0


        ' CO TO DIAXORISTIKO DEKADIKON ARITMON
        Dim CO As String = String.Format(1.1).Substring(1, 1)


        ' MsgBox("ΠΡΟΣΟΧΗ ΔΙΑΒΑΖΕΙ ΑΠΟ ΤΗΝ 2η ΣΕΙΡΑ ΜΕ ΓΡΑΜΟΓΡΑΦΗΣΗ:" + Chr(13) + "AJ1 AJ2	AJ3	AJ4	AJ5	AJI	FPA1	FPA2	FPA3	FPA4	ATIM	HME	EPO	AFM")


        ' Write the string as utf-8.
        ' This also writes the 3-byte utf-8 preamble at the beginning of the file.
        Dim appendMode As Boolean = False ' This overwrites the entire file.
        ' Dim sw As New StreamWriter("C:\MERCVB\out_utf9.export", appendMode, System.Text.Encoding.UTF8)
        'sw.Write(TextBox1.Text)
        'sw.Close()
        If checkServer() = False Then
            MsgBox("αποτυχία ενημέρωσης")
            Exit Sub
        End If


        Dim pol As String = " "
        Dim polepis As String = " "
        Dim ago As String = " "
        Dim AGOEPIS As String = " "



        Get_AJ_ASCII(pol, polepis, ago, AGOEPIS)



        '   Dim xlApp As Excel.Application
        '   Dim xlWorkBook As Excel.Workbook
        '    Dim xl As Excel.Worksheet

        '   xlApp = New Excel.ApplicationClass
        Dim par As String = " "
        Dim mf As String
        mf = "c:\mercvb\err3.txt"
        If Len(Dir(UCase(mf))) = 0 Then
            par = pol '  " 'G','g','Ξ','D'  "
            par = InputBox("ΠΑΡΑΣΤΑΤΙΚΑ", , par)
        Else
            FileOpen(1, mf, OpenMode.Input)
            '   Input(1, par)
            par = LineInput(1)
            FileClose(1)
        End If

        par = InputBox("ΠΑΡΑΣΤΑΤΙΚΑ", , par)

        FileOpen(1, mf, OpenMode.Output)
        PrintLine(1, par)
        FileClose(1)


        Dim synt As String
        If epan.CheckState = CheckState.Checked Then
            synt = ""

        Else
            synt = " and (B_C1 is null or LEFT(B_C1,1)<>'*') "

        End If
        ExecuteSQLQuery("update TIM SET AJ7=0 WHERE AJ7 IS NULL")

        '  Dim XL As DataTable
        Dim SQL As String   '   ID_NUM GEMISMA NA JEKINA APO 1
        SQL = "SELECT ID_NUM, AJ1  ,AJ2 , AJ3,AJ4,AJ5,AJI,FPA1,FPA2,FPA3,FPA4,ATIM,"
        SQL = SQL + "HME,PEL.EPO,PEL.AFM,KPE,PEL.DIE,PEL.XRVMA"    '"CONVERT(CHAR(10),HME,3) AS HMEP
        SQL = SQL + ",PEL.EPA,PEL.POL,AJ6,FPA6,AJ7,FPA7"

        SQL = SQL + "   FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD "
        SQL = SQL + " WHERE LEFT(ATIM,1) IN     (  " + par + "  )    and HME>='" + VB6.Format(apo, "mm/dd/yyyy") + "'  AND HME<='" + VB6.Format(eos, "mm/dd/yyyy") + "'  "
        SQL = SQL + "  AND AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7>0  " + synt
        SQL = SQL + " order by HME"





        '  SQL = "SELECT  top 20  AJ1 ,AJ2  from TIM  order by HME"

        ExecuteSQLQuery(SQL)

        If sqlDT.Rows.Count = 0 Then
            MsgBox("ΔΕΝ ΒΡΕΘΗΚΑΝ ΕΓΓΡΑΦΕΣ")
            Exit Sub
        End If


        If Len(TextBox1.Text) < 2 Then
            '  Exit Sub
        End If

        ' xlWorkBook = xlApp.Workbooks.Open(TextBox1.Text)
        '  xl = xlWorkBook.Worksheets(1) ' .Add




        '====================================================================================
        Dim ff As String = "c:\mercvb\m" + VB6.Format(Now, "YYYYddmmHHMM") + ".export" ' "\\Logisthrio\333\pr.export" '

        Dim writer As New XmlTextWriter(ff, System.Text.Encoding.UTF8)
        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2
        ' Create a Continent element and set its value to
        ' that of the New Continent dialog box
        'writer.WriteAttributeString("Table", , "sadasasd sdsd")

        '<Data Name="SX" Style="BRowse"><row name="Hd" mode="3" rowId="7">
        writer.WriteStartElement("Data")
        writer.WriteAttributeString("Name", "SX")
        writer.WriteAttributeString("Style", "Browse")
        '====================================================================================


        Dim enter_Renamed As String
        enter_Renamed = Chr(13)

        'FileOpen(1, "C:\MERCVB\A778.XML", OpenMode.Output)
        ROW = Val(ApoSeira.Text) - 1

        Dim hand As Integer = 0

        fnTimol = Val(nTimol.Text)
        fnLian = Val(nLian.Text)
        fnPistTim = Val(nPistTim.Text)
        fnPistLian = Val(nPistLian.Text)

        fnPistAg = Val(nPistAg.Text)
        fnTimAg = Val(nTimAg.Text)
        fnPAR = Val(nParox.Text)




        ' As Integer
        fcTimol = cTimol.Text
        fcLian = cLian.Text
        fcPistTim = cPistTim.Text
        fcPistLian = cPistLian.Text

        fcTimAg = cTimAg.Text
        fcPistAg = cPistAg.Text

        fcPAR = cParox.Text



        Dim ajia_ana_parast(30) As Single
        Dim parast(30) As String
        Dim OK, i, nSynal As Integer
        nSynal = 0
        '===============================================================================real onomatepvmymo 54100
        'Do While True
        'ROW = ROW + 1
        For ROW = 0 To sqlDT.Rows.Count - 1
            Me.Text = ROW
            kau13 = 0
            kau23 = 0
            kau16 = 0
            kau9 = 0
            kau0 = 0
            kau24 = 0
            kau13 = 0

            '    If IsDBNull(sqlDT.Rows(ROW)(12)) Then
            'Exit Do
            'End If

            'If Len(sqlDT.Rows(ROW)(11).ToString) < 2 Then
            'Exit Do
            'End If
            'I() 'f sqlDT.Rows(ROW)(11)  = Nothing Then
            'Exit Do
            'End If

            '1	 2	    3	4	5	6	7	    8	    9	    10	    11	    12	13	14	15	16	17	    18	19
            'AJ1 AJ2	AJ3	AJ4	AJ5	AJI	FPA1	FPA2	FPA3	FPA4	ATIM	HME	EPO	AFM	KPE	DIE	XRVMA	EPA	POL
            '  Party_IDParty = IIf(IsDBNull(sqlDT.Rows(ROW)(14)), "", sqlDT.Rows(ROW)(14)) ' As String  '12344   ΚΩΔ ΣΥΝΑΛΛΑΣΟΜΕΝΟΥ
            AM_DcTp_Dscr = "Τιμολόγιο"
            Party_AFM = Trim(IIf(IsDBNull(sqlDT.Rows(ROW)(14)), "", sqlDT.Rows(ROW)(14)))  'Dim Party_AFM As String ' =""999349996
            If Len(Trim(Party_AFM)) <= 4 Then
                Party_AFM = "000000000"
            End If

            Party_ADDRESS = IIf(IsDBNull(sqlDT.Rows(ROW)(16)), "", sqlDT.Rows(ROW)(16))  'ToString  ' "ΠΟΛΥΣΤΗΛΟ ΚΑΒΑΛΑΣ"
            AM_DcTp_cd = "#ΤΥΠ-0"
            AMO_Srl_DSCR = "Πωλήσεις" '"ΠΩΛΗΣΕΙΣ"
            Base_dt = VB6.Format(sqlDT.Rows(ROW)(12), "YYYY-mm-dd")
            Base_INVOICE = sqlDT.Rows(ROW)(11)  ' =""#ΤΥΠ-0/Π000/1/Τιμολόγιο Παροχής Υπηρεσιών"
            Party_SNAME = sqlDT.Rows(ROW)(13)  '=""Θ. ΓΡΑΜΜΑΤΗΣ Κ.ΣΙΑ Ε.Ε""
            f_logPel = pel30.Text ' "30-00-00-0000"

            KAU_AJIA = nVal(sqlDT.Rows(ROW)(1)) + nVal(sqlDT.Rows(ROW)(2)) + nVal(sqlDT.Rows(ROW)(3)) + nVal(sqlDT.Rows(ROW)(4)) + nVal(sqlDT.Rows(ROW)(5)) + nVal(sqlDT.Rows(ROW)("AJ6")) + nVal(sqlDT.Rows(ROW)("AJ7"))
            FPA = nVal(sqlDT.Rows(ROW)(7)) + nVal(sqlDT.Rows(ROW)(8)) + nVal(sqlDT.Rows(ROW)(9)) + nVal(sqlDT.Rows(ROW)(10)) + nVal(sqlDT.Rows(ROW)("FPA6")) + nVal(sqlDT.Rows(ROW)("FPA7"))


            kau13 = nVal(sqlDT.Rows(ROW)(1))
            kau23 = nVal(sqlDT.Rows(ROW)(2))
            kau16 = nVal(sqlDT.Rows(ROW)(3))
            kau9 = nVal(sqlDT.Rows(ROW)(4))
            kau0 = nVal(sqlDT.Rows(ROW)(5))
            kau24 = nVal(sqlDT.Rows(ROW)("AJ6"))
            kau17 = nVal(sqlDT.Rows(ROW)("AJ7"))

            fpa13 = nVal(sqlDT.Rows(ROW)(7))
            fpa23 = nVal(sqlDT.Rows(ROW)(8))
            fpa16 = nVal(sqlDT.Rows(ROW)(9))
            fpa9 = nVal(sqlDT.Rows(ROW)(10))
            fpa24 = nVal(sqlDT.Rows(ROW)("FPA6"))
            fpa17 = nVal(sqlDT.Rows(ROW)("FPA7"))

            LOG13 = pol13.Text : LOG23 = pol23.Text
            LOG16 = POL16.Text : LOG9 = POL9.Text
            LOG0 = POL0.Text

            LOG24 = POL24.Text : LOG17 = POL17.Text


            FL_Ledg_Dscr = "ΠΩΛΗΣΕΙΣ ΧΟΝΔΡΙΚΗΣ ΕΣ. ΦΠΑ23%"
            FL_Ledg_cd = pol23.Text ' "70-00-00-0057"

            MVTP = "1" '6=πιστωτικα 2=αγορες  7=πιστωτικα αγορών
            System_sys = "SB" '      'SB =POLISEIS FR



            If InStr(fcLian, Mid(Base_INVOICE, 1, fnLian)) > 0 Then
                IsHand = "1" 'LTrim(Str(hand))
                cdRetailIdentity = arTam.Text
                LOG13 = Lian13.Text : LOG23 = Lian23.Text
                LOG24 = lian24.Text
                LOG0 = LIAN0.Text
                'If InStr("yπ", Mid(Base_INVOICE, 1, 1)) > 0 Then 'yphresies
                '    LOG23 = "73-4057"
                'End If
                'If InStr("πφ", Mid(Base_INVOICE, 1, 1)) > 0 Then 'yphresies
                '    cdRetailIdentity = "ΣΥ09002067"
                '    LOG23 = "73-4057"
                '    IsHand = "0" 'LTrim(Str(hand))
                'End If
                MVTP = "1" '6=πιστωτικα 2=αγορες  7=πιστωτικα αγορών


                '=======================================================================================
                'ElseIf InStr("GΓgΞμ", Mid(Base_INVOICE, 1, 1)) > 0 Then
            ElseIf InStr(fcTimAg, Mid(Base_INVOICE, 1, fnTimAg)) > 0 Then  ' τιμολογια αγορας
                'θελει ψαξιμο.................
                MVTP = 2
                IsHand = "" 'LTrim(Str(hand))  BP=AGORES
                cdRetailIdentity = ""
                LOG13 = ago13.Text : LOG23 = ago23.Text
                LOG16 = ago16.Text : LOG9 = ago9.Text
                LOG24 = ago24_6.Text
                LOG0 = ago0.Text ' "20-00-00-0000"


                'LOG0 = LIAN0.Text
                'System_sys = "BP"

                'System_Dscr_1
                AMO_Srl_DSCR = "Αγορές"
                AMO_Srl_DSCR = "ΑΓΟΡΕΣ (ΧΕΙΡΟΓΡΑΦΗ)"
                System_sys = "BP"





            ElseIf InStr(fcPistAg, Mid(Base_INVOICE, 1, fnPistAg)) > 0 Then  'πιστωτικά  τιμολογια αγορας
                '          ElseIf InStr("D", Mid(Base_INVOICE, 1, 1)) > 0 Then
                'θελει ψαξιμο.................
                MVTP = 7
                IsHand = "" 'LTrim(Str(hand))  BP=AGORES
                cdRetailIdentity = ""
                LOG13 = ago13.Text : LOG23 = ago23.Text
                LOG16 = ago16.Text : LOG9 = ago9.Text
                LOG0 = ago0.Text ' "20-00-00-0000"
                LOG24 = ago24_6.Text
                'LOG0 = LIAN0.Text
                System_sys = "BP"

                'System_Dscr_1
                AMO_Srl_DSCR = "Αγορές"
                AMO_Srl_DSCR = "ΑΓΟΡΕΣ (ΧΕΙΡΟΓΡΑΦΗ)"
                System_sys = "BP"

            ElseIf InStr(fcTimol, Mid(Base_INVOICE, 1, fnTimol)) > 0 Then 'τιμολογια -πιστωτικά
                cdRetailIdentity = ""
                IsHand = ""
                'LOG13 = pol13.Text : LOG23 = pol23.Text
                'LOG16 = POL16.Text : LOG9 = POL9.Text
                'LOG0 = POL0.Text
                MVTP = "1" '6=πιστωτικα 2=αγορες  7=πιστωτικα αγορών



                'ElseIf InStr("Y", Mid(Base_INVOICE, 1, 1)) > 0 Then 'yphresies
                '   LOG23 = "73-0057"
                '  MVTP = "1" '6=πιστωτικα 2=αγορες  7=πιστωτικα αγορών





                '   POL24.Text = node.Attributes("POL24").Value
                ' EPIS24.Text = node.Attributes("EPIS24").Value

            ElseIf InStr(fcPAR, Mid(Base_INVOICE, 1, fnPAR)) > 0 Then  'επιστροφη λιανικης
                LOG13 = "" : LOG24 = PAR24.Text
                MVTP = 6
                IsHand = "1" 'LTrim(Str(hand))
                System_sys = "SB" '      'SB =POLISEIS FR
            ElseIf InStr(fcPistLian, Mid(Base_INVOICE, 1, fnPistLian)) > 0 Then  'επιστροφη λιανικης
                LOG13 = episLian13.Text : LOG23 = episLian23.Text
                LOG24 = episLian24.Text
                MVTP = 6
                IsHand = "1" 'LTrim(Str(hand))
                'kau13 = -kau13
                'kau23 = -kau23
                'fpa13 = -fpa13
                'fpa23 = -fpa23
                'KAU_AJIA = -KAU_AJIA
                'FPA = -FPA
                System_sys = "SB" '      'SB =POLISEIS FR
                'System_sys = "FR"

                'End If
            ElseIf InStr(fcPistTim, Mid(Base_INVOICE, 1, fnPistTim)) > 0 Then  'pistvtiko timologio
                'kau13 = -kau13
                'kau23 = -kau23
                'fpa13 = -fpa13
                'fpa23 = -fpa23
                'KAU_AJIA = -KAU_AJIA
                'FPA = -FPA
                LOG13 = EPIS13.Text : LOG23 = EPIS23.Text
                MVTP = 6
                'End If
                System_sys = "SB" '      'SB =POLISEIS FR

            End If
            KAU_AJIA1 = KAU_AJIA
            FPA1 = FPA
            writeBAgor_row(writer)
            rowId = rowId + 11
            'Loop
            DoEvents()

            If Mid(LOG0, 1, 1) = "7" Or Mid(LOG13, 1, 1) = "7" Or Mid(LOG9, 1, 1) = "7" Or Mid(LOG17, 1, 1) = "7" Or Mid(LOG23, 1, 1) = "7" Or Mid(LOG24, 1, 1) = "7" Then

                If MVTP = "6" Then                                     '    POL= "1" Then '6=πιστωτικα 2=αγορες  7=πιστωτικα αγορών
                    esb23 = esb23 + kau23
                    esb24 = esb24 + kau24
                    esb17 = esb17 + kau17
                    esb13 = esb13 + kau13
                    esb0 = esb0 + kau0
                Else
                    sb24 = sb24 + kau24
                    sb23 = sb23 + kau23
                    sb13 = sb13 + kau13
                    sb16 = sb16 + kau16
                    sb9 = sb9 + kau9
                    sb0 = sb0 + kau0
                End If

            End If
            If Mid(LOG13, 1, 1) = "2" Or Mid(LOG23, 1, 1) = "2" Then
                bp23 = bp23 + kau23
                bp13 = bp13 + kau13
                bp0 = bp0 + kau0
            End If


            OK = 0

            For i = 1 To 30

                If Mid(Base_INVOICE, 1, 1) = Mid(parast(i), 1, 1) Then
                    OK = 1
                    ajia_ana_parast(i) = ajia_ana_parast(i) + kau13 + kau23 + kau16 + kau9 + kau0 + kau24 + kau17
                End If

            Next

            If OK = 0 Then
                nSynal = nSynal + 1
                parast(nSynal) = Mid(Base_INVOICE, 1, 1)
                ajia_ana_parast(nSynal) = kau13 + kau23 + kau16 + kau9 + kau0 + kau24 + kau17
            End If




            ExecuteSQLQuery("UPDATE TIM SET B_C1= '*'+convert(CHAR(10),GETDATE(),3) WHERE ID_NUM=" + Str(nVal(sqlDT.Rows(ROW)("ID_NUM"))), SQLDT2)






        Next







        writer.WriteEndDocument()
        writer.Close()

        ListBox1.Items.Clear()

        ListBox1.Items.Add("ΠΩΛ 13% " + VB6.Format(sb13, "0000000.00"))
        ListBox1.Items.Add("ΠΩΛ 23% " + VB6.Format(sb23, "0000000.00"))


        ListBox1.Items.Add("ΠΩΛ 17% " + VB6.Format(sb17, "0000000.00"))
        ListBox1.Items.Add("ΠΩΛ 24% " + VB6.Format(sb24, "0000000.00"))




        ListBox1.Items.Add("ΠΩΛ 16% " + VB6.Format(sb16, "0000000.00"))
        ListBox1.Items.Add("ΠΩΛ  9% " + VB6.Format(sb9, "0000000.00"))


        ListBox1.Items.Add("ΠΩΛ 0%  " + VB6.Format(sb0, "0000000.00"))
        ListBox1.Items.Add(" ")
        ListBox1.Items.Add("ΣΥΝΟΛΟ  " + VB6.Format(sb13 + sb23 + sb17 + sb24 + sb16 + sb9 + sb0, "0000000.00"))
        ListBox1.Items.Add(" ")
        ListBox1.Items.Add(" ")




        ListBox1.Items.Add("ΕΠ.ΠΩΛ 13% " + VB6.Format(esb13, "0000000.00"))
        ListBox1.Items.Add("ΕΠ.ΠΩΛ 23% " + VB6.Format(esb23, "0000000.00"))
        ListBox1.Items.Add("EΠ.ΠΩΛ 0%  " + VB6.Format(esb0, "0000000.00"))
        ListBox1.Items.Add(" ")
        ListBox1.Items.Add("ΣΥΝΟΛΟ  " + VB6.Format(esb13 + esb23 + esb0, "0000000.00"))
        ListBox1.Items.Add("----------------------------- ")

        ListBox1.Items.Add("ΣΥΝΟΛΟ  " + VB6.Format(sb13 + sb23 + sb16 + sb9 + sb0 + sb24 + sb17 - Esb17 - esb24 - esb13 - esb23 - esb0, "0000,000,000.00"))








        ListBox1.Items.Add("ΑΓΟΡ 13% " + VB6.Format(bp13, "########.00"))
        ListBox1.Items.Add("ΑΓΟΡ 23% " + VB6.Format(bp23, "########.00"))

        ListBox1.Items.Add("ΑΓΟΡ  0% " + VB6.Format(bp0, "########.00"))

        ListBox1.Items.Add(" ")
        ListBox1.Items.Add("ΣΥΝΟΛΟ  " + VB6.Format(bp13 + bp0 + bp23, "0000000.00"))



        For i = 1 To nSynal

            ' If Len(parast(i)) >= 1 Then
            ListBox1.Items.Add(parast(i) + " " + VB6.Format(ajia_ana_parast(i), "########.00"))
            ' End If

        Next



        MsgBox("Ενημερώθηκαν " + Str(ROW) + " εγγραφές. Δημιουργήθηκε το αρχείο export στο " + ff)
        'xlApp.Quit()


    End Sub


    Public Function checkServer() As Boolean
        Dim c As String
        Dim tmpStr As String
        c = "c:\mercvb\Config.ini"


        Dim par As String = ""
        Dim mf As String
        mf = c   ' "c:\mercvb\err3.txt"
        If Len(Dir(UCase(mf))) = 0 Then
            par = ":(local)\sql2012:sa:12345678:1:EMP"    '" 'G','g','Ξ','D'  "
            par = InputBox("ΒΑΣΗ ΔΕΔΟΜΕΝΩΝ", , par)
        Else
            FileOpen(1, mf, OpenMode.Input)
            '   Input(1, par)
            par = LineInput(1)
            FileClose(1)
        End If
        par = InputBox("ΒΑΣΗ ΔΕΔΟΜΕΝΩΝ", ":Π.Χ. (local)\sql2012:sa:12345678:1:EMP", par)

        'Input = InputBox("Text:")

        If String.IsNullOrEmpty(par) Then
            ' Cancelled, or empty
            checkServer = False
            ' MsgBox("εξοδος λογω μη σύνδεσης με βάση δεδομένων")
            Exit Function
        Else
            ' Normal
        End If


        FileOpen(1, mf, OpenMode.Output)
        PrintLine(1, par)
        FileClose(1)





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

                gconnect = "Provider=SQLOLEDB.1;;Password=" & Split(tmpStr, ":")(3) & _
                ";Persist Security Info=True ;" & _
                ";User Id=" & Split(tmpStr, ":")(2) & _
                ";Initial Catalog=" & Trim(Split(tmpStr, ":")(5)) & _
                ";Data Source=" & Split(tmpStr, ":")(1)




            Else
                'local
                'MsgBox(Split(tmpStr, ":")(1))
                gconnect = "Provider=SQLOLEDB;Server=" & Split(tmpStr, ":")(1) & _
                           ";Database=" & Split(tmpStr, ":")(5) & "; Trusted_Connection=yes;"

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
            sqlCon.ConnectionString = gconnect
            sqlCon.Open()
            checkServer = True
            sqlCon.Close()

            '            Dim GDB As New ADODB.Connection

        Catch ex As Exception
            checkServer = False
            MsgBox("εξοδος λογω μη σύνδεσης με βάση δεδομένων")
            End
        End Try
    End Function

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
            MsgBox("Error No. " & Err.Number & " Invalid database or no database found !! Adjust settings first", MsgBoxStyle.Critical, "Sales And Inventory")
            MsgBox(SQLQuery)
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
        End Try
        'Return sqlDT
    End Sub

    Public Shared Sub DoEvents()

    End Sub

    Function Get_AJ_ASCII(ByRef pol As String, _
                          ByVal polepis As String, _
                          ByVal ago As String, _
                          ByVal AGOEPIS As String) As Boolean

        '<EhHeader>


        '</EhHeader>



        Dim R As New ADODB.Recordset
        Dim x As String

        'If gConnect = "Access" Then
        '   Set db = OpenDatabase(gDir, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If

        ExecuteSQLQuery("select POL,EIDOS,AJIA_APOU from PARASTAT", SQLDT2)

        pol = " "

        Dim row As Integer
        For row = 0 To SQLDT2.Rows.Count - 1

            If IsDBNull(SQLDT2.Rows(row)("eidos")) Or IsDBNull(SQLDT2.Rows(row)("pol")) Or IsDBNull(SQLDT2.Rows(row)("ajia_apou")) Then

            Else

                If SQLDT2.Rows(row)("pol") = "1" And SQLDT2.Rows(row)("ajia_apou") = "3" Then
                    pol = pol + "'" + SQLDT2.Rows(row)("eidos") + "',"
                End If

                If SQLDT2.Rows(row)("pol") = "1" And SQLDT2.Rows(row)("ajia_apou") = "4" Then
                    polepis = polepis + "'" + SQLDT2.Rows(row)("eidos") + "',"
                End If

                If SQLDT2.Rows(row)("pol") = "2" And SQLDT2.Rows(row)("ajia_apou") = "1" Then
                    ago = ago + "'" + SQLDT2.Rows(row)("eidos") + "',"
                End If

                If SQLDT2.Rows(row)("pol") = "2" And SQLDT2.Rows(row)("ajia_apou") = "2" Then
                    AGOEPIS = AGOEPIS + "'" + SQLDT2.Rows(row)("eidos") + "',"
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



End Class

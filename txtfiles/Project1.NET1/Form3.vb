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

Imports System
' Imports System.Xml
Imports System.Xml.Schema
Imports System.Xml.XPath

Public Class FORM3
    Inherits System.Windows.Forms.Form

    Public dataBytes() As Byte

    Public sqlDT As New DataTable
    Public sqlDaTaSet As New DataSet
    Public sqlDTx As New DataTable
    Public openedFileStream As System.IO.Stream

    Public gSplitter As String = ";"

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


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'pvlhsevn
        '==========================================================================================
        On Error Resume Next

        'DBGrid1.Clear
        If checkServer() Then
        Else
            Exit Sub
        End If
        ' gconnect=":HP530\SQL2012:sa:12345678:1:perp"
        Dim M_AFM As String : M_AFM = Text1.Text
        Dim m_mhnas As String : m_mhnas = Text2.Text
        Dim m_etos As String : m_etos = Text3.Text
        Dim hmer As String = Text4.Text  '"31/03/2015"
        hmer = VB6.Format(hmer, "yyyy-mm-dd")

        If Len(M_AFM) <> 9 Then
            ' MsgBox("λαθος στο ΑΦΜ")
            'Exit Sub
        End If

        If Len(Dir("C:\SYGK", vbDirectory)) = 0 Then
            MkDir("C:\SYGK")
        End If

        Dim file

        Dim F_CASH

        F_CASH = arTam.Text ' "ΣΥ09002067"
        file = "C:\SYGK\XML.TXT"
        Kill(file)

        Dim R As New ADODB.Recordset

        'Open "C:\SYGK\synola.txt" For Output As #5
        Dim m_filename As String

        m_filename = "C:\SYGK\" + Text1.Text + "_" + Text3.Text + Text2.Text + ".XML"

        Dim writer As New XmlTextWriter(m_filename, System.Text.Encoding.UTF8)
        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2
        ' Create a Continent element and set its value to
        ' that of the New Continent dialog box
        'writer.WriteAttributeString("Table", , "sadasasd sdsd")

        '<Data Name="SX" Style="BRowse"><row name="Hd" mode="3" rowId="7">
        writer.WriteStartElement("packages")
        writer.WriteStartElement("package")
        writer.WriteAttributeString("actor_afm", M_AFM)
        writer.WriteAttributeString("month", m_mhnas)
        writer.WriteAttributeString("year", m_etos)

        writer.WriteStartElement("groupedRevenues")
        writer.WriteAttributeString("action", "replace")





        '   Open m_filename For Output As #1
        '
        '  Print #1, "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""YES""?>"
        ' Print #1, "<packages>"
        '   Print #1, " <package actor_afm=""" + M_AFM + """ month=""" + m_mhnas + """ year=""" + m_etos + """>"
        '  Print #1, " <groupedRevenues action=""replace"">"


        ExecuteSQLQuery("SELECT * FROM SYNOLAKEPYO WHERE POLHS=1  ")

        Dim sxre As Single, spis As Single


        sxre = 0 : spis = 0
        Dim sxretax As Single = 0
        Dim spistax As Single = 0
        Dim sfpa(10) As Single


        '   DBGrid1.row = 0 : DBGrid1.Col = 1
        '  DBGrid1.Text = "Καθ.Αξία"

        '  DBGrid1.row = 0 : DBGrid1.Col = 2
        '  DBGrid1.Text = "Φ.Π.Α."

        ListBox1.Items.Clear()



        Dim a
        Dim k As Long

        FileOpen(1, "C:\SYGK\ERR.TXT", OpenMode.Output)




        For k = 0 To sqlDT.Rows.Count - 1

            writer.WriteStartElement("revenue")
            writer.WriteStartElement("afm") : writer.WriteString(sqlDT.Rows(k)("AFM").ToString.Trim) : writer.WriteEndElement()  'AFM
            writer.WriteStartElement("amount") : writer.WriteString(VB6.Format(sqlDT.Rows(k)("ajia"), "#######0.00")) : writer.WriteEndElement()  'AJIA
            writer.WriteStartElement("tax") : writer.WriteString(VB6.Format(sqlDT.Rows(k)("fpa"), "#######0.00")) : writer.WriteEndElement()  'FPA
            writer.WriteStartElement("invoices") : writer.WriteString(VB6.Format(sqlDT.Rows(k)("tem"), "########0")) : writer.WriteEndElement()  'NtIM

            writer.WriteStartElement("note") : writer.WriteString(sqlDT.Rows(k)("pis")) : writer.WriteEndElement()  'CREDIT
            writer.WriteStartElement("date") : writer.WriteString(VB6.Format(sqlDT.Rows(k)("shme"), "yyyy-mm-dd")) : writer.WriteEndElement()  'DATE
            writer.WriteEndElement()  'REVENUE



            a = check_afm(sqlDT.Rows(k)("AFM").ToString)
            If a = 0 Then
                PrintLine(1, sqlDT.Rows(k)("AFM").ToString)
            End If


            If sqlDT.Rows(k)("pis") = "normal" Then
                sxre = sxre + sqlDT.Rows(k)("ajia")
                sxretax = sxretax + sqlDT.Rows(k)("fpa")
                sfpa(1) = sfpa(1) + sqlDT.Rows(k)("fpa")
            Else
                spis = spis + sqlDT.Rows(k)("ajia")
                spistax = spistax + sqlDT.Rows(k)("fpa")

                sfpa(2) = sfpa(2) + sqlDT.Rows(k)("fpa")
            End If


            ' R.MoveNext()
        Next 'Loop

        FileClose(1)



        writer.WriteEndElement()  'groupedREVENUE


        '        DBGrid1.row = 1 : DBGrid1.Col = 0
        ListBox1.Items.Add("Σύνολο Χρ.τιμ. ")
        '      DBGrid1.Col = 1
        ListBox1.Items.Add(VB6.Format(sxre, "###,###.00"))


        'DBGrid1.Col = 2
        ListBox1.Items.Add(VB6.Format(sfpa(1), "###,###.00"))


        'SFPA(1) = SFPA(1) + R!FPA

        On Error GoTo 0


        'DBGrid1.row = 2 : DBGrid1.Col = 0
        ListBox1.Items.Add("Σύνολο πισ.τιμ. ")
        ' DBGrid1.Col = 1
        ListBox1.Items.Add(VB6.Format(-spis, "####,###.00"))

        ' DBGrid1.Col = 2
        ListBox1.Items.Add(VB6.Format(-sfpa(2), "###,###.00"))



        'Print #5, "Σύνολο Χρ.τιμ. " + Format(sxre, "########0.00")
        'Print #5, " Πιστωτικά " + Format(spis, "########0.00")

        ' writer.WriteEndElement()  'GROUPEDREVENUES





        'action=""replace"">"

        ExecuteSQLQuery("SELECT SUM(AJIA) AS SAJIA,SUM(FPA) AS SFPA FROM SYNOLAKEPYO WHERE POLHS=3")


        Dim slian As Single = 0

        Dim slianTax As Single = 0

        Dim EXEILIANIKES As Integer = 0

        If (Not IsDBNull(sqlDT.Rows(0)("sajia"))) Then
            If sqlDT.Rows(0)("sajia") > 0 Then

                writer.WriteStartElement("groupedCashRegisters")
                writer.WriteAttributeString("action", "replace")

                EXEILIANIKES = 1

                For k = 0 To sqlDT.Rows.Count - 1
                    If (Not IsDBNull(sqlDT.Rows(0)("sajia"))) Then
                        writer.WriteStartElement("cashregister")

                        writer.WriteStartElement("amount") : writer.WriteString(VB6.Format(sqlDT.Rows(k)("sajia"), "#######0.00")) : writer.WriteEndElement()  'AJIA
                        writer.WriteStartElement("tax") : writer.WriteString(VB6.Format(sqlDT.Rows(k)("sfpa"), "#######0.00")) : writer.WriteEndElement()  'FPA
                        '  writer.WriteStartElement("invoices") : writer.WriteString(VB6.Format(sqlDT.Rows(k)(""), "########0")) : writer.WriteEndElement()  'NtIM
                        writer.WriteStartElement("date") : writer.WriteString(hmer) : writer.WriteEndElement()  'DATE
                        writer.WriteEndElement()  'cashregister
                        slian = slian + sqlDT.Rows(k)("sajia")
                        slianTax = slianTax + sqlDT.Rows(k)("sfpa")
                    End If
                Next


            End If
        End If




        '  writer.WriteEndElement()  'groupedCashRegisters


        '  If Not R.EOF Then
        'Print #1, "  <cashregister>"
        'Print #1, "      <amount>" + Replace(Format(R!SAJIA, "######.00"), ".", ",") + "</amount>"
        'Print #1, "      <tax>" + Replace(Format(R!sfpa, "######.00"), ".", ",") + "</tax>"
        'Print #1, "      <date>" + Replace(Format(d2, "YYYY/MM/DD"), "/", "-") + "</date>"
        'Print #1, "  </cashregister>"
        '      slian = R!SAJIA
        '  Else
        '      slian = 0

        '  End If


        'Print #5, "Λιαν" + Format(slian, "######,##0.00")
        '    DBGrid1.row = 3 : DBGrid1.Col = 0
        ListBox1.Items.Add("Σύνολο λιανικών")
        ' DBGrid1.Col = 1
        ListBox1.Items.Add(VB6.Format(slian, "####,##0.00"))


        '  sfpa(3) = R!sfpa
        '   DBGrid1.Col = 2
        'ListBox1.Items.Add(VB6.Format(sqlDT.Rows(0)("sfpa"), "####,##0.00"))





        ' R.Close()

        ExecuteSQLQuery("SELECT SUM(AJIA) AS SAJIA,SUM(FPA) AS SFPA FROM SYNOLAKEPYO WHERE POLHS=4")


        Dim sTam As Single = 0
        Dim sTamTax As Single = 0
        If (Not IsDBNull(sqlDT.Rows(0)("sajia"))) Then
            If sqlDT.Rows(0)("sajia") > 0 Then


                ' GRAFO TO BEGGINNING
                If EXEILIANIKES = 0 Then
                    writer.WriteStartElement("groupedCashRegisters")
                    writer.WriteAttributeString("action", "replace")
                    EXEILIANIKES = 1
                End If




                For k = 0 To sqlDT.Rows.Count - 1
                    writer.WriteStartElement("cashregister")
                    writer.WriteStartElement("cashreg_id") : writer.WriteString(F_CASH) : writer.WriteEndElement()  'id
                    writer.WriteStartElement("amount") : writer.WriteString(VB6.Format(sqlDT.Rows(k)("sajia"), "########.00")) : writer.WriteEndElement()  'AJIA
                    writer.WriteStartElement("tax") : writer.WriteString(VB6.Format(sqlDT.Rows(k)("sfpa"), "########.00")) : writer.WriteEndElement()  'FPA
                    'writer.WriteStartElement("invoices") : writer.WriteString(VB6.Format(sqlDT.Rows(k)("tem"), "########0")) : writer.WriteEndElement()  'NtIM
                    writer.WriteStartElement("date") : writer.WriteString(hmer) : writer.WriteEndElement()  'DATE
                    writer.WriteEndElement()  'CashRegister
                    sTam = sTam + sqlDT.Rows(k)("sajia")
                    sTamTax = sTamTax + sqlDT.Rows(k)("sfpa")
                Next
            End If
        End If


        ' GRAFO TO BEGGINNING
        If EXEILIANIKES = 1 Then
            writer.WriteEndElement()  'groupedCashRegisters
        End If








        '    If Not R.EOF Then
        'Print #1, "  <cashregister>"
        'Print #1, "  <cashreg_id>" + f_tam + "</cashreg_id>"
        'Print #1, "      <amount>" + Replace(Format(R!SAJIA, "######.00"), ".", ",") + "</amount>"
        'Print #1, "      <tax>" + Replace(Format(R!sfpa, "######.00"), ".", ",") + "</tax>"
        'Print #1, "      <date>" + Format(d2, "YYYY-MM-DD") + "</date>"
        'Print #1, "  </cashregister>"
        'Print #1, "</groupedCashRegisters>"
        '        On Error Resume Next
        '        sTam = R!SAJIA
        '    Else
        '        sTam = 0
        '    End If


        'Print #5, "Λιαν" + Format(sTam, "####,##0.00")
        'DBGrid1.row = 4 : DBGrid1.Col = 0
        ListBox1.Items.Add("Σύνολο ταμειακών")
        'DBGrid1.Col = 1
        ListBox1.Items.Add(VB6.Format(sTam, "####,##0.00"))
        'Print #5, Format(sTam, "####,##00.00")
        If IsDBNull(sqlDT.Rows(0)("sfpa")) Then
            ListBox1.Items.Add("0")
        Else
            ListBox1.Items.Add(VB6.Format(sqlDT.Rows(0)("sfpa"), "##########.00"))
        End If

        ListBox1.Items.Add("------------------------")



        ListBox1.Items.Add(VB6.Format(sTam + slian + sxre - spis, "##########.00"))
        ListBox1.Items.Add(VB6.Format(sTamTax + slianTax + sxretax - spistax, "##########.00"))

        'DBGrid1.row = 4 : DBGrid1.Col = 2
        'DBGrid1.Text = Format(R!sfpa, "####,##0.00")

        'Print #5, Format(R!sfpa, "####,##00.00")

        ListBox2.Width = ListBox2.Width * 2





        'DBGrid1.row = 5
        'DBGrid1.Text = "Σύνολο  "
        'DBGrid1.Col = 1
        'DBGrid1.Text = Format(sxre - spis + sTam + slian, "###,##0.00")


        'DBGrid1.Col = 2
        'DBGrid1.Text = Format(sfpa(1) - sfpa(2) + sfpa(3) + sfpa(4), "###,##0.00")






        writer.WriteEndElement()  'PACKAGE
        writer.WriteEndElement()  'PACKAGES
        writer.WriteEndDocument()
        writer.Close()


        '  Print #1, "</package>"
        ' Print #1, "</packages>"

        ' Close #1

        '   Dim k As Integer

        '    For k = 1 To 5
        ' Print #5, Left(DBGrid1.TextMatrix(k, 0) + Space(30), 30) + Right(Space(30) + DBGrid1.TextMatrix(k, 1), 30) + Right(Space(30) + DBGrid1.TextMatrix(k, 2), 30)

        '    Next




        ExecuteSQLQuery("SELECT SUM(AJIA) AS SAJIA,SUM(FPA) AS SFPA FROM SYNOLAKEPYO WHERE POLHS=9")

        If (Not IsDBNull(sqlDT.Rows(0)("sajia"))) Then
            ListBox1.Items.Add("ΕΞΩΤΕΡΙΚΟΥ")
            ListBox1.Items.Add(VB6.Format(sqlDT.Rows(0)("sajia"), "##########.00"))

        End If



        'Close #5

        MsgBox("ΑΠΟΘΗΚΕΥΤΗΚΕ ΤΟ " + m_filename + Chr(3) + " Kαι c:\sygk\synola.txt το αρχείο με τα σύνολα")
        ' End Sub

        ListBox2.Items.Clear()



        'Public Class Form1

        ' Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim myDocument As New XmlDocument
        myDocument.Load(m_filename) ' m_filename)  ' "C:\somefile.xml"
        myDocument.Schemas.Add("", "c:\sygk\gsis_packages_schema.xsd") 'namespace here or empty string
        Dim eventHandler As ValidationEventHandler = New ValidationEventHandler(AddressOf ValidationEventHandler)
        myDocument.Validate(eventHandler)
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












    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '  Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim writer As New XmlTextWriter("c:\mercvb\product2.xml", System.Text.Encoding.UTF8)
        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2
        ' Create a Continent element and set its value to
        ' that of the New Continent dialog box
        'writer.WriteAttributeString("Table", , "sadasasd sdsd")

        '<Data Name="SX" Style="BRowse"><row name="Hd" mode="3" rowId="7">
        writer.WriteStartElement("packages")
        writer.WriteStartElement("package")
        writer.WriteAttributeString("actor_afm", "SX")
        writer.WriteAttributeString("month", "9")
        writer.WriteAttributeString("year", "2014")


        writer.WriteStartElement("groupedRevenues")
        writer.WriteAttributeString("action", "replace")

        writer.WriteStartElement("revenue")

        writer.WriteStartElement("afm") : writer.WriteString("028783755") : writer.WriteEndElement()  'AFM
        writer.WriteStartElement("amount") : writer.WriteString("028783755") : writer.WriteEndElement()  'AFM
        writer.WriteStartElement("tax") : writer.WriteString("028783755") : writer.WriteEndElement()  'AFM
        writer.WriteStartElement("invoices") : writer.WriteString("028783755") : writer.WriteEndElement()  'AFM

        writer.WriteStartElement("note") : writer.WriteString("028783755") : writer.WriteEndElement()  'AFM
        writer.WriteStartElement("date") : writer.WriteString("028783755") : writer.WriteEndElement()  'AFM





        writer.WriteEndElement()  'REVENUE

        writer.WriteEndElement()  'GROUPEDREVENUES




        writer.WriteEndElement()  'PACKAGE
        writer.WriteEndElement()  'PACKAGES

    End Sub



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
        If Len(Trim(par)) > 5 Then
            FileOpen(1, mf, OpenMode.Output)
            PrintLine(1, par)
            FileClose(1)

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

                gConnect = "Provider=SQLOLEDB.1;;Password=" & Split(tmpStr, ":")(3) & _
                ";Persist Security Info=True ;" & _
                ";User Id=" & Split(tmpStr, ":")(2) & _
                ";Initial Catalog=" & Trim(Split(tmpStr, ":")(5)) & _
                ";Data Source=" & Split(tmpStr, ":")(1)




            Else
                'local
                'MsgBox(Split(tmpStr, ":")(1))
                gConnect = "Provider=SQLOLEDB;Server=" & Split(tmpStr, ":")(1) & _
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
            sqlCon.ConnectionString = gConnect
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


    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click

    End Sub

    Private Sub Text3_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text3.Leave
        Text4.Text = "30/" + Text2.Text + "/" + Text3.Text
    End Sub

    Private Sub Text3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Text3.TextChanged

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EXCELKINHSEON.Click
        'ΑΡΧΕΙΟ ΚΙΝΗΣΕΩΝ
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

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        'ΑΡΧΕΙΟ ΠΕΛΑΤΩΝ
        If Len(Trim(TextBox2.Text)) = 0 Then
            CD1.ShowDialog()
            TextBox2.Text = CD1.FileName
        Else
            If Len(Dir(LTrim(TextBox2.Text), FileAttribute.Normal)) < 2 Then
                MsgBox("δεν υπάρχει το αρχείο " & TextBox1.Text)
                Exit Sub
            End If
        End If
    End Sub


    Private Sub AMBROSIADIS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AMBROSIADIS.Click
        'αμβροσιαδης

        'ΣΤΟ ΦΥΛΛΟ 2 ΕΧΩ ΤΟΥΣ ΠΕΛΑΤΕΣ ΜΕ ΑΦΜ ΚΑΙ ΣΤΟ ΦΥΛΛΟ1 ΤΑ ΤΙΜΟΛΟΓΙΑ ΜΕ ΤΑ ΠΟΣΑ
        'μεταφέρει το ΑΦΜ ΣΤΟ ΦΥΛΛΟ1(στηλη 14)  ΑΠΟ ΤΟ ΦΥΛΛΟ2

        ' pel(ROW, 2)  πινακας που φορτώνει ολους τους πελατες απο το φυλλο 2
        ' 
        Dim nHME As Integer = Val(cHME.Text)
        Dim nPAR As Integer = Val(cPAR.Text)

        Dim nKOD As Integer = Val(cKOD.Text)
        Dim nEPO As Integer = Val(cEPO.Text)

        Dim nKAU24 As Integer = Val(cKAU24.Text)
        Dim nKAU13 As Integer = Val(cKAU13.Text)
        Dim nKAU0 As Integer = Val(cKAU0.Text)

        Dim nFPA24 As Integer = Val(cFPA24.Text)
        Dim nFPA13 As Integer = Val(cFPA13.Text)




        If nHME = 0 Then
            MsgBox("φορτωστε τον πίνακα")
            Exit Sub

        End If



        ',nPAR,nKAU24,nKAU13,






        ' Label8.Text = "1-ΗΜΕΡ 2-ΠΑΡ 5-23%ΚΑΘ 7-13%ΚΑΘ 12-23%ΦΠΑ 10-0%ΚΑΘ 14-ΦΠΑ13% 36-συν.αξια "
        Dim debug As Boolean = False

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook

        Dim xlAppPel As Excel.Application
        Dim xlWorkBookPel As Excel.Workbook


        Dim xl As Excel.Worksheet
        Dim xlPEL As Excel.Worksheet


        Dim xlok As Excel.Worksheet

        xlApp = New Excel.ApplicationClass


        xlWorkBook = xlApp.Workbooks.Open(TextBox1.Text)

        xlWorkBook.Worksheets.Add()
        xl = xlWorkBook.Worksheets(2) ' .Add

        xlok = xlWorkBook.Worksheets(1)




        xlAppPel = New Excel.ApplicationClass
        xlWorkBookPel = xlApp.Workbooks.Open(TextBox2.Text)
        xlPEL = xlWorkBookPel.Worksheets(1)

        'metafora me σωστη γραμμογραφηση στο ΝΕΟ ΦΥΛΛΟ ΠΟΥ ΔΗΜΙΟΥΡΓΕΙ ΣΤΗΝ ΑΡΧΗ  3
        '=========================================
        '===============================================================================real onomatepvmymo 54100
        Dim nRows As Long  'ποσα τιμολογια εχει

        ROW = 0
        Do While True
            ROW = ROW + 1
            If debug Then
                If ROW > 100 Then Exit Do
            End If



            If xl.Cells(ROW, 1).value = Nothing Then
                Exit Do
            End If
            xlok.Cells(ROW, 1) = xl.Cells(ROW, nKAU13) '13% kauarh

            xlok.Cells(ROW, 2) = xl.Cells(ROW, nKAU24) ' 23%

            xlok.Cells(ROW, 5) = xl.Cells(ROW, nKAU0) '0%

            ' xlok.Cells(ROW, 6) = xl.Cells(ROW, 36) 'συνολικη αξια

            'fpa
            xlok.Cells(ROW, 7) = xl.Cells(ROW, nFPA13).value  'fpa 13
            xlok.Cells(ROW, 8) = xl.Cells(ROW, nFPA24).value  '23%




            'xlok.Cells(ROW, 7) = xl.Cells(ROW, 12).value - xl.Cells(ROW, 8).value  '13%
            '11 apa   12 hme   13 epo  14 afm



            xlok.Cells(ROW, 11) = xl.Cells(ROW, nPAR)   'apa
            xlok.Cells(ROW, 12) = xl.Cells(ROW, nHME).value  'hmeromhnia

            xlok.Cells(ROW, 13) = xl.Cells(ROW, nEPO)  'epvnymia

            xlok.Cells(ROW, 14) = xl.Cells(ROW, nKOD)   'afm





            Me.Text = ROW



        Loop

        nRows = ROW




        'MsgBox("ok")

        'xlWorkBook.Save()
        'xlApp.Quit()

        'Exit Sub


        '==========================================

        Dim pel(2000, 2) As String

        ROW = 1

        Dim hand As Integer = 0
        Dim cc As String

        '========================φορτωνω το αρχειο των πελατων στον πινακα pe() =======================================================real onomatepvmymo 54100
        'APO XLPEL 
        'Κωδικός	Α.Φ.Μ. 	Επωνυμία

        ' pel(ROW, 0) =   KODIKOS
        ' ' ΟΝΟΜΑ  pel(ROW, 1)
        ' AFM   PEL(ROW,2)


        Do While True
            ROW = ROW + 1
            If debug Then
                'If ROW > 100 Then Exit Do
            End If



            If xlPEL.Cells(ROW, 1).value = Nothing Then
                Exit Do
            End If
            'PEL(K, 2) + ";" + PEL(K, 0)  ' afm;kodikos
            pel(ROW, 0) = xlPEL.Cells(ROW, 1).value.ToString  ' ΚΩΔΙΚΟΣ   pel(ROW, 0)

            ' ΟΝΟΜΑ  pel(ROW, 1)

            If xlPEL.Cells(ROW, 3).value = Nothing Then
                pel(ROW, 1) = ""
            Else
                pel(ROW, 1) = xlPEL.Cells(ROW, 3).value.ToString
            End If


            'AFM
            If IsDBNull(xlPEL.Cells(ROW, 2).value) Then
                pel(ROW, 2) = ""
            Else
                If xlPEL.Cells(ROW, 2).value = Nothing Then
                    pel(ROW, 2) = ""
                Else
                    cc = xlPEL.Cells(ROW, 2).value.ToString
                    pel(ROW, 2) = "'" + cc
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


        '=========================================================real onomatepvmymo 54100
        ROW = 1
        Do While True
            ROW = ROW + 1

            If debug Then
                If ROW > 100 Then Exit Do
            End If

            If ROW >= nRows And xl.Cells(ROW, 13).value = Nothing Then
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


            If xlok.Cells(ROW, 14).VALUE = Nothing Then
                C = ""
            Else
                C = xlok.Cells(ROW, 14).VALUE.ToString
            End If





            cc = SCAN_PEL_SIMPLE(C, pel)
            'xlok.Cells(ROW, 14) = Split(cc, ";")(2)  'afm
            ' 
            If InStr(cc, ";") > 0 Then
                xlok.Cells(ROW, 14) = Split(cc, ";")(2)  'afm
                xlok.Cells(ROW, 15) = Split(cc, ";")(1)  'kodikos
            End If



            Me.Text = Str(ROW) + xl.Cells(ROW, 14).ToString

        Loop


















        '    xlWorkBook.Save()
        '        xlApp.Quit()

        xlWorkBook.Save()
        xlApp.Quit()


        xlAppPel.Quit()

        MsgBox("ok")


    End Sub


    Function SCAN_PEL_SIMPLE(ByVal X As String, ByRef PEL(,) As String) As String
        ' AYTO CAXNEI ΚΩΔΙΚΟΣ SKETO 

        Dim K As Integer
        SCAN_PEL_SIMPLE = ""
        Dim c As String
        Dim L As Integer
        If InStrRev(X, "-") - InStr(X, "-") > 0 And InStr(X, "-") < 5 Then   'an exei >= 2 payles me ayton ton tropo



        Else

            c = X  ' Split(X, "-")(0)
            For K = 1 To 2000
                If c = PEL(K, 0) Then
                    SCAN_PEL_SIMPLE = PEL(K, 1) + ";" + PEL(K, 0) + ";" + PEL(K, 2)
                    Exit For
                End If

            Next
        End If

    End Function



    Function SCAN_PEL(ByVal X As String, ByRef PEL(,) As String) As String
        ' AYTO CAXNEI ΚΩΔΙΚΟΣ-ΟΝΟΜΑ  ΟΤΑ ΗΤΑΝ ΜΑΖΙ 

        Dim K As Integer
        SCAN_PEL = ""
        Dim c As String
        Dim L As Integer
        If InStrRev(X, "-") - InStr(X, "-") > 0 And InStr(X, "-") < 5 Then   'an exei >= 2 payles me ayton ton tropo
            For K = 1 To 2000
                L = Len(PEL(K, 1)) 'ΜΗΚΟΣ ΚΩΔΙΚΟΥ
                If L > 1 Then
                    If Mid(X, 1, L) = PEL(K, 1) Then
                        SCAN_PEL = PEL(K, 2) + ";" + PEL(K, 0)  ' afm;kodikos
                        Exit For
                    End If
                End If

            Next
        Else

            c = Split(X, "-")(0)
            For K = 1 To 2000
                If c = PEL(K, 1) Then
                    SCAN_PEL = PEL(K, 2) + ";" + PEL(K, 0)
                    Exit For
                End If

            Next
        End If

    End Function

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '<EhHeader>
        ' On Error GoTo create2014_Click_Err

        If checkServer() Then
        Else
            Exit Sub
        End If



        '</EhHeader>
        Dim R As New ADODB.Recordset

        Dim RH As New ADODB.Recordset

        Dim rPEL As New ADODB.Recordset

        Dim SQLPEL As String

        Dim k As Integer

        Dim PolTam As String

        PolTam = "''"

        Dim PolTim, PolPis, PolTamMhx, PolTamMhxPIS, PolTamTam, AgoTim, AgoPis, AGOEXO As String

        gdb.Open(gConnect)

        PolTim = "''" : PolPis = "''" : PolTamMhx = "''" : PolTamTam = "''" : AgoTim = "''" : AgoPis = "''" : AGOEXO = "''"
        PolTamMhxPIS = "''"
        R.Open("seLect * FROM PARASTAT ", gdb, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        R.MoveFirst()
        k = 0

        'Dim M_TAM As String

        Dim f_tam As String = ""

        Do While Not R.EOF

            If (Not IsDBNull(R("myf").Value)) And R.Fields("pol").Value.ToString = "1" Then  ' poliseis

                Select Case R("myf").Value
                    'Select Case R.Fields("myf").Value
                    Case 1
                        PolTim = PolTim + ",ASCII('" + R("eidos").Value.ToString + "')"

                    Case 2
                        PolPis = PolPis + ",ASCII('" + R("eidos").Value.ToString + "')"

                    Case 3
                        PolTam = PolTam + ",ASCII('" + R("eidos").Value.ToString + "')"

                    Case 4
                        PolTamMhx = PolTamMhx + ",ASCII('" + R("eidos").Value.ToString + "')"
                        f_tam = Trim(R("TAMEIAKI").Value.ToString)

                    Case 7
                        PolTamMhxPIS = PolTamMhxPIS + ",ASCII('" + R("eidos").Value.ToString + "')"




                End Select






            End If

            If R.Fields("pol").Value.ToString = "2" Then  ' agores    'R("pol").ToString = "2"
                If (Not IsDBNull(R("myf").Value)) Then

                    Select Case R("myf").Value


                        Case 1
                            AgoTim = AgoTim + ",ASCII('" + R("eidos").Value.ToString + "')"

                        Case 2
                            AgoPis = AgoPis + ",ASCII('" + R("eidos").Value.ToString + "')"

                        Case 5
                            AGOEXO = AGOEXO + ",ASCII('" + R("eidos").Value.ToString + "')"
                    End Select
                End If


            End If

            R.MoveNext()
        Loop

        R.Close()

        ListBox2.Items.Add("POLTIM       " + PolTim)
        ListBox2.Items.Add("POLPIS       " + PolPis)
        ListBox2.Items.Add("POLTAM       " + PolTam)
        ListBox2.Items.Add("POLTAMMHX    " + PolTamMhx)
        ListBox2.Items.Add("POLTAMMHXPIS " + PolTamMhxPIS)








        '     PolTim = apo2kaimeta(PolTim)
        '     PolPis = apo2kaimeta(PolPis)
        '     PolTamMhx = apo2kaimeta(PolTamMhx)
        '     PolTam = apo2kaimeta(PolTam)
        '     AgoTim = apo2kaimeta(AgoTim)
        '     AgoPis = apo2kaimeta(AgoPis)

        gdb.Execute("update TIM SET KR2=ASCII(LEFT(ATIM,1))")




        Dim ALL As String

        ALL = PolTim + "," + PolPis + "," + PolTamMhx + "," + PolTam + "," + AgoTim + "," + AgoPis + "," + AGOEXO + "," + PolTamMhxPIS

        Dim sql As String, pol As String, ag As String

180:    sql = "select 1 AS POLHS,'normal' AS PIS,'          ' as TAM, "
        sql = sql + "(TIM.AJ1+TIM.AJ2+TIM.AJ3+TIM.AJ4+TIM.AJ5+TIM.AJ6+TIM.AJ7) as AJIA,KR2 AS TYPOS,"
        sql = sql + "(TIM.FPA1+TIM.FPA2+TIM.FPA3+TIM.FPA4+TIM.FPA6+TIM.FPA7) as FPA," & "1 AS AEG,PEL.AFM,PEL.EPO,PEL.DOY,"
        sql = sql + "PEL.EPA,PEL.DIE,PEL.POL,PLAISIO AS TK,TIM.EIDOS,HME " & " INTO TIMKEPYO "

        sql = sql + "FROM TIM  INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD  " & " "
        sql = sql + "WHERE TIM.AJ1+TIM.AJ2+TIM.AJ3+TIM.AJ4+TIM.AJ5>=0 AND KR2 IN (" + ALL + ")"
190:    sql = sql + " and HME>='" + Format(D1, "MM/DD/YYYY") + "' AND HME<='" + Format(d2, "MM/DD/YYYY") + "'"

        On Error Resume Next

200:    gdb.Execute("DROP TABLE TIMKEPYO")
210:    gdb.Execute("DROP TABLE TIMKEPYO2")
212:    gdb.Execute("DROP TABLE  SYNOLAKEPYO")

220:    ' Gdb.Execute "DROP TABLE PEL22"

        Dim nn As Integer, NNAG As Integer

230:    gdb.Execute(sql, nn)

        'αφαιρω τις λιανικές
        'gdb.Execute("UPDATE TIMKEPYO SET FPA=abs(FPA),AJIA=abs(AJIA) WHERE TYPOS=ascii('p')")


        'αφαιρω τις λιανικές
        gdb.Execute("UPDATE TIMKEPYO SET FPA=-ABS(FPA),AJIA=-abs(AJIA) WHERE TYPOS IN (" + PolTamMhxPIS + ")")



        gdb.Execute("UPDATE TIMKEPYO SET POLHS=1 WHERE TYPOS IN (" + PolTim + ")")







        R.Open("select count(*) from TIMKEPYO  WHERE AFM='000000000' and TYPOS IN (" + PolTim + ")")
        If R(0).Value > 0 Then

            MsgBox("Προσοχή βρέθηκαν " + Str(R(0)) + " πελατες με ΑΦΜ=000000000 και θα μεταφερθούν στις λιανικές")
            gdb.Execute("UPDATE TIMKEPYO SET POLHS=3,AFM='000000000',EPO='',EPA='',DIE='',DOY=''   WHERE AFM LIKE '00000000%' and TYPOS IN (" + PolTim + ")", k)
        End If


        'ΠΕΤΑΕΙ ΤΟ ΕΞΩΤΕΡΙΚΟ ΕΞΩ ΕΝΔΟΚ+3ΧΩΡΕΣ
        gdb.Execute("UPDATE TIMKEPYO SET POLHS=9,AFM='999999999',EPO='',EPA='',DIE='',DOY=''   WHERE AFM IN (SELECT AFM FROM PEL WHERE  LEFT(PEL.TYPOS ,1) IN ('3','2') )  and TYPOS IN (" + PolTim + ")", k)







        ' Gdb.Execute "UPDATE TIMKEPYO SET SET POLHS=3,AFM='000000000',EPO='',EPA='',DIE='',DOY=''   WHERE AFM='000000000' and POLHS=1", K

        gdb.Execute("UPDATE TIMKEPYO SET POLHS=2 WHERE TYPOS IN (" + AgoTim + "," + AgoPis + "," + AGOEXO + ")", NNAG)
        'gdb.Execute("UPDATE TIMKEPYO SET POLHS=2 WHERE TYPOS IN (" + AgoTim + ")")
        gdb.Execute("UPDATE TIMKEPYO SET POLHS=1,PIS='credit'  WHERE TYPOS IN (" + PolPis + ")")
        gdb.Execute("UPDATE TIMKEPYO SET POLHS=2,PIS='credit'  WHERE TYPOS IN (" + AgoPis + ")")

        gdb.Execute("UPDATE TIMKEPYO SET POLHS=3,PIS='CASH'  WHERE TYPOS IN (" + PolTam + "," + PolTamMhxPIS + ")")
        gdb.Execute("UPDATE TIMKEPYO SET AFM='000000000',EPO='',EPA='',DIE='',DOY=''   WHERE TYPOS IN (" + PolTam + ")")




        gdb.Execute("UPDATE TIMKEPYO SET POLHS=4,PIS='CASH'  WHERE TYPOS IN (" + PolTamMhx + ")")
        gdb.Execute("UPDATE TIMKEPYO SET POLHS=5,PIS='CREDIT'  WHERE TYPOS IN (" + AGOEXO + ")")

240:    If nn = 0 Then
250:        MsgBox("ΔΕΝ ΔΗΜΙΟΥΡΓΗΘΗΚΕ ΤΟ ΑΡΧΕΙΟ ΠΩΛΗΣΕΩΝ")
260:        MsgBox(Err.Description)

            ' Exit Sub
        Else
            MsgBox("ΔΗΜΙΟΥΡΓΗΘΗΚΕ ΤΟ ΑΡΧΕΙΟ ΠΩΛΗΣΕΩΝ")

        End If



        If NNAG = 0 Then
            MsgBox("ΔΕΝ ΔΗΜΙΟΥΡΓΗΘΗΚΕ ΤΟ ΑΡΧΕΙΟ ΑΓΟΡΩΝ")
            MsgBox(Err.Description)
            'gdb.Close()
            'Exit Sub
        Else
            MsgBox("ΔΗΜΙΟΥΡΓΗΘΗΚΕ ΤΟ ΑΡΧΕΙΟ ΑΓΟΡΩΝ")

        End If

        If nn = 0 And NNAG = 0 Then
            gdb.Close()
            Exit Sub
        End If





        Dim a As Integer
        R.Open("select * from TIMKEPYO", gdb, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        Do While Not R.EOF
            a = 1 ' check_afm(R("AFM").ToString)

            If a = 0 Then
                MsgBox(R("EPO"))
            End If

            R.MoveNext()
        Loop

        R.Close()

        gdb.Execute("SELECT POLHS,PIS,AFM,SUM(AJIA) AS AJIA,SUM(FPA) AS FPA,SUM(AEG) AS TEM,MAX(HME) AS SHME into SYNOLAKEPYO FROM TIMKEPYO GROUP BY POLHS,AFM,PIS ORDER BY POLHS,AFM")
        ' gdb.Execute("SELECT POLHS,PIS,AFM,SUM(AJIA) AS AJIA,SUM(FPA) AS FPA,SUM(AEG) AS TEM,MAX(HME) AS SHME,EPO,EPA,DIE,DOY into SYNOLAKEPYO FROM TIMKEPYO GROUP BY POLHS,AFM,PIS,EPO,DOY,EPA,DIE ORDER BY POLHS,AFM")
        gdb.Close()
        'print3_xar "SELECT POLHS,AFM,EPO,DOY,EPA,DIE,SUM(AJIA) AS [ΣΥΝ.ΑΞΙΑ],SUM(FPA) AS [ΣΥΝ.ΦΠΑ],SUM(AEG) AS [ΑΡ.ΤΙΜΟΛ]  FROM TIMKEPYO GROUP BY POLHS,AFM,EPO,DOY,EPA,DIE ORDER BY POLHS,AFM", "11111111", "ΣΥΓΚΕΝΤΡΩΤΙΚΗ ΑΠΟ " + Format(D1, "DD/MM/YYYY") + " ΕΩΣ " + Format(d2, "DD/MM/YYYY"), 0 ' RR.RecordSource
        MsgBox("ok")
        ' Adodc1.ConnectionString = gConnect
        'Adodc1.RecordSource = "SELECT * FROM SYNOLAKEPYO"
        ' Adodc1.Refresh()
    End Sub
    Function check_afm(ByVal M_AFM As String) As Integer

        '<EhHeader>
        On Error GoTo check_afm_Err
        M_AFM = M_AFM.Trim
        '</EhHeader>
        Dim SUMA, k As Long
        Dim l As Integer = Len(M_AFM)
100:    SUMA = 0
110:    check_afm = 1
120:    k = 1

130:    For k = 1 To 8
140:        SUMA = SUMA + Val(Mid(M_AFM, k, 1)) * 2 ^ (9 - k)
        Next

150:    If SUMA Mod 11 <> Val(Mid(M_AFM, l, 1)) Then
160:        If SUMA Mod 11 = 10 And Val(Mid(M_AFM, l, 1)) = 0 Then
            Else
170:            MsgBox("Λάθος στο ΑΦΜ " + M_AFM)
180:            check_afm = 0
            End If
        End If
        If l <> 9 Then
            check_afm = 0
        End If
        '<EhFooter>
        Exit Function

check_afm_Err:

        Resume Next

        '</EhFooter>

    End Function



    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NOTASIMPORT.Click
        'ΝΟΤΑΣ 
        '6 ΔΕΛ.ΛΙΑΝ ΑΛΠ   2=ΤΙΜ ΕΞΩ,ΤΔΑ 1=ΤΙΜ    3=ΔΕΛΤΙΑ ΑΠΟΣΤΟΛΗΣ ΔΑΠ    4=ΠΕΠ  7=ΕΠΙΣΤΡ.ΛΙΑΝ    

        If checkServer() Then
        Else
            Exit Sub
        End If

        Text2.Text = VB6.Format(DateTimePicker2, "MM")
        Text3.Text = VB6.Format(DateTimePicker2, "YYYY")

        Text4.Text = VB6.Format(DateTimePicker2, "DD/MM/YYYY")
        Text1.Text = "029234870"


        D1.Value = DateTimePicker1.Value
        D2.Value = DateTimePicker2.Value

        ' pel(ROW, 2)  πινακας που φορτώνει ολους τους πελατες απο το φυλλο 2
        ' 
        Dim debug As Boolean = False


        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        'Dim xlAppPel As Excel.Application
        'Dim xlWorkBookPel As Excel.Workbook
        Dim xl As Excel.Worksheet
        Dim xlPEL As Excel.Worksheet
        Dim xlok As Excel.Worksheet

        If Len(TextBox3.Text) < 2 Then
            MsgBox("ΔΕΝ ΔΙΑΛΕΞΑΤΕ ΑΡΧΕΙΟ EXCEL ")

            Exit Sub
        End If

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open(TextBox3.Text)
        ' xlWorkBook.Worksheets.Add()
        ' xl = xlWorkBook.Worksheets(2) ' .Add
        xlok = xlWorkBook.Worksheets(1)
        '        xlAppPel = New Excel.ApplicationClass
        '       xlWorkBookPel = xlApp.Workbooks.Open(TextBox2.Text)
        '      xlPEL = xlWorkBookPel.Worksheets(1)

        'metafora me σωστη γραμμογραφηση στο 3
        '=========================================
        '===============================================================================real onomatepvmymo 54100
        Dim nRows As Long  'ποσα τιμολογια εχει
        'ExecuteSQLQuery ("CREATE 
        ROW = 4

        '
        ExecuteSQLQuery("SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'NOTAS'")
        Dim SQL2 As String

        If sqlDT.Rows(0)(0) = 0 Then

            SQL2 = "CREATE TABLE [dbo].[NOTAS](" _
             & "[HME] [datetime] NULL," _
             & "[AFM] [nchar](10) NULL," _
             & "[AJI] [numeric](18, 2) NULL," _
             & "[FPA] [numeric](18, 2) NULL," _
             & "[ATIM] [nchar](15) NULL," _
             & "[TYPOS] [int] NULL," _
             & "[ID] [int] IDENTITY(1,1) NOT NULL" _
            & ") ON [PRIMARY]"

            ExecuteSQLQuery(SQL2)

        End If

        ExecuteSQLQuery("DELETE FROM NOTAS")


        Dim SQL As String
        Do While True
            ROW = ROW + 1
            If debug Then
                If ROW > 100 Then Exit Do
            End If



            If xlok.Cells(ROW, 1).value = Nothing Then
                Exit Do
            End If

            SQL = "INSERT INTO NOTAS (HME,AFM,AJI,FPA,ATIM,TYPOS) VALUES("
            SQL = SQL + "'" + VB6.Format(xlok.Cells(ROW, 1), "MM/DD/YYYY") + "',"  'HME

            If xlok.Cells(ROW, 2).VALUE.ToString = Nothing Then
                SQL = SQL + "'000000000',"
            Else
                SQL = SQL + "'" + xlok.Cells(ROW, 2).VALUE.ToString + "',"  ' AFM

            End If


            'If xlok.Cells(ROW, 1).value = Nothing Then
            'Exit Do
            'End If


            SQL = SQL + "" + Replace(VB6.Format(xlok.Cells(ROW, 3), "######0.00"), ",", ".") + "," 'AJI
            SQL = SQL + "" + Replace(VB6.Format(xlok.Cells(ROW, 4), "######0.00"), ",", ".") + "," 'FPA
            SQL = SQL + "'" + xlok.Cells(ROW, 5).VALUE.ToString + "',"  ' ATIM
            SQL = SQL + "" + Replace(VB6.Format(xlok.Cells(ROW, 6), "000"), ",", ".") + ")" 'TYPOS
            ExecuteSQLQuery(SQL)
            Me.Text = ROW
            'Exit Do

        Loop


        'xlWorkBook.Save()
        xlApp.Quit()


        ExecuteSQLQuery("select SUM(AJI) AS SAJI , TYPOS FROM NOTAS GROUP BY TYPOS")
        Dim K As Integer
        ListBox2.Items.Clear()
        For K = 0 To sqlDT.Rows.Count - 1
            ListBox2.Items.Add(sqlDT.Rows(K)(1).ToString + "--" + sqlDT.Rows(K)(0).ToString)

        Next


        'ΚΑΝΩ ΑΡΝΤΗΤΙΚΕΣ ΤΙΣ ΛΙΑΝΙΚΕΣ ΕΠΙΣΤΡΟΦΕΣ ΓΙΑ ΝΑ ΓΙΝΕΙ Η ΣΟΥΜΑ ΛΙΑΝΙΚΩΝ ΣΩΣΤΑ
        ExecuteSQLQuery("update NOTAS SET AJI=-AJI  WHERE  TYPOS IN (7) ")

        'ΚΑΝΩ ΤΙΣ ΕΠΙΣΤΡΟΦΕΣ =6 ΓΙΑ ΝΑ ΜΗΝ ΕΧΩ 2 ΤΥΠΟΥΣ ΣΤΙΣ ΛΙΑΝΙΚΕΣ 
        ExecuteSQLQuery("update NOTAS SET TYPOS=6   WHERE  TYPOS IN (7) ")

        'ΣΒΗΝΩ ΤΑ ΔΕΛΤΙΑ ΑΠΟΣΤΟΛΗΣ
        ExecuteSQLQuery("DELETE FROM  NOTAS WHERE TYPOS=3")







        '  ExecuteSQLQuery("if EXISTS (SELECT * FROM SYNOLAKEPYO2)  drop table  SYNOLAKEPYO2 ")

        ExecuteSQLQuery("if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = 'SYNOLAKEPYO2' AND TABLE_SCHEMA = 'dbo') drop table dbo.SYNOLAKEPYO2;")


        Dim aa As String
        '6 ΔΕΛ.ΛΙΑΝ ΑΛΠ   2=ΤΙΜ ΕΞΩ,ΤΔΑ 1=ΤΙΜ    3=ΔΕΛΤΙΑ ΑΠΟΣΤΟΛΗΣ ΔΑΠ    4=ΠΕΠ  7=ΕΠΙΣΤΡ.ΛΙΑΝ    
        aa = "SELECT 0 AS POLHS,'       ' AS PIS,SUM(AJI) AS SAJI,SUM(FPA) AS SFPA,COUNT(*) AS TEM,MAX(HME) AS SHME ," _
        & "TYPOS, AFM INTO SYNOLAKEPYO2 FROM NOTAS WHERE   HME>='" + Format(DateTimePicker1, "MM/DD/YYYY") + "' AND HME<='" + Format(DateTimePicker2, "MM/DD/YYYY") + "' GROUP BY TYPOS,AFM"
        ExecuteSQLQuery(aa)
        ExecuteSQLQuery("update SYNOLAKEPYO2 SET PIS=(CASE WHEN  TYPOS IN (4) THEN 'credit' else 'normal' END)")


        ExecuteSQLQuery("update SYNOLAKEPYO2 SET POLHS=(CASE WHEN  TYPOS IN (6,7) THEN 3 else 1 END)")


        ExecuteSQLQuery("if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = 'SYNOLAKEPYO' AND TABLE_SCHEMA = 'dbo') drop table dbo.SYNOLAKEPYO;")

        ExecuteSQLQuery("SELECT POLHS,PIS,AFM,SAJI AS AJIA,SFPA AS FPA,TEM,SHME INTO SYNOLAKEPYO FROM SYNOLAKEPYO2 ")


        MsgBox("OK")



    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'ΑΡΧΕΙΟ ΚΙΝΗΣΕΩΝ
        If Len(Trim(TextBox3.Text)) = 0 Then
            CD1.ShowDialog()
            TextBox3.Text = CD1.FileName
        Else
            If Len(Dir(LTrim(TextBox3.Text), FileAttribute.Normal)) < 2 Then
                MsgBox("δεν υπάρχει το αρχείο " & TextBox3.Text)
                Exit Sub
            End If
        End If
    End Sub

    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles xmlAGORON.Click
        ' αγορων
        'pvlhsevn
        '==========================================================================================
        On Error Resume Next

        'DBGrid1.Clear
        If checkServer() Then
        Else
            Exit Sub
        End If
        ' gconnect=":HP530\SQL2012:sa:12345678:1:perp"
        Dim M_AFM As String : M_AFM = Text1.Text
        Dim m_mhnas As String : m_mhnas = Text2.Text
        Dim m_etos As String : m_etos = Text3.Text
        Dim hmer As String = Text4.Text  '"31/03/2015"
        hmer = VB6.Format(hmer, "yyyy-mm-dd")

        If Len(M_AFM) <> 9 Then
            ' MsgBox("λαθος στο ΑΦΜ")
            'Exit Sub
        End If

        If Len(Dir("C:\SYGK", vbDirectory)) = 0 Then
            MkDir("C:\SYGK")
        End If

        Dim file

        Dim F_CASH

        F_CASH = arTam.Text ' "ΣΥ09002067"
        file = "C:\SYGK\XML.TXT"
        Kill(file)

        Dim R As New ADODB.Recordset

        'Open "C:\SYGK\synola.txt" For Output As #5
        Dim m_filename As String

        m_filename = "C:\SYGK\" + Text1.Text + "_" + Text3.Text + Text2.Text + "b.XML"

        Dim writer As New XmlTextWriter(m_filename, System.Text.Encoding.UTF8)
        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2
        ' Create a Continent element and set its value to
        ' that of the New Continent dialog box
        'writer.WriteAttributeString("Table", , "sadasasd sdsd")

        '<Data Name="SX" Style="BRowse"><row name="Hd" mode="3" rowId="7">
        writer.WriteStartElement("packages")
        writer.WriteStartElement("package")
        writer.WriteAttributeString("actor_afm", M_AFM)
        writer.WriteAttributeString("month", m_mhnas)
        writer.WriteAttributeString("year", m_etos)

        writer.WriteStartElement("groupedExpenses")
        writer.WriteAttributeString("action", "replace")





        '   Open m_filename For Output As #1
        '
        '  Print #1, "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""YES""?>"
        ' Print #1, "<packages>"
        '   Print #1, " <package actor_afm=""" + M_AFM + """ month=""" + m_mhnas + """ year=""" + m_etos + """>"
        '  Print #1, " <groupedRevenues action=""replace"">"


        ExecuteSQLQuery("SELECT * FROM SYNOLAKEPYO WHERE POLHS=2  ")

        Dim sxre As Single, spis As Single


        sxre = 0 : spis = 0
        Dim sxretax As Single = 0
        Dim spistax As Single = 0
        Dim sfpa(10) As Single


        '   DBGrid1.row = 0 : DBGrid1.Col = 1
        '  DBGrid1.Text = "Καθ.Αξία"

        '  DBGrid1.row = 0 : DBGrid1.Col = 2
        '  DBGrid1.Text = "Φ.Π.Α."

        ListBox1.Items.Clear()



        Dim a
        Dim k As Long

        FileOpen(1, "C:\SYGK\ERR.TXT", OpenMode.Output)

        '  <groupedExpenses action="replace">
        '<expense>
        '  <afm>044149178</afm>
        '  <amount>11017,86</amount>
        '  <tax>1432,38</tax>
        '  <invoices>38</invoices>
        '  <note>normal</note>
        '  <nonObl>0</nonObl>
        '  <date>2014-03-28</date>
        '  <!--kefcode=10&hyp;001-->
        '  <!--kefname=BAPANOY ΠPABITA OYPANIA-->
        '</expense>


        For k = 0 To sqlDT.Rows.Count - 1

            writer.WriteStartElement("expense")
            writer.WriteStartElement("afm") : writer.WriteString(sqlDT.Rows(k)("AFM").ToString.Trim) : writer.WriteEndElement()  'AFM
            writer.WriteStartElement("amount") : writer.WriteString(VB6.Format(sqlDT.Rows(k)("ajia"), "#######0.00")) : writer.WriteEndElement()  'AJIA
            writer.WriteStartElement("tax") : writer.WriteString(VB6.Format(sqlDT.Rows(k)("fpa"), "#######0.00")) : writer.WriteEndElement()  'FPA
            writer.WriteStartElement("invoices") : writer.WriteString(VB6.Format(sqlDT.Rows(k)("tem"), "########0")) : writer.WriteEndElement()  'NtIM

            writer.WriteStartElement("note") : writer.WriteString(sqlDT.Rows(k)("pis")) : writer.WriteEndElement()  'CREDIT
            writer.WriteStartElement("nonObl") : writer.WriteString("0") : writer.WriteEndElement()  'υποχρεος
            ' <nonObl>0</nonObl>
            writer.WriteStartElement("date") : writer.WriteString(VB6.Format(sqlDT.Rows(k)("shme"), "yyyy-mm-dd")) : writer.WriteEndElement()  'DATE
            writer.WriteEndElement()  'REVENUE



            a = check_afm(sqlDT.Rows(k)("AFM").ToString)
            If a = 0 Then
                PrintLine(1, sqlDT.Rows(k)("AFM").ToString)
            End If


            If sqlDT.Rows(k)("pis") = "normal" Then
                sxre = sxre + sqlDT.Rows(k)("ajia")
                sxretax = sxretax + sqlDT.Rows(k)("fpa")
                sfpa(1) = sfpa(1) + sqlDT.Rows(k)("fpa")
            Else
                spis = spis + sqlDT.Rows(k)("ajia")
                spistax = spistax + sqlDT.Rows(k)("fpa")

                sfpa(2) = sfpa(2) + sqlDT.Rows(k)("fpa")
            End If


            ' R.MoveNext()
        Next 'Loop

        FileClose(1)



        writer.WriteEndElement()  'groupedREVENUE


        '        DBGrid1.row = 1 : DBGrid1.Col = 0
        ListBox1.Items.Add("Σύνολο Χρ.τιμ. ")
        '      DBGrid1.Col = 1
        ListBox1.Items.Add(VB6.Format(sxre, "###,###.00"))


        'DBGrid1.Col = 2
        ListBox1.Items.Add(VB6.Format(sfpa(1), "###,###.00"))


        'SFPA(1) = SFPA(1) + R!FPA

        On Error GoTo 0


        'DBGrid1.row = 2 : DBGrid1.Col = 0
        ListBox1.Items.Add("Σύνολο πισ.τιμ. ")
        ' DBGrid1.Col = 1
        ListBox1.Items.Add(VB6.Format(-spis, "####,###.00"))

        ' DBGrid1.Col = 2
        ListBox1.Items.Add(VB6.Format(-sfpa(2), "###,###.00"))



        'Print #5, "Σύνολο Χρ.τιμ. " + Format(sxre, "########0.00")
        'Print #5, " Πιστωτικά " + Format(spis, "########0.00")

        ' writer.WriteEndElement()  'GROUPEDREVENUES





        'action=""replace"">"

        ExecuteSQLQuery("SELECT SUM(AJIA) AS SAJIA,SUM(FPA) AS SFPA FROM SYNOLAKEPYO WHERE POLHS=3")


        Dim slian As Single = 0

        Dim slianTax As Single = 0

        Dim EXEILIANIKES As Integer = 0

        'If (Not IsDBNull(sqlDT.Rows(0)("sajia"))) Then
        '    If sqlDT.Rows(0)("sajia") > 0 Then

        '        writer.WriteStartElement("groupedCashRegisters")
        '        writer.WriteAttributeString("action", "replace")

        '        EXEILIANIKES = 1

        '        For k = 0 To sqlDT.Rows.Count - 1
        '            If (Not IsDBNull(sqlDT.Rows(0)("sajia"))) Then
        '                writer.WriteStartElement("cashregister")

        '                writer.WriteStartElement("amount") : writer.WriteString(VB6.Format(sqlDT.Rows(k)("sajia"), "#######0.00")) : writer.WriteEndElement()  'AJIA
        '                writer.WriteStartElement("tax") : writer.WriteString(VB6.Format(sqlDT.Rows(k)("sfpa"), "#######0.00")) : writer.WriteEndElement()  'FPA
        '                '  writer.WriteStartElement("invoices") : writer.WriteString(VB6.Format(sqlDT.Rows(k)(""), "########0")) : writer.WriteEndElement()  'NtIM
        '                writer.WriteStartElement("date") : writer.WriteString(hmer) : writer.WriteEndElement()  'DATE
        '                writer.WriteEndElement()  'cashregister
        '                slian = slian + sqlDT.Rows(k)("sajia")
        '                slianTax = slianTax + sqlDT.Rows(k)("sfpa")
        '            End If
        '        Next


        '    End If
        'End If


        'ExecuteSQLQuery("SELECT SUM(AJIA) AS SAJIA,SUM(FPA) AS SFPA FROM SYNOLAKEPYO WHERE POLHS=4")


        'Dim sTam As Single = 0
        'Dim sTamTax As Single = 0
        'If (Not IsDBNull(sqlDT.Rows(0)("sajia"))) Then
        '    If sqlDT.Rows(0)("sajia") > 0 Then


        '        ' GRAFO TO BEGGINNING
        '        If EXEILIANIKES = 0 Then
        '            writer.WriteStartElement("groupedCashRegisters")
        '            writer.WriteAttributeString("action", "replace")
        '            EXEILIANIKES = 1
        '        End If




        '        For k = 0 To sqlDT.Rows.Count - 1
        '            writer.WriteStartElement("cashregister")
        '            writer.WriteStartElement("cashreg_id") : writer.WriteString(F_CASH) : writer.WriteEndElement()  'id
        '            writer.WriteStartElement("amount") : writer.WriteString(VB6.Format(sqlDT.Rows(k)("sajia"), "########.00")) : writer.WriteEndElement()  'AJIA
        '            writer.WriteStartElement("tax") : writer.WriteString(VB6.Format(sqlDT.Rows(k)("sfpa"), "########.00")) : writer.WriteEndElement()  'FPA
        '            'writer.WriteStartElement("invoices") : writer.WriteString(VB6.Format(sqlDT.Rows(k)("tem"), "########0")) : writer.WriteEndElement()  'NtIM
        '            writer.WriteStartElement("date") : writer.WriteString(hmer) : writer.WriteEndElement()  'DATE
        '            writer.WriteEndElement()  'CashRegister
        '            sTam = sTam + sqlDT.Rows(k)("sajia")
        '            sTamTax = sTamTax + sqlDT.Rows(k)("sfpa")
        '        Next
        '    End If
        'End If


        '' GRAFO TO BEGGINNING
        'If EXEILIANIKES = 1 Then
        '    writer.WriteEndElement()  'groupedCashRegisters
        'End If




  
        'ListBox1.Items.Add("Σύνολο ταμειακών")
        'ListBox1.Items.Add(VB6.Format(sTam, "####,##0.00"))
        'If IsDBNull(sqlDT.Rows(0)("sfpa")) Then
        '    ListBox1.Items.Add("0")
        'Else
        '    ListBox1.Items.Add(VB6.Format(sqlDT.Rows(0)("sfpa"), "##########.00"))
        'End If
        'ListBox1.Items.Add("------------------------")
        'ListBox1.Items.Add(VB6.Format(sTam + slian + sxre - spis, "##########.00"))
        'ListBox1.Items.Add(VB6.Format(sTamTax + slianTax + sxretax - spistax, "##########.00"))
        'ListBox2.Width = ListBox2.Width * 2




        writer.WriteEndElement()  'PACKAGE
        writer.WriteEndElement()  'PACKAGES
        writer.WriteEndDocument()
        writer.Close()


        '  Print #1, "</package>"
        ' Print #1, "</packages>"

        ' Close #1

        '   Dim k As Integer

        '    For k = 1 To 5
        ' Print #5, Left(DBGrid1.TextMatrix(k, 0) + Space(30), 30) + Right(Space(30) + DBGrid1.TextMatrix(k, 1), 30) + Right(Space(30) + DBGrid1.TextMatrix(k, 2), 30)

        '    Next




        ExecuteSQLQuery("SELECT SUM(AJIA) AS SAJIA,SUM(FPA) AS SFPA FROM SYNOLAKEPYO WHERE POLHS=9")

        If (Not IsDBNull(sqlDT.Rows(0)("sajia"))) Then
            ListBox1.Items.Add("ΕΞΩΤΕΡΙΚΟΥ")
            ListBox1.Items.Add(VB6.Format(sqlDT.Rows(0)("sajia"), "##########.00"))

        End If



        'Close #5

        MsgBox("ΑΠΟΘΗΚΕΥΤΗΚΕ ΤΟ " + m_filename + Chr(3) + " Kαι c:\sygk\synola.txt το αρχείο με τα σύνολα")
        ' End Sub

        ListBox2.Items.Clear()



        'Public Class Form1

        ' Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim myDocument As New XmlDocument
        myDocument.Load(m_filename) ' m_filename)  ' "C:\somefile.xml"
        myDocument.Schemas.Add("", "c:\sygk\gsis_packages_schema.xsd") 'namespace here or empty string
        Dim eventHandler As ValidationEventHandler = New ValidationEventHandler(AddressOf ValidationEventHandler)
        myDocument.Validate(eventHandler)
    End Sub

    Private Sub Button5_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'GRAMOGRAFHSH
        MsgBox("ΚΑΘΑΡΕΣ ΑΞΙΕΣ (στηλες): 5η(Ε) 13% , 6η(F) 23% , 10η(J) 0%  14η συνολ.με φπα  Φπα στηλες: 15η(O) 13% , 9η(I) 23% ")
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim ff As String = "AMBROS.XML"   'filexml.Text ' "GAT.XML" ' "\\Logisthrio\333\pr.export" '

        Dim writer As New XmlTextWriter(ff, System.Text.Encoding.UTF8)
        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2

        writer.WriteStartElement("Table")

        writer.WriteAttributeString("CHME", cHME.Text)
        writer.WriteAttributeString("CPAR", cPAR.Text)
        writer.WriteAttributeString("CKOD", cKOD.Text)
        writer.WriteAttributeString("CEPO", cEPO.Text)
        writer.WriteAttributeString("CKAU24", cKAU24.Text)


        writer.WriteAttributeString("CKAU13", cKAU13.Text)
        writer.WriteAttributeString("CKAU0", cKAU0.Text)

        writer.WriteAttributeString("CFPA24", cFPA24.Text)
        writer.WriteAttributeString("CFPA13", cFPA13.Text)

        writer.WriteEndElement() ' TABLE



        writer.WriteEndDocument()
        writer.Close()




    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim xmlDoc As New XmlDocument()


        xmlDoc.Load("AMBROS.XML") '"GAT.xml")

        Dim nodes As XmlNodeList = xmlDoc.DocumentElement.SelectNodes("/Table")
        Dim pID As String = "", pName As String = "", pPrice As String = ""


        For Each node As XmlNode In nodes
            cHME.Text = node.Attributes("CHME").Value  ' ΔΙΑΒΑΖΩ ΤΟ ATTRIBUTE
            cPAR.Text = node.Attributes("CPAR").Value

            cKOD.Text = node.Attributes("CKOD").Value  ' ΔΙΑΒΑΖΩ ΤΟ ATTRIBUTE
            cEPO.Text = node.Attributes("CEPO").Value

            cKAU24.Text = node.Attributes("CKAU24").Value  ' ΔΙΑΒΑΖΩ ΤΟ ATTRIBUTE
            cKAU13.Text = node.Attributes("CKAU13").Value
            cKAU0.Text = node.Attributes("CKAU0").Value  ' ΔΙΑΒΑΖΩ ΤΟ ATTRIBUTE
            cFPA24.Text = node.Attributes("CFPA24").Value
            cFPA13.Text = node.Attributes("CFPA13").Value  ' ΔΙΑΒΑΖΩ ΤΟ ATTRIBUTE
        Next

    End Sub
End Class
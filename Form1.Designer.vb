﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.toXML = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.EditConnString = New System.Windows.Forms.Button()
        Me.APO = New System.Windows.Forms.DateTimePicker()
        Me.EOS = New System.Windows.Forms.DateTimePicker()
        Me.ListBox2 = New System.Windows.Forms.ListBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.UPDATE_TIM = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.CancInv = New System.Windows.Forms.Button()
        Me.RequestTransmittedDocs = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.Button9 = New System.Windows.Forms.Button()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Button10 = New System.Windows.Forms.Button()
        Me.Button11 = New System.Windows.Forms.Button()
        Me.Button12 = New System.Windows.Forms.Button()
        Me.tickToParastat = New System.Windows.Forms.Button()
        Me.Button13 = New System.Windows.Forms.Button()
        Me.Button14 = New System.Windows.Forms.Button()
        Me.Button15 = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(161, Byte))
        Me.Button1.Location = New System.Drawing.Point(12, 27)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(311, 28)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "2.ΑΠΟΣΤΟΛΗ ΑΡΧΕΙΟΥ ΣΕ ΑΑΔΕ"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(12, 109)
        Me.TextBox2.Multiline = True
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(862, 282)
        Me.TextBox2.TabIndex = 1
        '
        'toXML
        '
        Me.toXML.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(161, Byte))
        Me.toXML.Location = New System.Drawing.Point(12, 3)
        Me.toXML.Name = "toXML"
        Me.toXML.Size = New System.Drawing.Size(311, 23)
        Me.toXML.TabIndex = 2
        Me.toXML.Text = "1.ΔΗΜΙΟΥΡΓΙΑ ΑΡΧΕΙΟΥ XML ΓΙΑ ΑΠΟΣΤΟΛΗ"
        Me.toXML.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(880, 291)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(264, 23)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "4.SendIncomeClassification"
        Me.Button3.UseVisualStyleBackColor = True
        Me.Button3.Visible = False
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'EditConnString
        '
        Me.EditConnString.Location = New System.Drawing.Point(880, 426)
        Me.EditConnString.Name = "EditConnString"
        Me.EditConnString.Size = New System.Drawing.Size(264, 23)
        Me.EditConnString.TabIndex = 4
        Me.EditConnString.Text = "opendatabase"
        Me.EditConnString.UseVisualStyleBackColor = True
        Me.EditConnString.Visible = False
        '
        'APO
        '
        Me.APO.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(161, Byte))
        Me.APO.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.APO.Location = New System.Drawing.Point(744, -1)
        Me.APO.Name = "APO"
        Me.APO.Size = New System.Drawing.Size(130, 26)
        Me.APO.TabIndex = 5
        '
        'EOS
        '
        Me.EOS.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(161, Byte))
        Me.EOS.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.EOS.Location = New System.Drawing.Point(744, 31)
        Me.EOS.Name = "EOS"
        Me.EOS.Size = New System.Drawing.Size(130, 26)
        Me.EOS.TabIndex = 6
        '
        'ListBox2
        '
        Me.ListBox2.FormattingEnabled = True
        Me.ListBox2.Location = New System.Drawing.Point(329, 60)
        Me.ListBox2.Name = "ListBox2"
        Me.ListBox2.Size = New System.Drawing.Size(545, 43)
        Me.ListBox2.TabIndex = 8
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(880, 455)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(264, 22)
        Me.Button2.TabIndex = 9
        Me.Button2.Text = "check xml"
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'UPDATE_TIM
        '
        Me.UPDATE_TIM.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(161, Byte))
        Me.UPDATE_TIM.Location = New System.Drawing.Point(12, 60)
        Me.UPDATE_TIM.Name = "UPDATE_TIM"
        Me.UPDATE_TIM.Size = New System.Drawing.Size(311, 43)
        Me.UPDATE_TIM.TabIndex = 10
        Me.UPDATE_TIM.Text = "3.ΕΝΗΜΕΡΩΣΗ ΒΑΣΗΣ ΜΕ ΤΗΝ ΑΠΟΣΤΟΛΗ"
        Me.UPDATE_TIM.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(12, 397)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidth = 51
        Me.DataGridView1.Size = New System.Drawing.Size(862, 237)
        Me.DataGridView1.TabIndex = 11
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Button4.Location = New System.Drawing.Point(919, 4)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(252, 28)
        Me.Button4.TabIndex = 12
        Me.Button4.Text = "ΤΙΜΟΛΟΓΗΣΕΙΣ ΠΡΟΣ ΕΜΑΣ"
        Me.Button4.UseVisualStyleBackColor = False
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(880, 323)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(264, 23)
        Me.Button5.TabIndex = 13
        Me.Button5.Text = "5.UpdateDB with IncResponse"
        Me.Button5.UseVisualStyleBackColor = True
        Me.Button5.Visible = False
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.Lime
        Me.Button6.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(161, Byte))
        Me.Button6.Location = New System.Drawing.Point(329, 1)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(403, 54)
        Me.Button6.TabIndex = 14
        Me.Button6.Text = "ΑΥΤΟΜΑΤΟΠΟΙΗΜΕΝΗ ΑΠΟΣΤΟΛΗ"
        Me.Button6.UseVisualStyleBackColor = False
        '
        'CancInv
        '
        Me.CancInv.BackColor = System.Drawing.Color.Lime
        Me.CancInv.Location = New System.Drawing.Point(919, 382)
        Me.CancInv.Name = "CancInv"
        Me.CancInv.Size = New System.Drawing.Size(253, 23)
        Me.CancInv.TabIndex = 15
        Me.CancInv.Text = "Ακύρωση Απεσταλμένου Παρ/κού"
        Me.CancInv.UseVisualStyleBackColor = False
        '
        'RequestTransmittedDocs
        '
        Me.RequestTransmittedDocs.BackColor = System.Drawing.Color.Lime
        Me.RequestTransmittedDocs.Location = New System.Drawing.Point(891, 526)
        Me.RequestTransmittedDocs.Name = "RequestTransmittedDocs"
        Me.RequestTransmittedDocs.Size = New System.Drawing.Size(253, 28)
        Me.RequestTransmittedDocs.TabIndex = 16
        Me.RequestTransmittedDocs.Text = "ΤΙΜΟΛΟΓΗΣΕΙΣ ΣΕ ΤΡΙΤΟΥΣ"
        Me.RequestTransmittedDocs.UseVisualStyleBackColor = False
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(896, 611)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(248, 23)
        Me.Button7.TabIndex = 17
        Me.Button7.Text = "Button7"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'Button8
        '
        Me.Button8.Location = New System.Drawing.Point(880, 497)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(264, 23)
        Me.Button8.TabIndex = 18
        Me.Button8.Text = "διαβαζει τα τιμολογια προμηθευτων"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'Button9
        '
        Me.Button9.Location = New System.Drawing.Point(981, 556)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(186, 30)
        Me.Button9.TabIndex = 19
        Me.Button9.Text = "Button9"
        Me.Button9.UseVisualStyleBackColor = True
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(309, 382)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(565, 264)
        Me.ListBox1.TabIndex = 20
        '
        'Button10
        '
        Me.Button10.Location = New System.Drawing.Point(919, 33)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(252, 28)
        Me.Button10.TabIndex = 21
        Me.Button10.Text = "ΚΤΕΛ1"
        Me.Button10.UseVisualStyleBackColor = True
        '
        'Button11
        '
        Me.Button11.Location = New System.Drawing.Point(919, 70)
        Me.Button11.Name = "Button11"
        Me.Button11.Size = New System.Drawing.Size(252, 28)
        Me.Button11.TabIndex = 22
        Me.Button11.Text = "ΚΤΕΛ2 login"
        Me.Button11.UseVisualStyleBackColor = True
        '
        'Button12
        '
        Me.Button12.Location = New System.Drawing.Point(919, 121)
        Me.Button12.Name = "Button12"
        Me.Button12.Size = New System.Drawing.Size(252, 28)
        Me.Button12.TabIndex = 23
        Me.Button12.Text = "get ticket"
        Me.Button12.UseVisualStyleBackColor = True
        '
        'tickToParastat
        '
        Me.tickToParastat.Location = New System.Drawing.Point(919, 174)
        Me.tickToParastat.Name = "tickToParastat"
        Me.tickToParastat.Size = New System.Drawing.Size(252, 28)
        Me.tickToParastat.TabIndex = 24
        Me.tickToParastat.Text = "ΕΙΣΙΤΗΡΙΑ->ΠΑΡΑΣΤΑΤΙΚΑ"
        Me.tickToParastat.UseVisualStyleBackColor = True
        '
        'Button13
        '
        Me.Button13.Location = New System.Drawing.Point(920, 208)
        Me.Button13.Name = "Button13"
        Me.Button13.Size = New System.Drawing.Size(252, 28)
        Me.Button13.TabIndex = 25
        Me.Button13.Text = "PAROXOS ΛΟΓΙΝ"
        Me.Button13.UseVisualStyleBackColor = True
        '
        'Button14
        '
        Me.Button14.Location = New System.Drawing.Point(920, 242)
        Me.Button14.Name = "Button14"
        Me.Button14.Size = New System.Drawing.Size(252, 28)
        Me.Button14.TabIndex = 26
        Me.Button14.Text = "Jinvoice"
        Me.Button14.UseVisualStyleBackColor = True
        '
        'Button15
        '
        Me.Button15.Location = New System.Drawing.Point(880, 352)
        Me.Button15.Name = "Button15"
        Me.Button15.Size = New System.Drawing.Size(263, 30)
        Me.Button15.TabIndex = 27
        Me.Button15.Text = "forologikos"
        Me.Button15.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1203, 661)
        Me.Controls.Add(Me.Button15)
        Me.Controls.Add(Me.Button14)
        Me.Controls.Add(Me.Button13)
        Me.Controls.Add(Me.tickToParastat)
        Me.Controls.Add(Me.Button12)
        Me.Controls.Add(Me.Button11)
        Me.Controls.Add(Me.Button10)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Button9)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.RequestTransmittedDocs)
        Me.Controls.Add(Me.CancInv)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.UPDATE_TIM)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.ListBox2)
        Me.Controls.Add(Me.EOS)
        Me.Controls.Add(Me.APO)
        Me.Controls.Add(Me.EditConnString)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.toXML)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents toXML As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents EditConnString As Button
    Friend WithEvents APO As DateTimePicker
    Friend WithEvents EOS As DateTimePicker
    Friend WithEvents ListBox2 As ListBox
    Friend WithEvents Button2 As Button
    Friend WithEvents UPDATE_TIM As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents CancInv As Button
    Friend WithEvents RequestTransmittedDocs As Button
    Friend WithEvents Button7 As Button
    Friend WithEvents Button8 As Button
    Friend WithEvents Button9 As Button
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents Button10 As Button
    Friend WithEvents Button11 As Button
    Friend WithEvents Button12 As Button
    Friend WithEvents tickToParastat As Button
    Friend WithEvents Button13 As Button
    Friend WithEvents Button14 As Button
    Friend WithEvents Button15 As Button
End Class

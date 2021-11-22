Public Class Form1

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        Dim ans As Integer
        ans = MsgBox("Do you want to save this document?", vbYesNo, "Save")
        If ans = 6 Then
            SaveFileDialog1.Filter = "Rich Text format(.rtf)|*.rtf|Text documents(*.txt)|*.txt|All Files(*.*)|(*.*)"
            If SaveFileDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
                RichTextBox1.SaveFile(SaveFileDialog1.FileName)
            End If
        Else
            RichTextBox1.Clear()

        End If
    End Sub

      Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        Dim ans As Integer
        ans = MsgBox("Do you want to save this document?", vbYesNo, "Save")
        If ans = 6 Then
            SaveFileDialog1.Filter = "Rich Text format(.rtf)|*.rtf|Text documents(*.txt)|*.txt|All Files(*.*)|(*.*)"
            If SaveFileDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
                RichTextBox1.SaveFile(SaveFileDialog1.FileName)
            End If
        Else
            RichTextBox1.Clear()

        End If
    End Sub

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click

        Dim ans As Integer
        ans = MsgBox("Do you want to save this document?", vbYesNo, "Save")
        If ans = 6 Then
            SaveFileDialog1.Filter = "Rich Text format(.rtf)|*.rtf|Text documents(*.txt)|*.txt|All Files(*.*)|(*.*)"
            If SaveFileDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
                RichTextBox1.SaveFile(SaveFileDialog1.FileName)
            End If
        Else
            RichTextBox1.Clear()

        End If
    End Sub


    Private Sub CloseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseToolStripMenuItem.Click
        PrintDialog1.Document = PrintDocument1
        PrintDialog1.PrinterSettings = PrintDocument1.PrinterSettings
        PrintDialog1.AllowSomePages = True
        If PrintDialog1.ShowDialog = DialogResult.OK Then
            PrintDocument1.PrinterSettings = PrintDialog1.PrinterSettings
            PrintDocument1.Print()
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Dim ex As Integer
        ex = MsgBox("Do you wnt to exit?", vbYesNo, "Exit")
        If ex = 6 Then
            End
        Else
            RichTextBox1.Show()
        End If
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CutToolStripMenuItem.Click
        RichTextBox1.Cut()
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem.Click
        RichTextBox1.Copy()
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteToolStripMenuItem.Click
        RichTextBox1.Paste()

    End Sub

    Private Sub FontToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FontToolStripMenuItem.Click
        If FontDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            RichTextBox1.SelectionFont = FontDialog1.Font
        End If
    End Sub

    Private Sub ColorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ColorToolStripMenuItem.Click
        If ColorDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            RichTextBox1.SelectionColor = ColorDialog1.Color
        End If
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim name, r_word, w_word, temp, result As String
        Dim length, r_length, w_length, i As Integer

        name = RichTextBox1.Text
        name = Trim(RichTextBox1.Text)
        length = Len(name)

        r_word = TextBox1.Text
        r_word = Trim(TextBox1.Text)
        r_length = Len(r_word)

        w_word = TextBox2.Text
        w_word = Trim(TextBox2.Text)
        w_length = Len(w_word)

        result = ""
        For i = 1 To length
            temp = Mid(name, i, r_length)

            If temp = r_word Then
                result = result & w_word
                i = i + r_length - 1
            Else
                result = result & Mid(name, i, 1)
            End If
        Next
        RichTextBox1.Text = result
        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub

    
    Private Sub SpaceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SpaceToolStripMenuItem.Click
        Dim st As String
        Dim ln, sp As Integer
        sp = 0
        st = Me.RichTextBox1.Text
        ln = Len(st)
        For i = 1 To ln
            If Mid(st, i, 1) = " " Then
                sp = sp + 1
            End If
        Next
        MsgBox("Total " & sp & " space are here!")
    End Sub

    Private Sub VowelToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VowelToolStripMenuItem.Click
        Dim st As String
        Dim ln, v As Integer
        v = 0
        st = Me.RichTextBox1.Text
        ln = Len(st)
        For i = 1 To ln
            If Mid(st, i, 1) = "a" Or Mid(st, i, 1) = "e" Or Mid(st, i, 1) = "i" Or Mid(st, i, 1) = "o" Or Mid(st, i, 1) = "u" Or Mid(st, i, 1) = "A" Or Mid(st, i, 1) = "E" Or Mid(st, i, 1) = "I" Or Mid(st, i, 1) = "O" Or Mid(st, i, 1) = "U" Then
                v = v + 1
            End If
        Next
        MsgBox("Total " & v & " Vowel are here!")
    End Sub

    Private Sub ConsonantToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConsonantToolStripMenuItem.Click
        Dim st As String
        Dim ln, no, sp, v, c As Integer
        no = 0
        sp = 0
        v = 0
        st = Me.RichTextBox1.Text
        ln = Len(st)
        For i = 1 To ln
            If Mid(st, i, 1) = "a" Or Mid(st, i, 1) = "e" Or Mid(st, i, 1) = "i" Or Mid(st, i, 1) = "o" Or Mid(st, i, 1) = "u" Or Mid(st, i, 1) = "A" Or Mid(st, i, 1) = "E" Or Mid(st, i, 1) = "I" Or Mid(st, i, 1) = "O" Or Mid(st, i, 1) = "U" Then
                v = v + 1
            End If
        Next
        For i = 1 To ln
            If Mid(st, i, 1) = " " Then
                sp = sp + 1
            End If
        Next
        For i = 1 To ln
            If Mid(st, i, 1) = "1" Or Mid(st, i, 1) = "2" Or Mid(st, i, 1) = "3" Or Mid(st, i, 1) = "4" Or Mid(st, i, 1) = "5" Or Mid(st, i, 1) = "6" Or Mid(st, i, 1) = "7" Or Mid(st, i, 1) = "8" Or Mid(st, i, 1) = "9" Or Mid(st, i, 1) = "0" Then
                no = no + 1
            End If
        Next
        c = ln - (sp + v + no)
        MsgBox("Total " & c & " consonant are here!")
    End Sub


    Private Sub NumberToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumberToolStripMenuItem.Click
        Dim st As String
        Dim ln, no As Integer
        no = 0
        st = Me.RichTextBox1.Text
        ln = Len(st)
        For i = 1 To ln
            If Mid(st, i, 1) = "1" Or Mid(st, i, 1) = "2" Or Mid(st, i, 1) = "3" Or Mid(st, i, 1) = "4" Or Mid(st, i, 1) = "5" Or Mid(st, i, 1) = "6" Or Mid(st, i, 1) = "7" Or Mid(st, i, 1) = "8" Or Mid(st, i, 1) = "9" Or Mid(st, i, 1) = "0" Then
                no = no + 1
            End If
        Next
        MsgBox("Total " & no & " numbers are here!")
    End Sub

    
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim a As String
        a = UCase(Trim(RichTextBox1.Text))
        RichTextBox1.Text = a
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim a As String
        a = LCase(Trim(RichTextBox1.Text))
        RichTextBox1.Text = a
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim l, i As Integer
        Dim x, y, z As String
        z = ""
        x = " " + Trim(RichTextBox1.Text)
        l = Len(x)
        For i = 1 To l
            y = Mid(x, i, 1)
            If y = " " Then
                z = z + UCase(Mid(x, i + 1, 1))
            Else
                z = z + LCase(Mid(x, i + 1, 1))


            End If
        Next
        RichTextBox1.Text = z
    End Sub

    Private Sub ContactUsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ContactUsToolStripMenuItem.Click
        MsgBox("For more information contact:--")
        MsgBox("Mobile->+917278422131")
        MsgBox("e-Mail->sudhir.barik981@gmail.com")
        MsgBox("sudhir.barik981@Yahoo.com")
        MsgBox(">sudhir.barik981@rediffmail.com")
    End Sub

    Private Sub UndoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UndoToolStripMenuItem.Click
        'check if textbox can undo
        If RichTextBox1.CanUndo Then
            RichTextBox1.Undo()
        Else
        End If
    End Sub

    Private Sub FontDialog1_Apply(sender As Object, e As EventArgs) Handles FontDialog1.Apply

    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged

    End Sub

    Private Sub FindToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindToolStripMenuItem.Click
        Dim a As String
        Dim b As String
        a = InputBox("Enter text to be found")
        b = InStr(RichTextBox1.Text, a)
        If b Then
            RichTextBox1.Focus()
            RichTextBox1.SelectionStart = b - 1
            RichTextBox1.SelectionLength = Len(a)
        Else
            MsgBox("Text not found.")
        End If
    End Sub

    Private Sub LeftToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LeftToolStripMenuItem.Click
        RichTextBox1.SelectionAlignment = HorizontalAlignment.Left
        LeftToolStripMenuItem.Checked = True
        CenterToolStripMenuItem.Checked = False
        RightToolStripMenuItem.Checked = False
    End Sub

    Private Sub CenterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CenterToolStripMenuItem.Click
        RichTextBox1.SelectionAlignment = HorizontalAlignment.Center
        LeftToolStripMenuItem.Checked = False
        CenterToolStripMenuItem.Checked = True
        RightToolStripMenuItem.Checked = False
    End Sub

    Private Sub RightToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RightToolStripMenuItem.Click
        RichTextBox1.SelectionAlignment = HorizontalAlignment.Right
        LeftToolStripMenuItem.Checked = False
        CenterToolStripMenuItem.Checked = False
        RightToolStripMenuItem.Checked = True
    End Sub

    Private Sub BackgroundColorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BackgroundColorToolStripMenuItem.Click
        ColorDialog1.ShowDialog()
        RichTextBox1.BackColor = ColorDialog1.Color
    End Sub
End Class
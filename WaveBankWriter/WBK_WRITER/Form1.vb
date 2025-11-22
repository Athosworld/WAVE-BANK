
Imports System.ComponentModel
Imports System.Drawing.Text
Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic

Public Class Form1

    Private currentWbkBytes() As Byte = Nothing
    Private currentFilePath As String = ""
    Private originalTitle As String
    Dim pfc As New PrivateFontCollection()
    'Import DLL for Osaka font display of the "Wave Bank" text on the toolbar
    <DllImport("gdi32.dll")>
    Private Shared Function AddFontMemResourceEx(
        ByVal pbFont As IntPtr,
        ByVal cbFont As Integer,
        ByVal pdv As IntPtr,
        ByRef pcFonts As Integer) As IntPtr
    End Function

    Public Sub New()
        InitializeComponent()
        If LicenseManager.UsageMode = LicenseUsageMode.Designtime Then
            Panel1.Visible = False
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Panel1.Visible = False
        Panel2.Visible = False
        Try

            Dim fontData As Byte() = My.Resources.osaka
            Dim fontPtr As IntPtr = Marshal.AllocCoTaskMem(fontData.Length)
            Marshal.Copy(fontData, 0, fontPtr, fontData.Length)
            Dim cFonts As Integer = 0
            AddFontMemResourceEx(fontPtr, fontData.Length, IntPtr.Zero, cFonts)
            pfc.AddMemoryFont(fontPtr, fontData.Length)

            Dim customFont As New Font(pfc.Families(0), 15)
            Label2.Font = customFont
            Label4.Font = customFont
        Catch ex As Exception

        End Try


        'Add columns to the internal wave bank data display list
        ListView1.View = View.Details
        ListView1.Columns.Add("ID", 60)
        ListView1.Columns.Add("Size (KB)", 80)
        ListView1.Columns.Add("Type", 100)
        ListView1.FullRowSelect = True
        ListView1.MultiSelect = False

        AddHandler ListView1.DoubleClick, AddressOf ListView1_DoubleClick

        Panel1.Visible = False
        Panel2.Visible = False
        originalTitle = Me.Text
        BindingNavigator1.Enabled = False

        UpdateFileSizeDisplay()
    End Sub

    'Update file size label
    Private Sub UpdateFileSizeDisplay()
        If currentFilePath <> "" AndAlso File.Exists(currentFilePath) Then
            Dim fi As New FileInfo(currentFilePath)
            Label8.Text = $"Total Bank Size: {fi.Length / 1024.0:N2} KB ({fi.Length / (1024.0 * 1024.0):N3} MB)"
        Else
            Label8.Text = "Total Bank Size: No Bank Selected"
        End If
    End Sub

    'Refresh the list of waves on the internal wave bank data list every time wave is added or deleted
    Private Sub RefreshWavList()
        ListView1.Items.Clear()
        If currentWbkBytes Is Nothing OrElse currentWbkBytes.Length < 16 Then Return

        Dim bank = currentWbkBytes
        Dim i As Integer = 0

        While i < bank.Length - 7

            If bank(i) = &H5 AndAlso bank(i + 1) = &H40 AndAlso bank(i + 3) = &H1 Then
                Dim typeByte = bank(i + 2)
                Dim a = bank(i + 4)
                Dim b = bank(i + 5)
                Dim c = bank(i + 6)


                If a >= &H30 AndAlso a <= &H39 AndAlso b >= &H30 AndAlso b <= &H39 AndAlso c >= &H30 AndAlso c <= &H39 Then
                    Dim id = ChrW(a) & ChrW(b) & ChrW(c)


                    Dim j = i + 7
                    While j < bank.Length - 7
                        If bank(j) = &H5 AndAlso bank(j + 1) = &H40 AndAlso bank(j + 2) = &H95 AndAlso bank(j + 3) = &H1 Then
                            Dim ea = bank(j + 4), eb = bank(j + 5), ec = bank(j + 6)
                            If ea = a AndAlso eb = b AndAlso ec = c Then

                                Dim wavStart = i + 7
                                Dim wavLen = j - wavStart
                                Dim sizeKB = wavLen / 1024.0
                                Dim typeName = If(typeByte = &H98, "Sound Effect", If(typeByte = &H99, "Music Stream", "Unknown"))

                                Dim item As New ListViewItem(id)
                                item.SubItems.Add(sizeKB.ToString("N1"))
                                item.SubItems.Add(typeName)
                                ListView1.Items.Add(item)

                                Exit While
                            End If
                        End If
                        j += 1
                    End While
                End If
            End If
            i += 1
        End While
    End Sub



    'Extract wave for playback based on requested ID
    Private Function ExtractWavByID(id As String) As String
        If currentWbkBytes Is Nothing Then Return String.Empty
        Dim bank = currentWbkBytes
        Dim a = AscW(id(0)), b = AscW(id(1)), c = AscW(id(2))
        Dim i As Integer = 0

        While i < bank.Length - 12

            If bank(i) = &H5 AndAlso bank(i + 1) = &H40 AndAlso bank(i + 3) = &H1 Then
                Dim sa = bank(i + 4), sb = bank(i + 5), sc = bank(i + 6)
                If sa = a AndAlso sb = b AndAlso sc = c Then

                    Dim j As Integer = i + 10
                    While j < bank.Length - 6
                        If bank(j) = &H5 AndAlso bank(j + 1) = &H40 AndAlso bank(j + 2) = &H95 AndAlso bank(j + 3) = &H1 Then
                            Dim ea = bank(j + 4), eb = bank(j + 5), ec = bank(j + 6)
                            If ea = sa AndAlso eb = sb AndAlso ec = sc Then

                                Dim wavStart = i + 10
                                Dim wavLen = j - wavStart
                                If wavLen <= 0 Then Return String.Empty
                                Dim outPath = Path.Combine(Path.GetTempPath(), $"wbk_extract_{id}.wav")
                                Dim outBytes(wavLen - 1) As Byte
                                Array.Copy(bank, wavStart, outBytes, 0, wavLen)
                                File.WriteAllBytes(outPath, outBytes)
                                Return outPath
                            End If
                        End If
                        j += 1
                    End While
                End If
            End If
            i += 1
        End While

        Return String.Empty
    End Function





    'UI functions from now on
    Private Sub ListView1_DoubleClick(sender As Object, e As EventArgs)
        If ListView1.SelectedItems.Count = 0 Then Return
        Dim id = ListView1.SelectedItems(0).Text
        Dim tmp = ExtractWavByID(id)
        If tmp = "" Then
            MessageBox.Show("Cannot extract WAV.")
            Return
        End If
        My.Computer.Audio.Play(tmp, AudioPlayMode.Background)
    End Sub


    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        Try
            Dim newFile = Path.Combine(Application.StartupPath, "NewBank.wbk")
            Dim header() As Byte = {&H57, &H41, &H56, &H45, &H42, &H4E, &H4B, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
            File.WriteAllBytes(newFile, header)
            currentFilePath = newFile
            currentWbkBytes = CType(header.Clone(), Byte())
            Me.Text = $"{originalTitle} - ({Path.GetFileName(currentFilePath)})"
            BindingNavigator1.Enabled = True
            RefreshWavList()
            UpdateFileSizeDisplay()
            MessageBox.Show("Created: " & newFile)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub


    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        Using dlg As New OpenFileDialog()
            dlg.Title = "Open WBK"
            dlg.Filter = "WBK Files (*.wbk)|*.wbk"
            If dlg.ShowDialog() = DialogResult.OK Then
                If File.Exists(dlg.FileName) Then
                    currentFilePath = dlg.FileName
                    currentWbkBytes = File.ReadAllBytes(currentFilePath)
                    Me.Text = $"{originalTitle} - ({Path.GetFileName(currentFilePath)})"
                    BindingNavigator1.Enabled = True
                    RefreshWavList()
                    UpdateFileSizeDisplay()
                    MessageBox.Show("Opened: " & dlg.FileName)
                End If
            End If
        End Using
    End Sub


    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        If currentFilePath = "" OrElse currentWbkBytes Is Nothing Then
            MessageBox.Show("No WBK loaded.")
            Return
        End If
        File.WriteAllBytes(currentFilePath, currentWbkBytes)
        UpdateFileSizeDisplay()
        MessageBox.Show("Saved: " & currentFilePath)
    End Sub


    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        If currentWbkBytes Is Nothing Then Return
        Using sfd As New SaveFileDialog()
            sfd.Filter = "WBK File (*.wbk)|*.wbk"
            If sfd.ShowDialog() = DialogResult.OK Then
                File.WriteAllBytes(sfd.FileName, currentWbkBytes)
                currentFilePath = sfd.FileName
                Me.Text = $"{originalTitle} - ({Path.GetFileName(currentFilePath)})"
                UpdateFileSizeDisplay()
                MessageBox.Show("Saved as: " & sfd.FileName)
            End If
        End Using
    End Sub


    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        currentFilePath = ""
        currentWbkBytes = Nothing
        ListView1.Items.Clear()
        Me.Text = originalTitle
        BindingNavigator1.Enabled = False
        UpdateFileSizeDisplay()
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If currentWbkBytes Is Nothing Then Return
        Using dlg As New OpenFileDialog()
            dlg.Filter = "WAV Files (*.wav)|*.wav"
            If dlg.ShowDialog() <> DialogResult.OK Then Return
            Dim wavBytes = File.ReadAllBytes(dlg.FileName)


            Dim numbers As New List(Of Integer)
            Dim i As Integer = 0
            While i < currentWbkBytes.Length - 9
                If currentWbkBytes(i) = &H5 AndAlso currentWbkBytes(i + 1) = &H40 AndAlso currentWbkBytes(i + 3) = &H1 Then
                    Dim n1 = currentWbkBytes(i + 4), n2 = currentWbkBytes(i + 5), n3 = currentWbkBytes(i + 6)
                    If n1 >= &H30 AndAlso n1 <= &H39 AndAlso n2 >= &H30 AndAlso n2 <= &H39 AndAlso n3 >= &H30 AndAlso n3 <= &H39 Then
                        Dim val As Integer
                        If Integer.TryParse($"{ChrW(n1)}{ChrW(n2)}{ChrW(n3)}", val) Then numbers.Add(val)
                    End If
                End If
                i += 1
            End While
            Dim nextID = If(numbers.Count > 0, numbers.Max() + 1, 1)
            Dim idStr = nextID.ToString("000")
            Dim nA = AscW(idStr(0))
            Dim nB = AscW(idStr(1))
            Dim nC = AscW(idStr(2))
            Dim typeByte = If(RadioButton1.Checked, &H98, &H99)
            Dim header() As Byte = {&H5, &H40, typeByte, &H1, nA, nB, nC, 0, 0, 0}
            Dim endTag() As Byte = {&H5, &H40, &H95, &H1, nA, nB, nC, 0, 0, 0}

            Dim newBytes(currentWbkBytes.Length + header.Length + wavBytes.Length + endTag.Length - 1) As Byte
            Array.Copy(currentWbkBytes, 0, newBytes, 0, currentWbkBytes.Length)
            Array.Copy(header, 0, newBytes, currentWbkBytes.Length, header.Length)
            Array.Copy(wavBytes, 0, newBytes, currentWbkBytes.Length + header.Length, wavBytes.Length)
            Array.Copy(endTag, 0, newBytes, currentWbkBytes.Length + header.Length + wavBytes.Length, endTag.Length)
            currentWbkBytes = newBytes


            File.WriteAllBytes(currentFilePath, currentWbkBytes)


            RefreshWavList()
            UpdateFileSizeDisplay()
            MessageBox.Show("Added WAV ID " & idStr)
        End Using
    End Sub



    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click

        If ListView1.SelectedItems.Count = 0 Then
            MessageBox.Show("Select a WAV to delete.")
            Return
        End If

        Dim id = ListView1.SelectedItems(0).Text
        Dim a = AscW(id(0)), b = AscW(id(1)), c = AscW(id(2))
        Dim bank = currentWbkBytes

        Dim found As Boolean = False
        Dim i As Integer = 0

        While i < bank.Length - 12

            If bank(i) = &H5 AndAlso bank(i + 1) = &H40 AndAlso bank(i + 3) = &H1 Then
                Dim sa = bank(i + 4), sb = bank(i + 5), sc = bank(i + 6)
                If sa = a AndAlso sb = b AndAlso sc = c Then

                    Dim j = i + 10
                    While j < bank.Length - 6
                        If bank(j) = &H5 AndAlso bank(j + 1) = &H40 AndAlso bank(j + 2) = &H95 AndAlso bank(j + 3) = &H1 Then
                            Dim ea = bank(j + 4), eb = bank(j + 5), ec = bank(j + 6)
                            If ea = sa AndAlso eb = sb AndAlso ec = sc Then

                                Dim lenRemove = (j + 10) - i

                                Dim newBank(bank.Length - lenRemove - 1) As Byte
                                Array.Copy(bank, 0, newBank, 0, i)
                                Array.Copy(bank, j + 10, newBank, i, bank.Length - (j + 10))
                                currentWbkBytes = newBank
                                File.WriteAllBytes(currentFilePath, currentWbkBytes)
                                found = True
                                Exit While
                            End If
                        End If
                        j += 1
                    End While
                End If
            End If
            If found Then Exit While
            i += 1
        End While

        If found Then
            MessageBox.Show("Deleted WAV " & id)
            RefreshWavList()
            UpdateFileSizeDisplay()
        Else
            MessageBox.Show("Could not find WAV " & id & " to delete.")
        End If
    End Sub


    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        If ListView1.SelectedItems.Count = 0 Then Return
        Dim id = ListView1.SelectedItems(0).Text
        Dim tmp = ExtractWavByID(id)
        If tmp = "" Then
            MessageBox.Show("Cannot extract WAV.")
            Return
        End If
        My.Computer.Audio.Play(tmp, AudioPlayMode.Background)
    End Sub


    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        My.Computer.Audio.Stop()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged

    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        Panel2.Visible = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Panel1.Visible = False
        Panel2.Visible = False
    End Sub
End Class

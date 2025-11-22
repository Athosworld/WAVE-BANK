Imports System.IO

Module WBKUnpack

    Sub Main()
        Dim args = Environment.GetCommandLineArgs()

        If args.Length < 2 Then
            ShowUsage()
            Return
        End If

        Dim inputPath As String = args(1)
        Dim outputFolder As String = ""


        If args.Length >= 3 Then
            outputFolder = args(2)
        Else
            Dim inputDir = Path.GetDirectoryName(Path.GetFullPath(inputPath))
            Dim inputNameNoExt = Path.GetFileNameWithoutExtension(inputPath)
            outputFolder = Path.Combine(inputDir, inputNameNoExt & "_unarchived")
        End If


        If Not File.Exists(inputPath) Then
            Console.WriteLine("Error: input bank not found: " & inputPath)
            Return
        End If


        Try
            If Not Directory.Exists(outputFolder) Then Directory.CreateDirectory(outputFolder)
        Catch ex As Exception
            Console.WriteLine("Error creating output folder: " & ex.Message)
            Return
        End Try

        Console.WriteLine("Wave Bank Unarchiver")
        Console.WriteLine("Wave Bank: " & inputPath)
        Console.WriteLine("Unarchived output folder: " & outputFolder)
        Console.WriteLine()

        Try
            ExtractFromWBK(inputPath, outputFolder)
        Catch ex As Exception
            Console.WriteLine("Unhandled error: " & ex.Message)
        End Try
    End Sub

    Private Sub ShowUsage()
        Console.WriteLine("Wave Bank Unarchiver Tool")
        Console.WriteLine("Programmed by Athos García (2025)")
        Console.WriteLine("*------------------------------------------------------*")
        Console.WriteLine("Usage:")
        Console.WriteLine("  WBKUNARCHIVE.exe <bank_filename.wbk> [output_folder]")
        Console.WriteLine()
        Console.WriteLine("If no output folder is provided, a folder named <bank_filename>_unarchived will be created next to the input bank.")
    End Sub

    Private Sub ExtractFromWBK(wbkPath As String, outputFolder As String)
        Dim bank() As Byte = File.ReadAllBytes(wbkPath)
        Dim len As Integer = bank.Length
        Dim i As Integer = 0
        Dim extracted As Integer = 0

        While i <= len - 10
            If bank(i) = &H5 AndAlso bank(i + 1) = &H40 AndAlso bank(i + 3) = &H1 Then
                Dim typeByte As Byte = bank(i + 2)
                Dim a As Byte = bank(i + 4)
                Dim b As Byte = bank(i + 5)
                Dim c As Byte = bank(i + 6)

                If a >= &H30 AndAlso a <= &H39 AndAlso
                   b >= &H30 AndAlso b <= &H39 AndAlso
                   c >= &H30 AndAlso c <= &H39 Then

                    Dim idStr = ChrW(a) & ChrW(b) & ChrW(c)
                    Dim dataStart As Integer = i + 10

                    Dim j As Integer = dataStart
                    Dim foundEnd As Boolean = False
                    While j <= len - 10
                        If bank(j) = &H5 AndAlso bank(j + 1) = &H40 AndAlso bank(j + 2) = &H95 AndAlso bank(j + 3) = &H1 Then
                            If bank(j + 4) = a AndAlso bank(j + 5) = b AndAlso bank(j + 6) = c Then
                                Dim dataLen As Integer = j - dataStart
                                If dataLen > 0 Then
                                    Dim outBytes(dataLen - 1) As Byte
                                    Array.Copy(bank, dataStart, outBytes, 0, dataLen)

                                    Dim outName As String = idStr & ".wav"
                                    Dim outPath As String = Path.Combine(outputFolder, outName)

                                    Dim k As Integer = 1
                                    While File.Exists(outPath)
                                        outPath = Path.Combine(outputFolder, $"{idStr}_{k}.wav")
                                        k += 1
                                    End While

                                    File.WriteAllBytes(outPath, outBytes)
                                    Console.WriteLine($"Unarchived: {Path.GetFileName(outPath)}  (type=0x{typeByte:X2}, bytes={dataLen})")
                                    extracted += 1
                                End If
                                foundEnd = True
                                Exit While
                            End If
                        End If
                        j += 1
                    End While

                    If foundEnd Then
                        i = j + 10
                        Continue While
                    Else
                        Console.WriteLine($"Warning: start tag at offset {i} for ID {idStr} has no matching end tag. Wave will not be extracted.")
                    End If
                End If
            End If

            i += 1
        End While

        Console.WriteLine()
        Console.WriteLine($"Finished. Extracted {extracted} WAV(s).")
    End Sub

End Module

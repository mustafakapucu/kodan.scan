﻿Imports System.Drawing
Imports System.Threading

Public Module Module1
    Dim m_szPaperSource As String = "3"                 ' 0 auto, 1 adf front, 2 adf rear, 3 adf duplex, 4 flatbed
    'Dim m_szShowScannerUI As String = "0"               ' 0 off, 1 on
    Dim m_szFileName As String = "img"
    Dim m_szFileNumber As String = "0"                  ' 0 to 999
    Dim m_szFilePathName As String = ""
    Dim m_szFileFolderName As String = "\twain"
    Dim currentDirectoryPath As String = ""
    Dim m_szLanguage As String = "0"                    ' 0 English, 1 Chinese Simplified, 2 Chinese Traditional
    Dim m_szScanAs As String = "2"                      ' 0 bw, 1 gray, 2 color
    'Dim m_szDpiResolution As String = "1200"             ' 100, 150, 200, 240, 250, 300, 400, 500, 600, 1200
    'Dim m_szDocumentType As String = "0"                ' 0 photo, 1 textwithgraphics, 2 textwithphoto, 3 text
    Dim m_szFileType As String = "0"                    ' 0 tiff, 4 jpeg
    Dim m_szCompressionType As String = "0"             ' 0 none, 5 g4, 6 jpeg
    'Dim m_szJpegQuality As String = "40"                ' 40 draft, 50 good, 80 better, 90 best, 100 superior
    'Dim m_szSharpen As String = "0"                     ' 0 none, 1 normal, 2 high, 3 exaggerated
    'Dim m_szOrthogonalRotation As String = "0"          ' 0 none, 1 auto, 2 90, 3 180, 4 270, 5 auto 90, 6 auto 180, 7 auto 270 
    Dim m_szImageRotation As String = "180"               ' 0 0, 1 auto, 2 none, 90 90, 180 180 , 270 270, 360 360
    'Dim m_szBlankImageDeletion As String = "2"          ' 1 none, 2 content
    'Dim m_szBlankImageDeletionPercent As String = "0"   ' 0 to 100
    Dim m_szScanner As String = "KODAK Scanner: i2000"
    ' Dim m_szScannerProfile As String = "1"              ' 1 is typically "Default" profile
    Dim m_szOnePage As String = "0"                     ' 0 scan mulitple pages, 1 scan 1 page
    Dim filenumber As Integer
    Dim hWnd As IntPtr = 1
    Dim imageExtension = ".png"
    Dim imageNameprefix As String = "ecm00000"
    Dim imageFormat As Imaging.ImageFormat = Imaging.ImageFormat.Png
    Dim processName As String = "kodak.scan"

    Dim KODAKSCANSDK As KODAKSCANSDK.Program
    Private WithEvents MyScanEvent As KODAKSCANSDK.Program


    Public Sub Main()

        Try
            KODAKSCANSDK = New KODAKSCANSDK.Program
            All()
            Console.Write(1)
            Environment.Exit(1)


        Catch ex As Exception
            Console.Write(0)
            Environment.Exit(0)
        End Try

    End Sub

    Public Function All()

        currentDirectoryPath = IO.Directory.GetCurrentDirectory()

        'set baseFileFullPath
        m_szFilePathName = currentDirectoryPath.Substring(0, currentDirectoryPath.LastIndexOf("\")) + m_szFileFolderName

        KillOldProcess()
        RemoveAllFiles()
        SelectScanner()
        OpenScanner()
        StartScan()

        '    My.Computer.FileSystem.CopyFile(
        '"C:\Users\mustafa.kapucu\Desktop\New folder\examples\img000001.tif",
        '"C:\Twain\img000001.tif")
        '    My.Computer.FileSystem.CopyFile(
        '"C:\Users\mustafa.kapucu\Desktop\New folder\examples\img000002.tif",
        '"C:\Twain\img000002.tif")
        '    My.Computer.FileSystem.CopyFile(
        '"C:\Users\mustafa.kapucu\Desktop\New folder\examples\img000003.tif",
        '"C:\Twain\img000003.tif")
        '    My.Computer.FileSystem.CopyFile(
        '"C:\Users\mustafa.kapucu\Desktop\New folder\examples\img000004.tif",
        '"C:\Twain\img000004.tif")
        '    My.Computer.FileSystem.CopyFile(
        '"C:\Users\mustafa.kapucu\Desktop\New folder\examples\img000005.tif",
        '"C:\Twain\img000005.tif")
        '    My.Computer.FileSystem.CopyFile(
        '"C:\Users\mustafa.kapucu\Desktop\New folder\examples\img000006.tif",
        '"C:\Twain\img000006.tif")

        If CheckScannerStatus() Then
            CombineImages()
        End If

    End Function

    Private Function KillOldProcess()
        Dim counter As Integer = 0
        Dim processlist As List(Of Process) = Process.GetProcesses().ToList().Where(Function(p) p.ProcessName.Equals(processName)).ToList().OrderBy(Function(o) o.TotalProcessorTime).ToList()

        If processlist IsNot Nothing And processlist.Count > 1 Then
            For Each item As Process In processlist
                If counter > 0 Then
                    item.Kill()
                End If
                counter = counter + 1
            Next
        End If
    End Function

    Private Function CheckScannerStatus() As Boolean
        Try
            Dim di As New IO.DirectoryInfo(m_szFilePathName)
            Dim startCount As Integer
            startCount = 0

            While True
                startCount = di.GetFiles("*").Length
                Thread.Sleep(2500)

                If di.GetFiles("*").Length > startCount Or startCount = 0 Then
                    Continue While
                Else
                    Exit While
                End If

            End While
            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function

    Private Function OpenScanner() As Integer

        Dim ExitCode As Integer

        ExitCode = KODAKSCANSDK.SetScanner(m_szScanner)
        If ExitCode <> 0 Then
            Throw New Exception()
        End If

        ExitCode = KODAKSCANSDK.SetFileNumber(m_szFileNumber)
        If ExitCode <> 0 Then
            Throw New Exception()
        End If

        filenumber = Convert.ToInt32(m_szFileNumber)

        ExitCode = KODAKSCANSDK.SetFileName(m_szFileName)
        If ExitCode <> 0 Then
            Throw New Exception()
        End If

        ExitCode = KODAKSCANSDK.SetFilePathName(m_szFilePathName)
        If ExitCode <> 0 Then
            Throw New Exception()
        End If

        ExitCode = KODAKSCANSDK.OpenScanner(hWnd)
        If ExitCode <> 0 Then
            Throw New Exception()
        End If

        ExitCode = KODAKSCANSDK.SetImageRotation(m_szImageRotation)
        If ExitCode <> 0 Then
            Throw New Exception()
        End If

        ExitCode = KODAKSCANSDK.SetScanAs(m_szScanAs)
        If ExitCode <> 0 Then
            Throw New Exception()
        End If

        'ExitCode = KODAKSCANSDK.SetDPIResolution(m_szDpiResolution)
        'If ExitCode <> 0 Then
        '    Throw New Exception()
        'End If

        ExitCode = KODAKSCANSDK.SetPaperSource(m_szPaperSource)
        If ExitCode <> 0 Then
            Throw New Exception()
        End If

        ExitCode = KODAKSCANSDK.SetFileType(m_szFileType)
        If ExitCode <> 0 Then
            Throw New Exception()
        End If


        ExitCode = KODAKSCANSDK.SetCompressionType(m_szCompressionType)
        If ExitCode <> 0 Then
            Throw New Exception()
        End If

        'ExitCode = KODAKSCANSDK.SetJPEGQuality(m_szJpegQuality)
        'If ExitCode <> 0 Then
        '    Throw New Exception()
        'End If

    End Function

    Private Function SelectScanner()
        Dim ExitCode As Integer

        ExitCode = KODAKSCANSDK.SetLanguage(m_szLanguage)

        If ExitCode <> 0 Then
            Throw New Exception()
        End If

        ExitCode = KODAKSCANSDK.Init(hWnd)

        If ExitCode <> 0 Then
            Throw New Exception()
        End If
        KODAKSCANSDK.SelectScanner()

    End Function

    Private Function StartScan() As Integer
        Return KODAKSCANSDK.StartScan(m_szOnePage, hWnd)
    End Function

    Function CombineImages()
        Dim img1 As Bitmap
        Dim img2 As Bitmap
        Dim counter As Integer = 1

        Dim strFileSize As String = ""
        Dim di As New IO.DirectoryInfo(m_szFilePathName)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*")

        For index = 0 To aryFi.Length - 1

            If index > 0 And (index + 1) Mod 2 = 0 Then
                img2 = New Bitmap(aryFi.ElementAt(index - 1).FullName, True)
                img1 = New Bitmap(aryFi.ElementAt(index).FullName, True)

                Dim bmp As New Bitmap(img1.Width + img2.Width, Math.Max(img1.Height, img2.Height))
                Dim g As Graphics = Graphics.FromImage(bmp)

                g.DrawImage(img1, 0, 0, img1.Width, img1.Height)
                g.DrawImage(img2, img1.Width, 0, img2.Width, img2.Height)


                'Dim bmp As New Bitmap(Math.Max(img1.Width, img2.Width), img1.Height + img2.Height)
                'Dim g As Graphics = Graphics.FromImage(bmp)

                'g.DrawImage(img1, 0, 0, img1.Width, img1.Height)
                'g.DrawImage(img2, 0, img1.Height, img2.Width, img2.Width)


                g.Dispose()
                bmp.Save(m_szFilePathName + "\" + imageNameprefix + counter.ToString() + imageExtension, imageFormat)
                counter = counter + 1
            End If

        Next

#Disable Warning BC42105 ' Function doesn't return a value on all code paths
    End Function
#Enable Warning BC42105 ' Function doesn't return a value on all code paths

    Private Function RemoveAllFiles() As Boolean
        Try
            Dim di As New IO.DirectoryInfo(m_szFilePathName)
            Dim aryFi As IO.FileInfo() = di.GetFiles("*")

            If aryFi.Length > 0 Then
                For Each item In aryFi
                    If System.IO.File.Exists(item.FullName) = True Then
                        System.IO.File.Delete(item.FullName)
                    End If
                Next
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

End Module

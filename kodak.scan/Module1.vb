Module Module1
    Dim m_szPaperSource As String = "3"                 ' 0 auto, 1 adf front, 2 adf rear, 3 adf duplex, 4 flatbed
    Dim m_szShowScannerUI As String = "0"               ' 0 off, 1 on
    Dim m_szFileName As String = "img"
    Dim m_szFileNumber As String = "0"                  ' 0 to 999
    Dim m_szFilePathName As String = "C:\twain"
    Dim m_szLanguage As String = "0"                    ' 0 English, 1 Chinese Simplified, 2 Chinese Traditional
    Dim m_szScanAs As String = "2"                      ' 0 bw, 1 gray, 2 color
    Dim m_szDpiResolution As String = "1200"             ' 100, 150, 200, 240, 250, 300, 400, 500, 600, 1200
    Dim m_szDocumentType As String = "1"                ' 0 photo, 1 textwithgraphics, 2 textwithphoto, 3 text
    Dim m_szFileType As String = "0"                    ' 0 tiff, 4 jpeg
    Dim m_szCompressionType As String = "0"             ' 0 none, 5 g4, 6 jpeg
    Dim m_szJpegQuality As String = "100"                ' 40 draft, 50 good, 80 better, 90 best, 100 superior
    Dim m_szSharpen As String = "0"                     ' 0 none, 1 normal, 2 high, 3 exaggerated
    Dim m_szOrthogonalRotation As String = "0"          ' 0 none, 1 auto, 2 90, 3 180, 4 270, 5 auto 90, 6 auto 180, 7 auto 270 
    Dim m_szImageRotation As String = "0"               ' 0 0, 1 auto, 2 none, 90 90, 180 180 , 270 270, 360 360
    Dim m_szBlankImageDeletion As String = "2"          ' 1 none, 2 content
    Dim m_szBlankImageDeletionPercent As String = "0"   ' 0 to 100
    Dim m_szScanner As String = "KODAK Scanner: i2000"
    Dim m_szScannerProfile As String = "1"              ' 1 is typically "Default" profile
    Dim m_szOnePage As String = "0"                     ' 0 scan mulitple pages, 1 scan 1 page
    Dim filenumber As Integer
    Dim hWnd As IntPtr = 1

    Dim KODAKSCANSDK As KODAKSCANSDK.Program
    Private WithEvents MyScanEvent As KODAKSCANSDK.Program

    Sub Main()
        KODAKSCANSDK = New KODAKSCANSDK.Program
        SelectScanner()
        OpenScanner()
        StartScan()
    End Sub

    Private Sub OpenScanner()
        Dim ExitCode As Integer

        ExitCode = KODAKSCANSDK.SetScanner(m_szScanner)
        If ExitCode <> 0 Then
            MsgBox("An error occurred: SetScanner" & vbCrLf & ExitCode.ToString("X"))
        End If

        ExitCode = KODAKSCANSDK.SetFileNumber(m_szFileNumber)
        If ExitCode <> 0 Then
            MsgBox("An error occurred: SetFileNumber" & vbCrLf & ExitCode.ToString("X"))
        End If
        filenumber = Convert.ToInt32(m_szFileNumber)

        ExitCode = KODAKSCANSDK.SetFileName(m_szFileName)
        If ExitCode <> 0 Then
            MsgBox("An error occurred: SetFileName" & vbCrLf & ExitCode.ToString("X"))
        End If

        ExitCode = KODAKSCANSDK.SetFilePathName(m_szFilePathName)
        If ExitCode <> 0 Then
            MsgBox("An error occurred: SetFilePathName" & vbCrLf & ExitCode.ToString("X"))
        End If

        ExitCode = KODAKSCANSDK.OpenScanner(hWnd)
        If ExitCode <> 0 Then
            MsgBox("An error occurred: OpenScanner" & vbCrLf & ExitCode.ToString("X"))
        End If

        ExitCode = KODAKSCANSDK.SetScannerProfile(m_szScannerProfile)
        If ExitCode <> 0 Then
            MsgBox("An error occurred: SetScannerProfile" & vbCrLf & ExitCode.ToString("X"))
        End If

    End Sub

    Private Sub SelectScanner()
        Dim ExitCode As Integer
        Dim strScanners As String

        ExitCode = KODAKSCANSDK.SetLanguage(m_szLanguage)
        If ExitCode <> 0 Then
            MsgBox("An error occurred: SetLanguage" & vbCrLf & ExitCode.ToString("X"))
        End If

        ExitCode = KODAKSCANSDK.Init(hWnd)
        If ExitCode <> 0 Then
            MsgBox("An error occurred: Init" & vbCrLf & ExitCode.ToString("X"))
        End If

        strScanners = KODAKSCANSDK.SelectScanner()

    End Sub

    Private Sub StartScan()
        Dim ExitCode As Integer

        ExitCode = KODAKSCANSDK.StartScan(m_szOnePage, hWnd)
        If ExitCode <> 0 Then
            MsgBox("An error occurred: StartScan" & vbCrLf & ExitCode.ToString("X"))

        End If

    End Sub

End Module

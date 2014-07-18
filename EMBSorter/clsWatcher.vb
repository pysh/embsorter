'
' Сделано в SharpDevelop.
' Пользователь: pnosov
' Дата: 11.12.2013
' Время: 17:50
' 
' Для изменения этого шаблона используйте Сервис | Настройка | Кодирование | Правка стандартных заголовков.
'
Imports System.Linq
Imports System.IO
Imports System.Threading
Imports System.Timers
Imports System.Collections


Public Class clsWatcher
    Public WithEvents _fsw As New System.IO.FileSystemWatcher
    Private WithEvents _timerWaitNext As New System.Timers.Timer
    Private WithEvents _timerTick As New System.Timers.Timer

    Private _FolderName As String
    Private _FileMask As String
    Private _sec As Integer
    Private _strNewFilename As String
    
    Private Shared iniFile As New ini(IO.Path.ChangeExtension(Application.ExecutablePath,".ini"))

    Public Class ShowPopupEventArgs
        Inherits EventArgs
        Implements IDisposable
        Public Property strMessage As String = ""
        Public Property strHeader As String = ""
        Public Property intShowInterval As Integer = 0
        Public Property tipIcon As ToolTipIcon = ToolTipIcon.Info
        Public Sub Dispose() Implements IDisposable.Dispose
        End Sub
    End Class
    
    Public Event ShowPopup As EventHandler(Of ShowPopupEventArgs)
    
    Protected Overridable Sub OnShowPopup(e As ShowPopupEventArgs)
        RaiseEvent ShowPopup(Me, e)
    End Sub
    
    Private Sub _ShowPopup(timeOut As Integer, tipTitle As String, tipText As String, tipIcon As ToolTipIcon)' strMessage As String, strHeader As String, intInterval As Integer,)
        Using args As New ShowPopupEventArgs
            args.intShowInterval = 2000
            args.strHeader = tipTitle
            args.strMessage = tipText
            OnShowPopup(args)
        End Using
    End Sub
    
    
    Private Class cEFile
        Public FileName As String = ""
        Public PromoName As String = ""
        Public CardQty As Integer = -1
        Public PlasticCode As String = ""
        Public PlasticID As String = ""
        Public CreationDate As Date = Date.Now
        Public Sub Clear
            FileName = ""
            PromoName = ""
            CardQty = -1
            PlasticCode = ""
            PlasticID = ""
            CreationDate = Date.Now
        End Sub
    End Class

    ' <Subs>
    Public Sub New(ByVal FolderName As String, ByVal FileMask As String, Optional intIntervalTimer As Integer = 25000)
        TimerWaitNextInterval = intIntervalTimer
        _InitializeTimers()
        _FolderName = FolderName
        _FileMask = FileMask
        _fsw.Path = _FolderName
        _fsw.Filter = _FileMask
        _fsw.IncludeSubdirectories = False
        AddHandler _fsw.Created, AddressOf _OnFileCreate
        AddHandler _fsw.Renamed, AddressOf _OnFileCreate
        _fsw.EnableRaisingEvents = True
    End Sub

    ' </Subs>

    '<Properties>
    Property TimerWaitNextInterval() As Double
        Get
            Return _timerWaitNext.Interval
        End Get
        Set(ByVal value As Double)
            _timerWaitNext.Interval = value
            Debug.WriteLine("TimerWaitNextInterval_Set")
        End Set
    End Property

    Property TimerTickInterval() As Double
        Get
            Return _timerTick.Interval
        End Get
        Set(ByVal value As Double)
            _timerTick.Interval = value
        End Set
    End Property

    Property FolderName() As String
        Get
            Return _FolderName
        End Get
        Set(ByVal value As String)
            _FolderName = value
        End Set
    End Property
    
    Property FileMask() As String
        Get
            Return _FileMask
        End Get
        Set(ByVal value As String)
            _FileMask = value
        End Set
    End Property
    
    
    '</Properties>
    
    '<Main>

    
    Private Sub _timerWaitNextElapsed()
        _timerWaitNext.Stop
        _timerTick.Stop
        Debug.WriteLine("timerWaitNextElapsed")
        _MoveFiles
    End Sub
    
    Private Sub _OnFileCreate(sender As Object, e As System.IO.FileSystemEventArgs)
'        notifyIcon.ShowBalloonTip(20000, Now.ToString & ". Обнаружен новый файл", e.Name, ToolTipIcon.Info)
        _strNewFilename = e.FullPath
        Using args As New ShowPopupEventArgs
            args.intShowInterval = 1000
            args.strMessage = "Обнаружен новый файл: " & IO.Path.GetFileName(e.Name)
            args.strHeader = String.Format("{0} - {1}", Now(), IO.Path.GetDirectoryName(Me._FolderName))
            OnShowPopup(args)
        End Using
        
        _timerWaitNext.Stop
        _timerTick.Stop
        Debug.WriteLine("_OnFileCreate")
        ' _InitializeTimers()
        _sec = 0
        _timerWaitNext.Start 
        _timerTick.Start
        ' MainTimerStart()
    End Sub        
    
    Private Sub _InitializeTimers()
        _sec = 0
        _timerTick.Stop
        _timerWaitNext.Stop
        _timerWaitNext.Interval = TimerWaitNextInterval
        _timerTick.Interval = 1000
        AddHandler _timerWaitNext.Elapsed, AddressOf _timerWaitNextElapsed 
        AddHandler _timerTick.Elapsed, AddressOf _timerTickElapsed
    End Sub
    
    
    Private Sub _timerTickElapsed()
        _sec = _sec + 1
        Debug.WriteLine (_timerWaitNext.Interval/1000 - _sec, "TimerTickElapsed")
        Using args As New ShowPopupEventArgs
            args.intShowInterval = 30000
            args.strMessage = "Обнаружен новый файл: " & IO.Path.GetFileName(_strNewFilename) & vbCrLf & _
                              "Ожидание... " & (_timerWaitNext.Interval/1000 - _sec).ToString & " сек."
            args.strHeader = String.Format("{0} - {1}", Now(), IO.Path.GetDirectoryName(Me._FolderName))
            args.tipIcon = ToolTipIcon.None
            OnShowPopup(args)
        End Using
    End Sub

    Public Sub MoveFiles()
        _timerWaitNextElapsed()
    End Sub
    
    Function ReplaceInvalidFileNameChars(strIn As String) As String
        strIn.Replace("/", "_")
        strIn.Replace("\", "_")
        strIn.Replace("|", "|")
        strIn.Replace(":", ";")
        strIn.Replace("*", "#")
        strIn.Replace("?", "$")
        strIn.Replace("""", "'")
        strIn.Replace("<", "{")
        strIn.Replace(">", "}")
        Return strIn
    End Function
    
#Region "Работа с файлами"
    Private Sub _MoveFiles()
        Dim EFile As cEFile
        Dim strOutPath As String
        Dim strOutFile As String
        Dim dt As String = Date.Now.ToString("HHmmssfff")
        Dim EFiles As List(Of cEFile) = GetEFiles(_FolderName, _FileMask)
        For Each EFile In EFiles
            'FileListView.Items.Add(f.FullName).SubItems.Add(String.Format("{0}.{1:D3}", f.CreationTime,f.CreationTime.Millisecond))
            strOutPath = Path.Combine(_FolderName, _
                String.Format("{0:yyyyMMdd}", Date.Now), _
                String.Format("{0}_{1}__{2}_{3}@{4}", EFile.PlasticID, GetEFilesQty(EFiles), EFile.PlasticCode, ReplaceInvalidFileNameChars(EFile.PromoName), dt))
            ' EFile = GetPromoNameOfFile(f.FullName)
            strOutFile = IO.Path.Combine(strOutPath, IO.path.GetFileName(EFile.FileName))
            If Not IO.Directory.Exists(strOutPath) Then CreateDirectoryRecursive(strOutPath)
            IO.File.Move(EFile.FileName, strOutFile)
            Debug.WriteLine(strOutFile, "Move")
        Next
        Using args As New ShowPopupEventArgs
            args.intShowInterval = 0
            args.strMessage = "Перемещено файлов: " & EFiles.Count.ToString
            args.strHeader = String.Format("{0} - {1}", Now(), IO.Path.GetDirectoryName(Me._FolderName))
            OnShowPopup(args)
        End Using
    End Sub
    
    Private Function GetEFilesQty(EFiles As List(Of cEFile)) As Integer
        Dim Qty As Integer = 0
        For Each EFile As cEFile In EFiles
            Qty = Qty + EFile.CardQty
        Next
        Return Qty
    End Function
    
    Private Sub CreateDirectoryRecursive(strDirectoryName As String)
        Dim pfn As String = IO.Path.GetDirectoryName(strDirectoryName)
        If Not IO.Directory.Exists(pfn) Then CreateDirectoryRecursive(pfn)
        IO.Directory.CreateDirectory(strDirectoryName)
    End Sub
    
    Private Function GetEFiles(strPath As String, strFileMask As String) as List(Of cEFile)
        Dim RetVal As New List(Of cEFile)
        Dim DirInfo as New IO.DirectoryInfo(strPath)
        Dim fls As IEnumerable = From fl In DirInfo.EnumerateFiles(strFileMask) _
            Order By fl.CreationTime

        For Each EFileInfo As IO.FileInfo In fls
            Dim ef As New cEFile
            ef.Clear
            ef.FileName = EFileInfo.FullName
            Try
                ' Create an instance of StreamReader to read from a file.
                ' The using statement also closes the StreamReader.
                Using sr As New System.IO.StreamReader(EFileInfo.FullName, System.Text.Encoding.GetEncoding(1251))
                    Dim line As String
                    Dim c As Integer = 0
                    ' Read and display lines from the file until the end of
                    ' the file is reached.
                    Do
                        line = sr.ReadLine()
                        If Not (line Is Nothing) Then
                            c = c + 1
                            ef.CardQty = c
                            ef.PromoName   = GetPromo(line.Substring(456,100).Trim)
                            ef.PlasticCode = line.Substring(110,5).Trim
                            ef.CreationDate= EFileInfo.CreationTime
                            ef.PlasticID   = iniFile.GetString("PLASTIC_ID", ef.PromoName,"")
                        End If
                    Loop Until line Is Nothing
                    ' f.ProductName = f.ProductName.
                End Using
            Catch e As Exception
                ' Let the user know what went wrong.
                MessageBox.Show(e.Message & vbCrLf, "The file could not be read", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            If ef.PlasticID = "" Then
                iniFile.WriteString("PLASTIC_ID", ef.PromoName, "")
'                notifyIcon.ShowBalloonTip(20000, ef.PromoName, "Неизвестный продукт", ToolTipIcon.Warning)
            End If
            RetVal.Add(ef)
        Next
        Return RetVal       
    End Function    
    

    Private Function GetPromo(strIN As String) As String
        strIN = strIN.Remove(strIN.LastIndexOf(" ")).TrimEnd
        If strIN.IndexOf(" ") > -1 Then
            strIN = strIN.Substring(strIN.IndexOf(" "))
        End If
        Return strIN.Trim
    End Function
    
#End Region




End Class

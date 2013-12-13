'
' Сделано в SharpDevelop.
' Пользователь: pnosov
' Дата: 11.12.2013
' Время: 17:50
' 
' Для изменения этого шаблона используйте Сервис | Настройка | Кодирование | Правка стандартных заголовков.
'
Public Class clsWatcher
	Private WithEvents _fsw As New System.IO.FileSystemWatcher
	Private WithEvents _timerWaitNext As New System.Timers.Timer
	Private WithEvents _timerTick As New System.Timers.Timer
	Private _FolderName As String
	Private _FileMask As String
	Private _sec As Integer

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
	Public Sub New(ByVal FolderName, ByVal FileMask)
		
		_FolderName = FolderName
		_FileMask = FileMask
	End Sub
	
	
	Public Sub Run()
		_fsw.Path = _FolderName
		_fsw.Filter = _FileMask
		_fsw.IncludeSubdirectories = False
		AddHandler _fsw.Created, AddressOf _OnFileCreate
		AddHandler _fsw.Renamed, AddressOf _OnFileCreate
		_fsw.EnableRaisingEvents = True
		'		_timerWaitNext.Start
	End Sub

	' </Subs>
	
	'<Properties>
	Property TimerWaitNextInterval() As Double
        Get
            Return _timerWaitNext.Interval
        End Get
        Set(ByVal value As Double)
            _timerWaitNext.Interval = value
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
	Private Sub New()
		AddHandler _timerWaitNext.Elapsed, AddressOf _timerWaitNextElapsed
'		AddHandler _timerTick.Elapsed, AddressOf _timerTickElapsed
		_timerWaitNext.Interval = 23000
'		_timerTick.Interval = 1000
	End Sub
	
	Private Sub _timerWaitNextElapsed()
		_timerWaitNext.Stop
'		_timerTick.Stop
		_MoveFiles
	End Sub
	
	Private Sub _OnFileCreate(sender As Object, e As System.IO.FileSystemEventArgs)
	'	notifyIcon.ShowBalloonTip(20000, Now.ToString & ". Обнаружен новый файл", e.Name, ToolTipIcon.Info)
		_timerWaitNext.Start ' MainTimerStart()
	End Sub		
	
	Private Sub _MoveFiles()
		
	End Sub
	
'	Private Sub _timerTickElapsed()
'		Debug.WriteLine ("TimerTickElapsed", _sec.ToString)
'		_sec = _sec + 1
''		notifyIcon.ShowBalloonTip(20000, "...", _
''			"Ожидание... " & FormatDateTime(TimeSerial(0,0,CInt(((timer1.Interval/1000) - _sec))), DateFormat.LongTime).ToString, _
''			ToolTipIcon.Info)
'		timerTick()
'	End Sub
'	
'	Public Sub timerTick()
'		
'	End Sub



#Region "Работа с файлами"
	Private Sub MoveFiles()
		Dim EFile As cEFile
		Dim strOutPath As String
		Dim strOutFile As String
		Dim dt As String = Date.Now.ToString("HHmmssfff")
		Dim EFiles As Collection = GetEFiles(cFilePath, "*.001")
		For Each EFile In EFiles
			'FileListView.Items.Add(f.FullName).SubItems.Add(String.Format("{0}.{1:D3}", f.CreationTime,f.CreationTime.Millisecond))
			strOutPath = Path.Combine(cFilePath, _
				String.Format("{0:yyyyMMdd}", Date.Now), _
				String.Format("{0}_{1}__{2}_{3}@{4}", EFile.PlasticID, GetEFilesQty(EFiles), EFile.PlasticCode, EFile.PromoName, dt))
			Debug.WriteLine (strOutPath)
			' EFile = GetPromoNameOfFile(f.FullName)
			strOutFile = IO.Path.Combine(strOutPath, IO.path.GetFileName(EFile.FileName))
			If Not IO.Directory.Exists(strOutPath) Then CreateDirectoryRecursive(strOutPath)
			IO.File.Move(EFile.FileName, strOutFile)
		Next		
	End Sub
	
	Private Function GetEFilesQty(EFiles As Collection) As Integer
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
	
	Private Function GetEFiles(strPath As String, strFileMask As String) as Collection
        Dim RetVal As New Collection
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
	            MessageBox.Show(e.Message, "The file could not be read", MessageBoxButtons.OK, MessageBoxIcon.Error)
	        End Try
	        If ef.PlasticID = "" Then
	        	iniFile.WriteString("PLASTIC_ID", ef.PromoName, "")
	        	notifyIcon.ShowBalloonTip(20000, ef.PromoName, "Неизвестный продукт", ToolTipIcon.Warning)
	        End If
	        RetVal.Add(ef, ef.FileName)
        Next
		Return RetVal       
	End Function	
	
#End Region




End Class

'
' Сделано в SharpDevelop.
' Пользователь: pnosov
' Дата: 22.11.2012
' Время: 12:40
' 
' Для изменения этого шаблона используйте Сервис | Настройка | Кодирование | Правка стандартных заголовков.
'
Imports System.Linq
Imports System.IO
Imports System.Threading
Imports System.Timers
Imports System.Collections

Public NotInheritable Class NotificationIcon
	Private notifyIcon As NotifyIcon
	Private notificationMenu As ContextMenu
	Private fsw As System.IO.FileSystemWatcher
	Private WithEvents timer1 As System.Timers.Timer
	Private WithEvents timer2 As System.Timers.Timer
	Private t2 As Integer = 0
	Private Const cFilePath As String = "s:\Way4\OWSWork\WAY4P\Data\CARD_PRD\Embs\"
	Private iniFile As New ini(IO.Path.ChangeExtension(Application.ExecutablePath,".ini"))
	
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
	
	Private cEFiles000 As Collection
	
	#Region "Initialize icon and menu"
	Public Sub New()
		notifyIcon = New NotifyIcon()
		notificationMenu = New ContextMenu(InitializeMenu())
		
		InitializeTimers()
		InitializeFileSystemWatcher(cFilePath,"*.001")

		AddHandler notifyIcon.DoubleClick, AddressOf IconDoubleClick
		Dim resources As New System.ComponentModel.ComponentResourceManager(GetType(NotificationIcon))
		'notifyIcon.Icon = DirectCast(resources.GetObject("$this.Icon"), Icon)
		notifyIcon.Icon = DirectCast(resources.GetObject("page_white_go"), Icon)
		notifyIcon.ContextMenu = notificationMenu
	End Sub
	
	Private Sub InitializeTimers()
		timer1 = New system.timers.Timer
		timer2 = New system.Timers.Timer
		AddHandler timer1.Elapsed, AddressOf timer1tick
		AddHandler timer2.Elapsed, AddressOf timer2tick
		timer1.Interval = 23000
		timer2.Interval = 1000
	End Sub
	
	Private Sub InitializeFileSystemWatcher(strPath As String, strFilter As String)
		fsw = New System.IO.FileSystemWatcher(strPath, strFilter)
		' fsw.NotifyFilter = IO.NotifyFilters. NotifyFilters.
		AddHandler fsw.Created, AddressOf OnFileCreate
		AddHandler fsw.Renamed, AddressOf OnFileCreate
		fsw.EnableRaisingEvents = True
	End Sub
	
	Private Function InitializeMenu() As MenuItem()
		Dim menu As MenuItem() = New MenuItem() { _
			New MenuItem("Обработать", AddressOf menuProcessClick), _
			New MenuItem("Exit", AddressOf menuExitClick)}
		Return menu
	End Function
	#End Region

	#Region "Main - Program entry point"
	''' <summary>Program entry point.</summary>
	''' <param name="args">Command Line Arguments</param>
	<STAThread> _
	Public Shared Sub Main(args As String())
		Application.EnableVisualStyles()
		Application.SetCompatibleTextRenderingDefault(False)

		Dim isFirstInstance As Boolean
		' Please use a unique name for the mutex to prevent conflicts with other programs
		Using mtx As New Mutex(True, "EMBSorter", isFirstInstance)
			If isFirstInstance Then
				Dim notificationIcon As New NotificationIcon()
				notificationIcon.notifyIcon.Visible = True
				Application.Run()
				notificationIcon.notifyIcon.Dispose()
				' The application is already running
				' TODO: Display message box or change focus to existing application instance
			Else
			End If
		End Using
		' releases the Mutex
	End Sub


	
	Private Sub MoveFiles() ' sender As Object, e As System.IO.RenamedEventArgs
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
	
	' MCS_GRN_19__EG008_GREEN_TP104@123226667
	
	
	Function GetEFilesQty(EFiles As Collection) As Integer
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
	
	Private Function GetPromo(strIN As String) As String
		strIN = strIN.Substring(InStr(strIN," ")).Trim
		Return strIN.Remove(strIN.LastIndexOf(" "))
	End Function
	

	
	
	
	#End Region

	#Region "Event Handlers"
	' Обработка таймеров
	Private Sub MainTimerStart()
		t2 = 0
		timer1.Stop
		timer2.Stop
		Debug.WriteLine ("MainTimerStart")
		timer1.Start
		timer2.Start
	End Sub	
	
	Private Sub Timer1T
		timer1.Stop()
		timer2.Stop()
		Debug.WriteLine ("Перемещение файлов")
		' MoveFiles()
		t2 = 0
	End Sub
	
	' Обработка события "Создан файл"	
	Private Sub OnFileCreate(sender As Object, e As System.IO.FileSystemEventArgs)
		notifyIcon.ShowBalloonTip(20000, Now.ToString & ". Обнаружен новый файл", e.Name, ToolTipIcon.Info)
		MainTimerStart()
	End Sub	
	
	Private Sub timer1tick(source As Object, e As ElapsedEventArgs)
		Debug.WriteLine ("Timer1Tick")
		Timer1T
	End Sub

	Private Sub timer2tick(source As Object, e As ElapsedEventArgs)
		Debug.WriteLine ("Timer2Tick")
		t2 = t2 + 1
		notifyIcon.ShowBalloonTip(20000, "...", _
			"Ожидание... " & FormatDateTime(TimeSerial(0,0,CInt(((timer1.Interval/1000)-t2))), DateFormat.LongTime).ToString, _
			ToolTipIcon.Info)
	End Sub	

	Private Sub menuProcessClick(sender As Object, e As EventArgs)
		Timer1T()
		' MessageBox.Show("About This Application")
	End Sub

	Private Sub menuExitClick(sender As Object, e As EventArgs)
		Application.[Exit]()
	End Sub

	Private Sub IconDoubleClick(sender As Object, e As EventArgs)
		MessageBox.Show("The icon was double clicked")
	End Sub
	#End Region
End Class

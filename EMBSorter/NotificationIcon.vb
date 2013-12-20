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
'	Private fsw As System.IO.FileSystemWatcher
'	Private WithEvents timer1 As System.Timers.Timer
'	Private WithEvents timer2 As System.Timers.Timer
'	Private t2 As Integer = 0
'	Private Const cFilePath As String = "c:\temp000\" 's:\Way4\OWSWork\WAY4P\Data\CARD_PRD\Embs\"
		
	
	#Region "Initialize icon and menu"
	Public Sub New()
		
		notifyIcon = New NotifyIcon()
		notificationMenu = New ContextMenu(InitializeMenu())
		Dim p0 As New clsWatcher("S:\Way4\OWSWork\WAY4P\Data\CARD_PRD\Embs\","*.001")
		Dim p1 As New clsWatcher("S:\Way4\OWSWork\WAY4P\Data\CARD_PRD\Embs\Transport\","*.001")
		
'		
'		InitializeTimers()
'		InitializeFileSystemWatcher(cFilePath,"*.001")
'
'		AddHandler notifyIcon.DoubleClick, AddressOf IconDoubleClick
		Dim resources As New System.ComponentModel.ComponentResourceManager(GetType(NotificationIcon))
		'notifyIcon.Icon = DirectCast(resources.GetObject("$this.Icon"), Icon)
		notifyIcon.Icon = DirectCast(resources.GetObject("page_white_go"), Icon)
		notifyIcon.ContextMenu = notificationMenu
	End Sub
	
	Private Sub InitializeTimers()
'		timer1 = New system.timers.Timer
'		timer2 = New system.Timers.Timer
'		AddHandler timer1.Elapsed, AddressOf timer1tick
'		AddHandler timer2.Elapsed, AddressOf timer2tick
'		timer1.Interval = 23000
'		timer2.Interval = 1000
	End Sub
	
	Private Sub InitializeFileSystemWatcher(strPath As String, strFilter As String)
'		fsw = New System.IO.FileSystemWatcher(strPath, strFilter)
'		' fsw.NotifyFilter = IO.NotifyFilters. NotifyFilters.
'		AddHandler fsw.Created, AddressOf OnFileCreate
'		AddHandler fsw.Renamed, AddressOf OnFileCreate
'		fsw.EnableRaisingEvents = True
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


	
	

	

	
	
	
	#End Region

	#Region "Event Handlers"
	' Обработка таймеров
	Private Sub MainTimerStart()
'		t2 = 0
'		timer1.Stop
'		timer2.Stop
'		Debug.WriteLine ("MainTimerStart")
'		timer1.Start
'		timer2.Start
	End Sub	
	
	Private Sub Timer1T
'		timer1.Stop()
'		timer2.Stop()
'		Debug.WriteLine ("Перемещение файлов")
'		' MoveFiles()
'		t2 = 0
	End Sub
	
	' Обработка события "Создан файл"	
	Private Sub OnFileCreate(sender As Object, e As System.IO.FileSystemEventArgs)
'		notifyIcon.ShowBalloonTip(20000, Now.ToString & ". Обнаружен новый файл", e.Name, ToolTipIcon.Info)
'		MainTimerStart()
	End Sub	
	
	Private Sub timer1tick(source As Object, e As ElapsedEventArgs)
'		Debug.WriteLine ("Timer1Tick")
'		Timer1T
	End Sub

	Private Sub timer2tick(source As Object, e As ElapsedEventArgs)
'		Debug.WriteLine ("Timer2Tick")
'		t2 = t2 + 1
'		notifyIcon.ShowBalloonTip(20000, "...", _
'			"Ожидание... " & FormatDateTime(TimeSerial(0,0,CInt(((timer1.Interval/1000)-t2))), DateFormat.LongTime).ToString, _
'			ToolTipIcon.Info)
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

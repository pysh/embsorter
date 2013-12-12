'
' Сделано в SharpDevelop.
' Пользователь: pnosov
' Дата: 11.12.2013
' Время: 17:50
' 
' Для изменения этого шаблона используйте Сервис | Настройка | Кодирование | Правка стандартных заголовков.
'
<MTAThread>_
Public Class clsWatcher
	Private WithEvents Shared _timerWaitNext As New System.Timers.Timer
	Private WithEvents Shared _timerTick As New System.Timers.Timer
	Private _FolderName As String
	Private _FileMask As String
	Private _sec As Integer	
	
	' <Subs>	
	Public Sub Run()
		_timerWaitNext.Start
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
	
	Property TimerTickInterval() As Integer
        Get
            Return _timerTick.Interval
        End Get
        Set(ByVal value As Integer)
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
		AddHandler _timerTick.Elapsed, AddressOf _timerTickElapsed
		_timerWaitNext.Interval = 23000
		_timerTick.Interval = 1000
	End Sub
	
	Private Sub _timerWaitNextElapsed()
		_timerWaitNext.Stop
		_timerTick.Stop
		_MoveFiles
	End Sub
	
	Private Sub _timerTickElapsed()
		Debug.WriteLine ("TimerTickElapsed", _sec.ToString)
		_sec = _sec + 1
'		notifyIcon.ShowBalloonTip(20000, "...", _
'			"Ожидание... " & FormatDateTime(TimeSerial(0,0,CInt(((timer1.Interval/1000) - _sec))), DateFormat.LongTime).ToString, _
'			ToolTipIcon.Info)
		timerTick()
	End Sub
	
	Public Sub timerTick()
		
	End Sub
	
End Class

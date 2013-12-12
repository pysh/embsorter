'
' Сделано в SharpDevelop.
' Пользователь: pnosov
' Дата: 11.12.2013
' Время: 17:50
' 
' Для изменения этого шаблона используйте Сервис | Настройка | Кодирование | Правка стандартных заголовков.
'
Public Class clsWatcher
	Private Shared _timerWaitNext As New System.Timers.Timer
	Private Shared _timerTick As New Timer
	
	Public Sub Run()
		_timerWaitNext.Start	
	End Sub
	
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
	
	
	
End Class


Public Module RLOG

    Public Const RLOG_PROCESS As String = "Process"
    Public Const RLOG_COMM As String = "Comm"
    Public Const RLOG_TRANS As String = "Transfer"

    Public Const RLOG_TARGET_DEBUGGER As String = "VsDebugger"

    Public logger As NLog.Logger
    Public loggerComm As NLog.Logger
    Public loggerTrans As NLog.Logger

    Private config As NLog.Config.LoggingConfiguration

    Private targetVsDebugger As NLog.Targets.DebuggerTarget
    Private targetCommLog As NLog.Targets.FileTarget

    Private ruleProcess As NLog.Config.LoggingRule
    Private ruleComm As NLog.Config.LoggingRule
    Private ruleTrans As NLog.Config.LoggingRule

    ''' <summary>
    ''' 初始化設定
    ''' </summary>
    ''' <param name="mylogdir">Log輸出資料夾</param>
    ''' <param name="myargument">程式Log用參數</param>
    ''' <param name="mydevicetype">通訊Log用裝置</param>
    ''' <param name="mydeviceid">通訊Log用ID</param>
    Public Sub InitConfig(mylogdir As String, myargument As String, mydevicetype As String, mydeviceid As String)

        ' 初始化
        config = New NLog.Config.LoggingConfiguration

        ' NLog.config內定義的variable
        config.Variables.Clear()
        config.Variables.Add("mylogdir", mylogdir)
        config.Variables.Add("myargument", myargument)
        config.Variables.Add("mydevicetype", mydevicetype)
        config.Variables.Add("mydeviceid", mydeviceid)

        ' NLog.config內定義的target
        targetVsDebugger = New NLog.Targets.DebuggerTarget
        targetVsDebugger.Name = RLOG_TARGET_DEBUGGER
        targetVsDebugger.Layout = New NLog.Layouts.SimpleLayout("${longdate} [${level}][${callsite}:${callsite-linenumber}] ${message}${onexception:inner=${exception:format=ToString}}")

        ' NLog.config內定義的rule
        ruleProcess = New NLog.Config.LoggingRule(RLOG_PROCESS)
        ruleProcess.LoggerNamePattern = RLOG_PROCESS
        ruleProcess.EnableLoggingForLevels(NLog.LogLevel.Info, NLog.LogLevel.Fatal)
        ruleProcess.Targets.Add(targetVsDebugger)

        ruleComm = New NLog.Config.LoggingRule(RLOG_COMM)
        ruleComm.LoggerNamePattern = RLOG_COMM
        ruleComm.EnableLoggingForLevels(NLog.LogLevel.Info, NLog.LogLevel.Info)

        ruleTrans = New NLog.Config.LoggingRule(RLOG_TRANS)
        ruleTrans.LoggerNamePattern = RLOG_TRANS
        ruleTrans.EnableLoggingForLevels(NLog.LogLevel.Info, NLog.LogLevel.Info)

        ' NLog.config
        config.AddTarget(targetVsDebugger)
        config.LoggingRules.Add(ruleProcess)
        config.LoggingRules.Add(ruleComm)
        config.LoggingRules.Add(ruleTrans)

        NLog.LogManager.Configuration = config
        NLog.LogManager.KeepVariablesOnReload = True
        NLog.LogManager.ReconfigExistingLoggers()

        'Dim b As Boolean = ruleProcess.NameMatches(RLOG_PROCESS)
        logger = NLog.LogManager.LogFactory.GetLogger(RLOG_PROCESS)
        loggerComm = NLog.LogManager.LogFactory.GetLogger(RLOG_COMM)
        loggerTrans = NLog.LogManager.LogFactory.GetLogger(RLOG_TRANS)
    End Sub

    ''' <summary>
    ''' 更新設定
    ''' </summary>
    Public Sub UpdateConfig()
        NLog.LogManager.ReconfigExistingLoggers()
    End Sub

    ''' <summary>
    ''' 寫入通訊Log
    ''' </summary>
    ''' <param name="strRS">Log定義</param>
    ''' <param name="strMessage">通訊傳文</param>
    Public Sub WriteCommLog(strRS As String, strMessage As String)
        loggerComm.Info(String.Format("[{0}]{1}", strRS, strMessage))
    End Sub

End Module
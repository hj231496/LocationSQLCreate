
Public Module RLOG_FILE
    Public Const RLOG_TARGET_PROCESSLOG As String = "ProcessLog"
    Public Const RLOG_TARGET_COMMLOG As String = "CommLog"
    Public Const RLOG_TARGET_TRANSLOG As String = "TransLog"

    Public Sub InitProcessFile(Optional ByVal intType As Integer = 0)

        ' 設定檔
        Dim config As NLog.Config.LoggingConfiguration = NLog.LogManager.Configuration
        If config Is Nothing Then
            config = New NLog.Config.LoggingConfiguration
        End If

        ' 目標
        Dim targetProcessLog As NLog.Targets.FileTarget = config.FindTargetByName(Of NLog.Targets.FileTarget)(RLOG_TARGET_PROCESSLOG)
        If targetProcessLog Is Nothing Then
            'intType(0:一般程式用(需要AUTOMATIC程式刪),1:GUI用(會自動刪LOG))
            If intType = 0 Then
                targetProcessLog = New NLog.Targets.FileTarget
                targetProcessLog.Name = RLOG_TARGET_PROCESSLOG
                targetProcessLog.Layout = New NLog.Layouts.SimpleLayout("${longdate} [${level}][${callsite}:${callsite-linenumber}] ${message}${onexception:inner=${exception:format=ToString}}")
                targetProcessLog.Encoding = System.Text.Encoding.UTF8
                targetProcessLog.FileName = New NLog.Layouts.SimpleLayout("${var:mylogdir}/${shortdate}/ProcessLog/${processname}${var:myargument}.log")
                targetProcessLog.ArchiveFileName = New NLog.Layouts.SimpleLayout("${var:mylogdir}/${shortdate}/ProcessLog/${processname}${var:myargument}.{##}.log")
                targetProcessLog.ArchiveNumbering = NLog.Targets.ArchiveNumberingMode.DateAndSequence
                targetProcessLog.ArchiveEvery = NLog.Targets.FileArchivePeriod.Hour
                targetProcessLog.ArchiveDateFormat = "HH"
                targetProcessLog.ArchiveAboveSize = "3072000"
            Else
                targetProcessLog = New NLog.Targets.FileTarget
                targetProcessLog.Name = RLOG_TARGET_PROCESSLOG
                targetProcessLog.Layout = New NLog.Layouts.SimpleLayout("${longdate} [${level}] ${message}${onexception:inner=${exception:format=ToString}}")
                targetProcessLog.Encoding = System.Text.Encoding.UTF8
                targetProcessLog.FileName = New NLog.Layouts.SimpleLayout("${var:mylogdir}/GuiLog/${var:mydevicetype}.log")
                targetProcessLog.ArchiveFileName = New NLog.Layouts.SimpleLayout("${var:mylogdir}/GuiLog/${var:mydevicetype}.{#}.log")
                targetProcessLog.ArchiveNumbering = NLog.Targets.ArchiveNumberingMode.Date
                targetProcessLog.ArchiveEvery = NLog.Targets.FileArchivePeriod.Day
                targetProcessLog.ArchiveDateFormat = "yyyyMMdd"
                'targetGuiLog.ArchiveAboveSize = "3072000" '3M
                targetProcessLog.MaxArchiveFiles = 30
            End If
            'targetProcessLog = New NLog.Targets.FileTarget
            'targetProcessLog.Name = RLOG_TARGET_PROCESSLOG
            'targetProcessLog.Layout = New NLog.Layouts.SimpleLayout("${longdate} [${level}][${callsite}:${callsite-linenumber}] ${message}${onexception:inner=${exception:format=ToString}}")
            'targetProcessLog.Encoding = System.Text.Encoding.UTF8
            'targetProcessLog.FileName = New NLog.Layouts.SimpleLayout("${var:mylogdir}/${shortdate}/ProcessLog/${processname}${var:myargument}.log")
            'targetProcessLog.ArchiveFileName = New NLog.Layouts.SimpleLayout("${var:mylogdir}/${shortdate}/ProcessLog/${processname}${var:myargument}.{##}.log")
            'targetProcessLog.ArchiveNumbering = NLog.Targets.ArchiveNumberingMode.DateAndSequence
            'targetProcessLog.ArchiveEvery = NLog.Targets.FileArchivePeriod.Hour
            'targetProcessLog.ArchiveDateFormat = "HH"
            'targetProcessLog.ArchiveAboveSize = "3072000"
            'targetProcessLog.MaxArchiveFiles = 20

            config.AddTarget(targetProcessLog)

            ' 規則
            Dim ruleProcess As NLog.Config.LoggingRule = config.FindRuleByName(RLOG_PROCESS)
            If ruleProcess IsNot Nothing Then
                ruleProcess.Targets.Add(targetProcessLog)
            End If
        End If

        ' 重新設定
        NLog.LogManager.Configuration = config

    End Sub

    Public Sub InitCommFile()

        ' 設定檔
        Dim config As NLog.Config.LoggingConfiguration = NLog.LogManager.Configuration
        If config Is Nothing Then
            config = New NLog.Config.LoggingConfiguration
        End If

        ' 目標
        Dim targetCommLog As NLog.Targets.FileTarget = config.FindTargetByName(Of NLog.Targets.FileTarget)(RLOG_TARGET_COMMLOG)
        If targetCommLog Is Nothing Then
            targetCommLog = New NLog.Targets.FileTarget
            targetCommLog.Name = RLOG_TARGET_COMMLOG
            targetCommLog.Layout = New NLog.Layouts.SimpleLayout("${longdate} ${message}")
            targetCommLog.Encoding = System.Text.Encoding.UTF8
            targetCommLog.FileName = New NLog.Layouts.SimpleLayout("${var:mylogdir}/${shortdate}/CommLog/${var:mydevicetype}_${var:mydeviceid}.log")
            targetCommLog.ArchiveFileName = New NLog.Layouts.SimpleLayout("${var:mylogdir}/${shortdate}/CommLog/${var:mydevicetype}_${var:mydeviceid}_{##}.log")
            targetCommLog.ArchiveNumbering = NLog.Targets.ArchiveNumberingMode.DateAndSequence
            targetCommLog.ArchiveEvery = NLog.Targets.FileArchivePeriod.Hour
            targetCommLog.ArchiveDateFormat = "HH"
            targetCommLog.ArchiveAboveSize = "3072000"
            'targetCommLog.MaxArchiveFiles = 20

            config.AddTarget(targetCommLog)

            ' 規則
            Dim ruleComm As NLog.Config.LoggingRule = config.FindRuleByName(RLOG_COMM)
            If ruleComm IsNot Nothing Then
                ruleComm.Targets.Add(targetCommLog)
            End If
        End If

        ' 重新設定
        NLog.LogManager.Configuration = config

    End Sub

    Public Sub InitTransFile()

        ' 設定檔
        Dim config As NLog.Config.LoggingConfiguration = NLog.LogManager.Configuration
        If config Is Nothing Then
            config = New NLog.Config.LoggingConfiguration
        End If

        ' 目標
        Dim targetTransLog As NLog.Targets.FileTarget = config.FindTargetByName(Of NLog.Targets.FileTarget)(RLOG_TARGET_TRANSLOG)
        If targetTransLog Is Nothing Then
            targetTransLog = New NLog.Targets.FileTarget
            targetTransLog.Name = RLOG_TARGET_TRANSLOG
            targetTransLog.Layout = New NLog.Layouts.SimpleLayout("${longdate} [${level}] ${message}${onexception:inner=${exception:format=ToString}}")
            targetTransLog.Encoding = System.Text.Encoding.UTF8
            targetTransLog.FileName = New NLog.Layouts.SimpleLayout("${var:mylogdir}/${shortdate}/TransLog/${var:mydevicetype}_${var:mydeviceid}.log")
            targetTransLog.ArchiveFileName = New NLog.Layouts.SimpleLayout("${var:mylogdir}/${shortdate}/TransLog/${var:mydevicetype}_${var:mydeviceid}_{##}.log")
            targetTransLog.ArchiveNumbering = NLog.Targets.ArchiveNumberingMode.DateAndSequence
            targetTransLog.ArchiveEvery = NLog.Targets.FileArchivePeriod.Hour
            targetTransLog.ArchiveDateFormat = "HH"
            targetTransLog.ArchiveAboveSize = "3072000" '3M
            'targetTransLog.MaxArchiveFiles = 20

            config.AddTarget(targetTransLog)

            ' 規則
            Dim ruleComm As NLog.Config.LoggingRule = config.FindRuleByName(RLOG_TRANS)
            If ruleComm IsNot Nothing Then
                ruleComm.Targets.Add(targetTransLog)
            End If
        End If

        ' 重新設定
        NLog.LogManager.Configuration = config

    End Sub

End Module
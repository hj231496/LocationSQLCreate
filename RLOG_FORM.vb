
Public Module RLOG_FORM
    Public Const RLOG_TARGET_PROCESSBOX As String = "ProcessRichTextBox"
    Public Const RLOG_TARGET_COMMBOX As String = "CommRichTextBox"

    ''' <summary>
    ''' 初始化Process Log用
    ''' </summary>
    ''' <param name="formName">Form名稱</param>
    ''' <param name="controlName">RichTextBox名稱</param>
    Public Sub InitProcessBox(formName As String, controlName As String)

        ' 設定檔
        Dim config As NLog.Config.LoggingConfiguration = NLog.LogManager.Configuration
        If config Is Nothing Then
            config = New NLog.Config.LoggingConfiguration
        End If

        ' 目標
        Dim targetProcessRichTextBox As NLog.Windows.Forms.RichTextBoxTarget = config.FindTargetByName(Of NLog.Windows.Forms.RichTextBoxTarget)(RLOG_TARGET_PROCESSBOX)
        If targetProcessRichTextBox Is Nothing Then
            targetProcessRichTextBox = New NLog.Windows.Forms.RichTextBoxTarget With {
                .Name = RLOG_TARGET_PROCESSBOX,
                .Layout = New NLog.Layouts.SimpleLayout("${longdate} [${level}][${callsite}:${callsite-linenumber}] ${message}"),
                .AutoScroll = True,
                .MaxLines = 1000,
                .ControlName = controlName,
                .FormName = formName,
                .UseDefaultRowColoringRules = True,
                .AllowAccessoryFormCreation = False
            }

            config.AddTarget(targetProcessRichTextBox)

            ' 規則
            Dim ruleProcess As NLog.Config.LoggingRule = config.FindRuleByName(RLOG_PROCESS)
            If ruleProcess IsNot Nothing Then
                ruleProcess.Targets.Add(targetProcessRichTextBox)
            End If
        End If

        ' 重新設定
        NLog.LogManager.Configuration = config

    End Sub

    Public Sub InitCommBox(formName As String, controlName As String)

        ' 設定檔
        Dim config As NLog.Config.LoggingConfiguration = NLog.LogManager.Configuration
        If config Is Nothing Then
            config = New NLog.Config.LoggingConfiguration
        End If

        ' 目標
        Dim targetCommRichTextBox As NLog.Windows.Forms.RichTextBoxTarget = config.FindTargetByName(Of NLog.Windows.Forms.RichTextBoxTarget)(RLOG_TARGET_COMMBOX)
        If targetCommRichTextBox Is Nothing Then
            targetCommRichTextBox = New NLog.Windows.Forms.RichTextBoxTarget With {
                .Name = RLOG_TARGET_COMMBOX,
                .Layout = New NLog.Layouts.SimpleLayout("${longdate} ${message}"),
                .AutoScroll = True,
                .MaxLines = 1000,
                .ControlName = controlName,
                .FormName = formName,
                .UseDefaultRowColoringRules = True,
                .AllowAccessoryFormCreation = True
            }

            config.AddTarget(targetCommRichTextBox)

            ' 規則
            Dim ruleComm As NLog.Config.LoggingRule = config.FindRuleByName(RLOG_COMM)
            If ruleComm IsNot Nothing Then
                ruleComm.Targets.Add(targetCommRichTextBox)
            End If
        End If

        ' 重新設定
        NLog.LogManager.Configuration = config

    End Sub

End Module
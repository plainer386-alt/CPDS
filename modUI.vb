Attribute VB_Name = "modUI"
Option Explicit

'==============================================================================
' Compressor Preliminary Design System
'
' Module : modUI
' Purpose: User Interface Management
'
'==============================================================================

'==============================================================================
' INITIALIZATION
'==============================================================================

Public Sub InitializeMainForm(frm As Object)

    ApplyMainTheme frm

    CenterForm frm

    FillNavigation frm.lstNavigation

    SetStatus frm.lblStatus, MSG_READY

End Sub

Public Sub InitializeStabilityForm(frm As Object)

    ApplyMainTheme frm

    CenterForm frm

    FillCriteria frm

    SetStatus frm.lblStatus, MSG_READY

End Sub

'==============================================================================
' NAVIGATION
'==============================================================================

Public Sub FillNavigation(lst As MSForms.ListBox)

    lst.Clear

    lst.AddItem "Home"

    lst.AddItem "Geometry"

    lst.AddItem "Aerodynamics"

    lst.AddItem "Stability"

    lst.AddItem "Performance"

    lst.AddItem "Reports"

    lst.AddItem "Settings"

    lst.ListIndex = 0

End Sub

Public Function SelectedModule( _
    ByVal lst As MSForms.ListBox) _
    As eModule

    Select Case lst.ListIndex

        Case 0

            SelectedModule = mdHome

        Case 1

            SelectedModule = mdGeometry

        Case 2

            SelectedModule = mdAerodynamics

        Case 3

            SelectedModule = mdStability

        Case 4

            SelectedModule = mdPerformance

        Case 5

            SelectedModule = mdReports

        Case 6

            SelectedModule = mdSettings

        Case Else

            SelectedModule = mdHome

    End Select

End Function

'==============================================================================
' STABILITY METHODS
'==============================================================================

Public Sub FillCriteria(frm As Object)

    With frm

        .chkKoch.Value = True

        .chkAungier.Value = True

        .chkJianLi.Value = True

        .chkLieblein.Value = False

        .chkDeHaller.Value = False

    End With

End Sub

'==============================================================================
' STATUS
'==============================================================================

Public Sub UIReady(frm As Object)

    EnableInterface frm

    SetStatus frm.lblStatus, _
        MSG_READY, _
        stReady

End Sub

Public Sub UIRunning(frm As Object)

    DisableInterface frm

    frm.cmdRun.Enabled = True

    SetStatus frm.lblStatus, _
        MSG_RUNNING, _
        stRunning

End Sub

Public Sub UICompleted(frm As Object)

    EnableInterface frm

    SetStatus frm.lblStatus, _
        MSG_COMPLETE, _
        stCompleted

End Sub

Public Sub UIError( _
    frm As Object, _
    ByVal Message As String)

    EnableInterface frm

    SetStatus frm.lblStatus, _
        Message, _
        stError

End Sub

'==============================================================================
' INPUT VALIDATION
'==============================================================================

Public Function ValidateInteger( _
    txt As MSForms.TextBox) _
    As Boolean

    ValidateInteger = False

    If Trim$(txt.Text) = "" Then Exit Function

    If Not IsNumeric(txt.Text) Then Exit Function

    ValidateInteger = True

End Function

Public Function ValidateDouble( _
    txt As MSForms.TextBox) _
    As Boolean

    ValidateDouble = False

    If Trim$(txt.Text) = "" Then Exit Function

    If Not IsNumeric(txt.Text) Then Exit Function

    ValidateDouble = True

End Function

Public Function ValidateRange( _
    txt As MSForms.TextBox) _
    As Boolean

    ValidateRange = False

    If Len(Trim$(txt.Text)) = 0 Then Exit Function

    ValidateRange = True

End Function

'==============================================================================
' ENABLE / DISABLE
'==============================================================================

Public Sub EnableRun(frm As Object)

    frm.cmdRun.Enabled = True

End Sub

Public Sub DisableRun(frm As Object)

    frm.cmdRun.Enabled = False

End Sub

Public Sub EnableOutput(frm As Object)

    frm.txtOutput.Enabled = True

    frm.cmdBrowseOutput.Enabled = True

End Sub

Public Sub DisableOutput(frm As Object)

    frm.txtOutput.Enabled = False

    frm.cmdBrowseOutput.Enabled = False

End Sub

'==============================================================================
' WORKSHEETS
'==============================================================================

Public Sub FillWorksheetList(cmb As MSForms.ComboBox)

    Dim ws As Worksheet

    cmb.Clear

    For Each ws In ThisWorkbook.Worksheets

        cmb.AddItem ws.Name

    Next ws

End Sub

Public Function WorksheetExists( _
    ByVal SheetName As String) _
    As Boolean

    Dim ws As Worksheet

    WorksheetExists = False

    For Each ws In ThisWorkbook.Worksheets

        If StrComp(ws.Name, SheetName, vbTextCompare) = 0 Then

            WorksheetExists = True

            Exit Function

        End If

    Next ws

End Function

'==============================================================================
' PROJECT
'==============================================================================

Public Sub UpdateProjectName( _
    frm As Object, _
    ByVal ProjectName As String)

    If Trim$(ProjectName) = "" Then

        frm.lblProject.Caption = PROJECT_UNNAMED

    Else

        frm.lblProject.Caption = ProjectName

    End If

End Sub

Public Sub UpdateVersion(frm As Object)

    frm.lblVersion.Caption = _
        "Version " & APP_VERSION

End Sub

'==============================================================================
' INPUT
'==============================================================================

Public Function ReadInteger( _
    txt As MSForms.TextBox, _
    Optional ByVal DefaultValue As Long = 0) _
    As Long

    If ValidateInteger(txt) Then

        ReadInteger = CLng(txt.Text)

    Else

        ReadInteger = DefaultValue

    End If

End Function

Public Function ReadDouble( _
    txt As MSForms.TextBox, _
    Optional ByVal DefaultValue As Double = 0#) _
    As Double

    If ValidateDouble(txt) Then

        ReadDouble = CDbl(txt.Text)

    Else

        ReadDouble = DefaultValue

    End If

End Function

Public Function ReadString( _
    txt As MSForms.TextBox) _
    As String

    ReadString = Trim$(txt.Text)

End Function

'==============================================================================
' RANGE
'==============================================================================

Public Function GetInputRange( _
    frm As Object) _
    As String

    GetInputRange = Trim$(frm.txtInput.Text)

End Function

Public Function GetOutputRange( _
    frm As Object) _
    As String

    GetOutputRange = Trim$(frm.txtOutput.Text)

End Function

Public Sub ClearInput(frm As Object)

    frm.txtInput.Text = ""

    frm.txtOutput.Text = ""

End Sub

'==============================================================================
' CHECKBOXES
'==============================================================================

Public Function UseKoch(frm As Object) As Boolean

    UseKoch = frm.chkKoch.Value

End Function

Public Function UseAungier(frm As Object) As Boolean

    UseAungier = frm.chkAungier.Value

End Function

Public Function UseJianLi(frm As Object) As Boolean

    UseJianLi = frm.chkJianLi.Value

End Function

Public Function UseLieblein(frm As Object) As Boolean

    UseLieblein = frm.chkLieblein.Value

End Function

Public Function UseDeHaller(frm As Object) As Boolean

    UseDeHaller = frm.chkDeHaller.Value

End Function

'==============================================================================
' BUTTONS
'==============================================================================

Public Sub LockButtons(frm As Object)

    frm.cmdRun.Enabled = False

    frm.cmdClose.Enabled = False

    frm.cmdBrowseInput.Enabled = False

    frm.cmdBrowseOutput.Enabled = False

End Sub

Public Sub UnlockButtons(frm As Object)

    frm.cmdRun.Enabled = True

    frm.cmdClose.Enabled = True

    frm.cmdBrowseInput.Enabled = True

    frm.cmdBrowseOutput.Enabled = True

End Sub

'==============================================================================
' RESET
'==============================================================================

Public Sub ResetStabilityForm(frm As Object)

    frm.txtInput.Text = ""

    frm.txtOutput.Text = ""

    frm.txtStages.Text = ""

    frm.txtFirstRow.Text = ""

    frm.chkKoch.Value = True

    frm.chkAungier.Value = True

    frm.chkJianLi.Value = True

    frm.chkLieblein.Value = False

    frm.chkDeHaller.Value = False

    SetStatus frm.lblStatus, _
              MSG_READY, _
              stReady

End Sub

'==============================================================================
' INFORMATION
'==============================================================================

Public Sub ShowReady(frm As Object)

    SetStatus frm.lblStatus, _
              MSG_READY, _
              stReady

End Sub

Public Sub ShowCalculation(frm As Object)

    SetStatus frm.lblStatus, _
              MSG_RUNNING, _
              stRunning

End Sub

Public Sub ShowFinished(frm As Object)

    SetStatus frm.lblStatus, _
              MSG_COMPLETE, _
              stCompleted

End Sub

Public Sub ShowError( _
    frm As Object, _
    ByVal ErrorMessage As String)

    SetStatus frm.lblStatus, _
              ErrorMessage, _
              stError

End Sub

'==============================================================================
' BROWSE DIALOGS
'==============================================================================

Public Function BrowseInputRange(frm As Object) As Boolean

    Dim rg As Range

    On Error Resume Next

    Set rg = Application.InputBox( _
                    Prompt:="Select input data range", _
                    Title:="CPDS", _
                    Type:=8)

    On Error GoTo 0

    If rg Is Nothing Then

        BrowseInputRange = False

        Exit Function

    End If

    frm.txtInput.Text = rg.Address(False, False)

    BrowseInputRange = True

End Function

Public Function BrowseOutputRange(frm As Object) As Boolean

    Dim rg As Range

    On Error Resume Next

    Set rg = Application.InputBox( _
                    Prompt:="Select output range", _
                    Title:="CPDS", _
                    Type:=8)

    On Error GoTo 0

    If rg Is Nothing Then

        BrowseOutputRange = False

        Exit Function

    End If

    frm.txtOutput.Text = rg.Address(False, False)

    BrowseOutputRange = True

End Function

'==============================================================================
' FORM VALIDATION
'==============================================================================

Public Function ValidateStabilityForm(frm As Object) As Boolean

    ValidateStabilityForm = False

    If Not ValidateRange(frm.txtInput) Then

        MsgBox "Input range is not specified.", _
               vbExclamation, _
               APP_SHORT

        frm.txtInput.SetFocus

        Exit Function

    End If

    If Not ValidateRange(frm.txtOutput) Then

        MsgBox "Output range is not specified.", _
               vbExclamation, _
               APP_SHORT

        frm.txtOutput.SetFocus

        Exit Function

    End If

    If Not ValidateInteger(frm.txtStages) Then

        MsgBox "Invalid number of stages.", _
               vbExclamation, _
               APP_SHORT

        frm.txtStages.SetFocus

        Exit Function

    End If

    If Not ValidateInteger(frm.txtFirstRow) Then

        MsgBox "Invalid first stage row.", _
               vbExclamation, _
               APP_SHORT

        frm.txtFirstRow.SetFocus

        Exit Function

    End If

    ValidateStabilityForm = True

End Function

'==============================================================================
' PROGRESS
'==============================================================================

Public Sub BeginCalculation(frm As Object)

    LockButtons frm

    ShowCalculation frm

    DoEvents

End Sub

Public Sub EndCalculation(frm As Object)

    UnlockButtons frm

    ShowFinished frm

    DoEvents

End Sub

Public Sub AbortCalculation(frm As Object)

    UnlockButtons frm

    ShowReady frm

End Sub

'==============================================================================
' WORKSPACE
'==============================================================================

Public Sub ClearWorkspace(frm As Object)

    Dim ctl As Control

    For Each ctl In frm.Controls

        Select Case TypeName(ctl)

            Case "TextBox"

                ctl.Text = ""

            Case "CheckBox"

                ctl.Value = False

            Case "ComboBox"

                ctl.ListIndex = -1

        End Select

    Next ctl

End Sub

'==============================================================================
' NAVIGATION
'==============================================================================

Public Sub SelectModule( _
            frm As Object, _
            ByVal ModuleID As eModule)

    Select Case ModuleID

        Case mdHome

            frm.lstNavigation.ListIndex = 0

        Case mdGeometry

            frm.lstNavigation.ListIndex = 1

        Case mdAerodynamics

            frm.lstNavigation.ListIndex = 2

        Case mdStability

            frm.lstNavigation.ListIndex = 3

        Case mdPerformance

            frm.lstNavigation.ListIndex = 4

        Case mdReports

            frm.lstNavigation.ListIndex = 5

        Case mdSettings

            frm.lstNavigation.ListIndex = 6

    End Select

End Sub

'==============================================================================
' CALCULATION OPTIONS
'==============================================================================

Public Function SelectedCriteria(frm As Object) As Collection

    Dim col As New Collection

    If frm.chkKoch.Value Then
        col.Add crKoch
    End If

    If frm.chkAungier.Value Then
        col.Add crAungier
    End If

    If frm.chkJianLi.Value Then
        col.Add crJianLi
    End If

    If frm.chkLieblein.Value Then
        col.Add crLieblein
    End If

    If frm.chkDeHaller.Value Then
        col.Add crDeHaller
    End If

    Set SelectedCriteria = col

End Function

'==============================================================================
' RUN
'==============================================================================

Public Function CanRunCalculation(frm As Object) As Boolean

    CanRunCalculation = False

    If Not ValidateStabilityForm(frm) Then Exit Function

    If SelectedCriteria(frm).Count = 0 Then

        MsgBox "Select at least one stability criterion.", _
               vbInformation, _
               APP_SHORT

        Exit Function

    End If

    CanRunCalculation = True

End Function

'==============================================================================
' END OF MODULE
'==============================================================================

'==============================================================================
' MAIN WINDOW
'==============================================================================

Public Sub OpenModule( _
            frm As Object, _
            ByVal ModuleID As eModule)

    Select Case ModuleID

        Case mdHome

            ShowHome frm

        Case mdGeometry

            ShowGeometry frm

        Case mdAerodynamics

            ShowAerodynamics frm

        Case mdStability

            ShowStability frm

        Case mdPerformance

            ShowPerformance frm

        Case mdReports

            ShowReports frm

        Case mdSettings

            ShowSettings frm

    End Select

End Sub

'==============================================================================
' HOME
'==============================================================================

Public Sub ShowHome(frm As Object)

    frm.lblWorkspace.Caption = _
        "Welcome to Compressor Preliminary Design System"

    frm.lblDescription.Caption = _
        "Select a module from the navigation panel."

    SetStatus frm.lblStatus, _
              MSG_READY, _
              stReady

End Sub

'==============================================================================
' GEOMETRY
'==============================================================================

Public Sub ShowGeometry(frm As Object)

    frm.lblWorkspace.Caption = "Geometry"

    frm.lblDescription.Caption = _
        "Geometry design module is under development."

End Sub

'==============================================================================
' AERODYNAMICS
'==============================================================================

Public Sub ShowAerodynamics(frm As Object)

    frm.lblWorkspace.Caption = "Aerodynamics"

    frm.lblDescription.Caption = _
        "Aerodynamic calculation module is under development."

End Sub

'==============================================================================
' STABILITY
'==============================================================================

Public Sub ShowStability(frm As Object)

    frm.Hide

    frmStability.Show

End Sub

'==============================================================================
' PERFORMANCE
'==============================================================================

Public Sub ShowPerformance(frm As Object)

    frm.lblWorkspace.Caption = "Performance"

    frm.lblDescription.Caption = _
        "Performance module is under development."

End Sub

'==============================================================================
' REPORTS
'==============================================================================

Public Sub ShowReports(frm As Object)

    frm.lblWorkspace.Caption = "Reports"

    frm.lblDescription.Caption = _
        "Report generation module is under development."

End Sub

'==============================================================================
' SETTINGS
'==============================================================================

Public Sub ShowSettings(frm As Object)

    frm.lblWorkspace.Caption = "Settings"

    frm.lblDescription.Caption = _
        "Application settings."

End Sub

'==============================================================================
' NAVIGATION EVENTS
'==============================================================================

Public Sub NavigationClick(frm As Object)

    Dim m As eModule

    m = SelectedModule(frm.lstNavigation)

    OpenModule frm, m

End Sub

'==============================================================================
' STATUS BAR
'==============================================================================

Public Sub UpdateStatus( _
            frm As Object, _
            ByVal Text As String)

    frm.lblStatus.Caption = Text

End Sub

Public Sub ClearStatus(frm As Object)

    frm.lblStatus.Caption = MSG_READY

End Sub

'==============================================================================
' MAIN FORM EVENTS
'==============================================================================

Public Sub MainFormLoad(frm As Object)

    InitializeMainForm frm

    ShowHome frm

End Sub

Public Sub MainFormClose(frm As Object)

    Unload frm

End Sub

'==============================================================================
' STABILITY FORM EVENTS
'==============================================================================

Public Sub StabilityFormLoad(frm As Object)

    InitializeStabilityForm frm

    FillWorksheetList frm.cmbWorksheet

End Sub

Public Sub StabilityFormClose(frm As Object)

    frm.Hide

    frmMain.Show

End Sub

'==============================================================================
' RUN EVENTS
'==============================================================================

Public Sub RunStability(frm As Object)

    If Not CanRunCalculation(frm) Then Exit Sub

    BeginCalculation frm

    On Error GoTo CalculationError

    '==========================================================
    ' Solver will be connected here
    '
    ' Call RunKoch(...)
    ' Call RunAungier(...)
    ' Call RunJianLi(...)
    '==========================================================

    EndCalculation frm

    Exit Sub

CalculationError:

    UIError frm, Err.Description

End Sub

'==============================================================================
' HELPERS
'==============================================================================

Public Function YesNo( _
            ByVal Question As String) _
            As Boolean

    YesNo = _
        (MsgBox(Question, _
                vbQuestion + vbYesNo, _
                APP_SHORT) = vbYes)

End Function

Public Sub Information( _
            ByVal Text As String)

    MsgBox Text, _
           vbInformation, _
           APP_SHORT

End Sub

Public Sub Warning( _
            ByVal Text As String)

    MsgBox Text, _
           vbExclamation, _
           APP_SHORT

End Sub

Public Sub Critical( _
            ByVal Text As String)

    MsgBox Text, _
           vbCritical, _
           APP_SHORT

End Sub

'==============================================================================
' END OF FILE
'==============================================================================

'==============================================================================
' USER INTERFACE CONTROL
'==============================================================================

Public Sub LockInput(frm As Object)

    On Error Resume Next

    frm.txtInput.Enabled = False
    frm.txtOutput.Enabled = False
    frm.txtStages.Enabled = False
    frm.txtFirstRow.Enabled = False

    frm.cmbWorksheet.Enabled = False

    frm.chkKoch.Enabled = False
    frm.chkAungier.Enabled = False
    frm.chkJianLi.Enabled = False
    frm.chkLieblein.Enabled = False
    frm.chkDeHaller.Enabled = False

    On Error GoTo 0

End Sub

Public Sub UnlockInput(frm As Object)

    On Error Resume Next

    frm.txtInput.Enabled = True
    frm.txtOutput.Enabled = True
    frm.txtStages.Enabled = True
    frm.txtFirstRow.Enabled = True

    frm.cmbWorksheet.Enabled = True

    frm.chkKoch.Enabled = True
    frm.chkAungier.Enabled = True
    frm.chkJianLi.Enabled = True
    frm.chkLieblein.Enabled = True
    frm.chkDeHaller.Enabled = True

    On Error GoTo 0

End Sub

'==============================================================================
' BROWSE BUTTONS
'==============================================================================

Public Sub EnableBrowseButtons(frm As Object)

    On Error Resume Next

    frm.cmdBrowseInput.Enabled = True
    frm.cmdBrowseOutput.Enabled = True

    On Error GoTo 0

End Sub

Public Sub DisableBrowseButtons(frm As Object)

    On Error Resume Next

    frm.cmdBrowseInput.Enabled = False
    frm.cmdBrowseOutput.Enabled = False

    On Error GoTo 0

End Sub

'==============================================================================
' CALCULATION BUTTONS
'==============================================================================

Public Sub EnableCalculationButtons(frm As Object)

    On Error Resume Next

    frm.cmdRun.Enabled = True
    frm.cmdClose.Enabled = True
    frm.cmdReset.Enabled = True

    On Error GoTo 0

End Sub

Public Sub DisableCalculationButtons(frm As Object)

    On Error Resume Next

    frm.cmdRun.Enabled = False
    frm.cmdClose.Enabled = False
    frm.cmdReset.Enabled = False

    On Error GoTo 0

End Sub

'==============================================================================
' WAIT CURSOR
'==============================================================================

Public Sub BeginBusy()

    DoEvents

    Application.Cursor = xlWait

    Application.StatusBar = "CPDS : Running calculation..."

End Sub

Public Sub EndBusy()

    DoEvents

    Application.Cursor = xlDefault

    Application.StatusBar = False

End Sub

'==============================================================================
' PROGRESS
'==============================================================================

Public Sub ProgressStart(frm As Object)

    On Error Resume Next

    frm.lblProgress.Caption = "0 %"

    frm.lblStatus.Caption = "Preparing..."

    DoEvents

End Sub

Public Sub ProgressUpdate( _
            frm As Object, _
            ByVal Percent As Double, _
            Optional ByVal Message As String = "")

    On Error Resume Next

    If Percent < 0 Then Percent = 0
    If Percent > 100 Then Percent = 100

    frm.lblProgress.Caption = _
        Format(Percent, "0") & " %"

    If Len(Message) > 0 Then

        frm.lblStatus.Caption = Message

    End If

    DoEvents

End Sub

Public Sub ProgressFinish(frm As Object)

    On Error Resume Next

    frm.lblProgress.Caption = "100 %"

    frm.lblStatus.Caption = MSG_COMPLETE

End Sub

'==============================================================================
' FIELD HIGHLIGHTING
'==============================================================================

Public Sub HighlightControl(ctrl As Control)

    On Error Resume Next

    ctrl.BackColor = RGB(255, 225, 225)

End Sub

Public Sub ClearHighlight(ctrl As Control)

    On Error Resume Next

    ctrl.BackColor = vbWhite

End Sub

Public Sub HighlightError( _
            ctrl As Control, _
            ByVal Message As String)

    HighlightControl ctrl

    ctrl.SetFocus

    MsgBox Message, _
           vbExclamation, _
           APP_SHORT

End Sub

'==============================================================================
' FORM RESET
'==============================================================================

Public Sub ResetHighlights(frm As Object)

    Dim ctl As Control

    For Each ctl In frm.Controls

        Select Case TypeName(ctl)

            Case "TextBox"

                ctl.BackColor = vbWhite

        End Select

    Next ctl

End Sub

'==============================================================================
' ERROR HANDLING
'==============================================================================

Private m_LastError As String

Public Property Get LastError() As String

    LastError = m_LastError

End Property

Public Sub ClearLastError()

    m_LastError = ""

End Sub

Public Sub ShowError( _
            ByVal Source As String, _
            ByVal Description As String)

    m_LastError = Description

    MsgBox _
        "Module : " & Source & vbCrLf & _
        vbCrLf & _
        Description, _
        vbCritical, _
        APP_NAME

End Sub

Public Sub HandleError( _
            ByVal Source As String)

    ShowError Source, Err.Description

End Sub

Public Function SafeExecute( _
            ByVal Source As String) _
            As Boolean

    If Err.Number = 0 Then

        SafeExecute = True

        Exit Function

    End If

    ShowError Source, Err.Description

    SafeExecute = False

End Function

'==============================================================================
' COMMON DIALOGS
'==============================================================================

Public Sub ShowInformation( _
            ByVal Text As String)

    MsgBox Text, _
           vbInformation, _
           APP_NAME

End Sub

Public Sub ShowWarning( _
            ByVal Text As String)

    MsgBox Text, _
           vbExclamation, _
           APP_NAME

End Sub

Public Function AskQuestion( _
            ByVal Question As String) _
            As Boolean

    AskQuestion = _
        (MsgBox(Question, _
                vbQuestion + vbYesNo, _
                APP_NAME) = vbYes)

End Function

'==============================================================================
' FORM STATE
'==============================================================================

Public Sub SetReadyState(frm As Object)

    UnlockInput frm

    EnableBrowseButtons frm

    EnableCalculationButtons frm

    EndBusy

    SetStatus frm.lblStatus, _
              MSG_READY, _
              stReady

End Sub

Public Sub SetRunningState(frm As Object)

    LockInput frm

    DisableBrowseButtons frm

    DisableCalculationButtons frm

    frm.cmdRun.Enabled = True

    BeginBusy

    SetStatus frm.lblStatus, _
              MSG_RUNNING, _
              stRunning

End Sub

Public Sub SetFinishedState(frm As Object)

    UnlockInput frm

    EnableBrowseButtons frm

    EnableCalculationButtons frm

    EndBusy

    SetStatus frm.lblStatus, _
              MSG_COMPLETE, _
              stCompleted

End Sub

Public Sub SetErrorState( _
            frm As Object, _
            ByVal ErrorText As String)

    UnlockInput frm

    EnableBrowseButtons frm

    EnableCalculationButtons frm

    EndBusy

    SetStatus frm.lblStatus, _
              ErrorText, _
              stError

End Sub

'==============================================================================
' PLACEHOLDER MODULES
'==============================================================================

Public Sub OpenGeometry()

    ShowInformation _
        "Geometry module is not available in this version."

End Sub

Public Sub OpenAerodynamics()

    ShowInformation _
        "Aerodynamics module is not available in this version."

End Sub

Public Sub OpenPerformance()

    ShowInformation _
        "Performance module is not available in this version."

End Sub

Public Sub OpenReports()

    ShowInformation _
        "Reports module is not available in this version."

End Sub

Public Sub OpenSettings()

    ShowInformation _
        "Settings module is not available in this version."

End Sub

'==============================================================================
' FORM UTILITIES
'==============================================================================

Public Sub SetFormCaption( _
            frm As Object, _
            ByVal Caption As String)

    frm.Caption = _
        APP_SHORT & " - " & Caption

End Sub

Public Sub SetWorkspaceTitle( _
            frm As Object, _
            ByVal Title As String)

    On Error Resume Next

    frm.lblWorkspace.Caption = Title

End Sub

Public Sub SetWorkspaceDescription( _
            frm As Object, _
            ByVal Description As String)

    On Error Resume Next

    frm.lblDescription.Caption = Description

End Sub

Public Sub RefreshUI(frm As Object)

    DoEvents

    frm.Repaint

End Sub

'==============================================================================
' DEVELOPMENT
'==============================================================================

Public Function IsDevelopmentMode() As Boolean

    IsDevelopmentMode = True

End Function

Public Sub NotImplemented( _
            ByVal FeatureName As String)

    MsgBox _
        FeatureName & vbCrLf & vbCrLf & _
        "This functionality will be implemented in a future version.", _
        vbInformation, _
        APP_NAME

End Sub

'==============================================================================
' END OF MODULE
'==============================================================================
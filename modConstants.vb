Attribute VB_Name = "modConstants"
Option Explicit

'==============================================================================
' Compressor Preliminary Design System
' Global Constants
' Version 1.0
'==============================================================================

'----------------------------------------------------------------------------
' Application
'----------------------------------------------------------------------------

Public Const APP_NAME As String = "Compressor Preliminary Design System"
Public Const APP_SHORT As String = "CPDS"
Public Const APP_VERSION As String = "1.0"

'----------------------------------------------------------------------------
' Main Form
'----------------------------------------------------------------------------

Public Const MAIN_WIDTH As Long = 1100
Public Const MAIN_HEIGHT As Long = 700

Public Const HEADER_HEIGHT As Long = 48
Public Const STATUS_HEIGHT As Long = 24

Public Const NAV_WIDTH As Long = 220

Public Const CONTENT_MARGIN As Long = 12

'----------------------------------------------------------------------------
' Frames
'----------------------------------------------------------------------------

Public Const FRAME_MARGIN As Long = 10
Public Const FRAME_SPACING As Long = 12

'----------------------------------------------------------------------------
' Controls
'----------------------------------------------------------------------------

Public Const LABEL_HEIGHT As Long = 18

Public Const TEXTBOX_HEIGHT As Long = 22

Public Const COMBO_HEIGHT As Long = 22

Public Const BUTTON_HEIGHT As Long = 28

Public Const BUTTON_WIDTH As Long = 120

Public Const CHECKBOX_HEIGHT As Long = 18

Public Const LISTBOX_ROW_HEIGHT As Long = 18

'----------------------------------------------------------------------------
' Fonts
'----------------------------------------------------------------------------

Public Const FONT_NAME As String = "Segoe UI"

Public Const FONT_SMALL As Integer = 8

Public Const FONT_NORMAL As Integer = 9

Public Const FONT_HEADER As Integer = 11

Public Const FONT_TITLE As Integer = 14

'----------------------------------------------------------------------------
' Navigation
'----------------------------------------------------------------------------

Public Enum eModule

    mdHome = 0

    mdGeometry = 1

    mdAerodynamics = 2

    mdStability = 3

    mdPerformance = 4

    mdReports = 5

    mdSettings = 6

End Enum

'----------------------------------------------------------------------------
' Stability Methods
'----------------------------------------------------------------------------

Public Enum eCriterion

    crKoch = 1

    crAungier = 2

    crJianLi = 3

    crLieblein = 4

    crDeHaller = 5

End Enum

'----------------------------------------------------------------------------
' Status
'----------------------------------------------------------------------------

Public Enum eStatus

    stReady = 0

    stRunning = 1

    stCompleted = 2

    stError = 3

End Enum

'----------------------------------------------------------------------------
' Default Sheets
'----------------------------------------------------------------------------

Public Const DEFAULT_INPUT_SHEET As String = "Stage Calculation"

Public Const DEFAULT_OUTPUT_SHEET As String = "Results"

'----------------------------------------------------------------------------
' Worksheet Columns
'----------------------------------------------------------------------------

Public Const COL_STAGE As Long = 1

Public Const COL_RADIUS As Long = 2

Public Const COL_FLOW As Long = 3

Public Const COL_BETA1 As Long = 4

Public Const COL_BETA2 As Long = 5

Public Const COL_ALPHA2 As Long = 6

Public Const COL_ALPHA3 As Long = 7

'----------------------------------------------------------------------------
' Messages
'----------------------------------------------------------------------------

Public Const MSG_READY As String = "Ready"

Public Const MSG_RUNNING As String = "Running..."

Public Const MSG_COMPLETE As String = "Calculation completed."

Public Const MSG_ERROR As String = "Calculation failed."

'----------------------------------------------------------------------------
' Buttons
'----------------------------------------------------------------------------

Public Const BTN_RUN As String = "Run Analysis"

Public Const BTN_CLOSE As String = "Close"

Public Const BTN_BROWSE As String = "Browse..."

'----------------------------------------------------------------------------
' Project
'----------------------------------------------------------------------------

Public Const PROJECT_UNNAMED As String = "Untitled Project"

'----------------------------------------------------------------------------
' File Filters
'----------------------------------------------------------------------------

Public Const FILTER_EXCEL As String = _
"Excel Workbook (*.xlsx;*.xlsm),*.xlsx;*.xlsm"

'----------------------------------------------------------------------------
' End of Module
'----------------------------------------------------------------------------
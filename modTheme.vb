Attribute VB_Name = "modTheme"
Option Explicit

'==============================================================================
' Compressor Preliminary Design System
' Theme Module
'==============================================================================

Public Enum eThemeColor

    tcHeader = 1
    tcBackground
    tcPanel
    tcBorder
    tcAccent
    tcText
    tcTextSecondary
    tcSuccess
    tcWarning
    tcDanger
    tcNavigation
    tcNavigationSelected

End Enum

'==============================================================================
' Colors
'==============================================================================

Public Function ThemeColor(ByVal ColorID As eThemeColor) As Long

    Select Case ColorID

        Case tcHeader
            ThemeColor = RGB(31, 78, 121)

        Case tcBackground
            ThemeColor = RGB(242, 244, 247)

        Case tcPanel
            ThemeColor = RGB(255, 255, 255)

        Case tcBorder
            ThemeColor = RGB(214, 214, 214)

        Case tcAccent
            ThemeColor = RGB(0, 120, 215)

        Case tcText
            ThemeColor = RGB(35, 35, 35)

        Case tcTextSecondary
            ThemeColor = RGB(110, 110, 110)

        Case tcSuccess
            ThemeColor = RGB(0, 153, 51)

        Case tcWarning
            ThemeColor = RGB(255, 170, 0)

        Case tcDanger
            ThemeColor = RGB(200, 60, 60)

        Case tcNavigation
            ThemeColor = RGB(236, 239, 241)

        Case tcNavigationSelected
            ThemeColor = RGB(225, 235, 248)

        Case Else
            ThemeColor = vbBlack

    End Select

End Function

'==============================================================================
' UserForm
'==============================================================================

Public Sub ApplyMainTheme(ByVal frm As Object)

    With frm

        .BackColor = ThemeColor(tcBackground)

        .Caption = APP_NAME & "   v" & APP_VERSION

        .Font.Name = FONT_NAME
        .Font.Size = FONT_NORMAL

        .Width = MAIN_WIDTH
        .Height = MAIN_HEIGHT

    End With

End Sub

'==============================================================================
' Frames
'==============================================================================

Public Sub StyleFrame(ByVal fra As MSForms.Frame)

    With fra

        .BackColor = ThemeColor(tcPanel)

        .BorderStyle = fmBorderStyleSingle

        .SpecialEffect = fmSpecialEffectFlat

        .Font.Name = FONT_NAME
        .Font.Size = FONT_NORMAL
        .Font.Bold = True

        .ForeColor = ThemeColor(tcHeader)

    End With

End Sub

'==============================================================================
' Labels
'==============================================================================

Public Sub StyleLabel(ByVal lbl As MSForms.Label)

    With lbl

        .BackStyle = fmBackStyleTransparent

        .Font.Name = FONT_NAME
        .Font.Size = FONT_NORMAL

        .ForeColor = ThemeColor(tcText)

    End With

End Sub

Public Sub StyleHeader(ByVal lbl As MSForms.Label)

    With lbl

        .BackStyle = fmBackStyleTransparent

        .Font.Name = FONT_NAME
        .Font.Size = FONT_TITLE
        .Font.Bold = True

        .ForeColor = ThemeColor(tcHeader)

    End With

End Sub

Public Sub StyleSection(ByVal lbl As MSForms.Label)

    With lbl

        .BackStyle = fmBackStyleTransparent

        .Font.Name = FONT_NAME
        .Font.Size = FONT_HEADER
        .Font.Bold = True

        .ForeColor = ThemeColor(tcHeader)

    End With

End Sub

Public Sub StyleHint(ByVal lbl As MSForms.Label)

    With lbl

        .BackStyle = fmBackStyleTransparent

        .Font.Name = FONT_NAME
        .Font.Size = FONT_SMALL

        .ForeColor = ThemeColor(tcTextSecondary)

    End With

End Sub

'==============================================================================
' TextBox
'==============================================================================

Public Sub StyleTextBox(ByVal txt As MSForms.TextBox)

    With txt

        .BackColor = vbWhite

        .ForeColor = ThemeColor(tcText)

        .BorderStyle = fmBorderStyleSingle

        .SpecialEffect = fmSpecialEffectSunken

        .Font.Name = FONT_NAME
        .Font.Size = FONT_NORMAL

    End With

End Sub

'==============================================================================
' ComboBox
'==============================================================================

Public Sub StyleComboBox(ByVal cmb As MSForms.ComboBox)

    With cmb

        .BackColor = vbWhite

        .ForeColor = ThemeColor(tcText)

        .Style = fmStyleDropDownList

        .Font.Name = FONT_NAME
        .Font.Size = FONT_NORMAL

    End With

End Sub

'==============================================================================
' CheckBox
'==============================================================================

Public Sub StyleCheckBox(ByVal chk As MSForms.CheckBox)

    With chk

        .BackStyle = fmBackStyleTransparent

        .ForeColor = ThemeColor(tcText)

        .Font.Name = FONT_NAME
        .Font.Size = FONT_NORMAL

    End With

End Sub

'==============================================================================
' OptionButton
'==============================================================================

Public Sub StyleOption(ByVal opt As MSForms.OptionButton)

    With opt

        .BackStyle = fmBackStyleTransparent

        .ForeColor = ThemeColor(tcText)

        .Font.Name = FONT_NAME
        .Font.Size = FONT_NORMAL

    End With

End Sub

'==============================================================================
' CommandButton
'==============================================================================

Public Sub StyleButton(ByVal cmd As MSForms.CommandButton)

    With cmd

        .Width = BUTTON_WIDTH
        .Height = BUTTON_HEIGHT

        .Font.Name = FONT_NAME
        .Font.Size = FONT_NORMAL
        .Font.Bold = True

        .TakeFocusOnClick = False

        .BackColor = ThemeColor(tcAccent)

        .ForeColor = vbWhite

    End With

End Sub

Public Sub StyleSecondaryButton(ByVal cmd As MSForms.CommandButton)

    With cmd

        .Width = BUTTON_WIDTH
        .Height = BUTTON_HEIGHT

        .Font.Name = FONT_NAME
        .Font.Size = FONT_NORMAL

        .TakeFocusOnClick = False

        .BackColor = RGB(235, 235, 235)

        .ForeColor = ThemeColor(tcText)

    End With

End Sub

'==============================================================================
' ListBox
'==============================================================================

Public Sub StyleNavigation(ByVal lst As MSForms.ListBox)

    With lst

        .Font.Name = FONT_NAME
        .Font.Size = FONT_NORMAL

        .BorderStyle = fmBorderStyleNone

        .BackColor = ThemeColor(tcNavigation)

        .ForeColor = ThemeColor(tcText)

        .IntegralHeight = False

    End With

End Sub

'==============================================================================
' Status Label
'==============================================================================

Public Sub SetStatus(ByVal lbl As MSForms.Label, _
                     ByVal StatusText As String, _
                     Optional ByVal Status As eStatus = stReady)

    lbl.Caption = StatusText

    Select Case Status

        Case stReady
            lbl.ForeColor = ThemeColor(tcText)

        Case stRunning
            lbl.ForeColor = ThemeColor(tcAccent)

        Case stCompleted
            lbl.ForeColor = ThemeColor(tcSuccess)

        Case stError
            lbl.ForeColor = ThemeColor(tcDanger)

    End Select

End Sub

'==============================================================================
' Window
'==============================================================================

Public Sub CenterForm(ByVal frm As Object)

    frm.StartUpPosition = 1

End Sub

Public Sub DisableInterface(ByVal frm As Object)

    Dim ctrl As Control

    For Each ctrl In frm.Controls

        On Error Resume Next
        ctrl.Enabled = False
        On Error GoTo 0

    Next ctrl

End Sub

Public Sub EnableInterface(ByVal frm As Object)

    Dim ctrl As Control

    For Each ctrl In frm.Controls

        On Error Resume Next
        ctrl.Enabled = True
        On Error GoTo 0

    Next ctrl

End Sub

'==============================================================================
' End of module
'==============================================================================
VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   2505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1995
   Icon            =   "form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "form1.frx":030A
   MousePointer    =   99  'Custom
   ScaleHeight     =   167
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   133
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   165
      Top             =   225
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'=              Agustin Rodriguez             =
'=         virtual_guitar_1@hotmail.com       =
'=    http://www.foreverbahia.com.br/agustin  =
'=  http://www.geocities.com/virtual_quality  =
'==============================================

'THIS PROGRAM USES CLASSES THEN DON'T PRESS THE IDE STOP BUTTON TO PREVENT GPF

'RUN ONLY IN XP

'Press SHIFT + ESC or the Button END to exit

'PNG Class credits and thanks to Apeiron

Option Explicit

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long  ' Declare API
 
Private PNG As LayeredWindow
Private XX As Long
Private YY As Long
Private capture As Integer
Private pt As POINTAPI

Private Sub Form_Load()

  Dim x As String

    Set PNG = New LayeredWindow

    x = Right$("00" + Trim$(Str(Int(Rnd * 11))), 2)

    PNG.MakeTrans App.Path & "\\Drops\Drop " & x & ".png", Me

    Timer1.Interval = Abs(Int(Rnd * Height))

    If Nr = 20 Then
        x = "11"
        PNG.MakeTrans App.Path & "\\Drops\Drop " & x & ".png", Me
        Timer1.Interval = 0

    End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If Button = 1 Then
        XX = x * Screen.TwipsPerPixelX: YY = Y * Screen.TwipsPerPixelY
        capture = True
        ReleaseCapture
        SetCapture Me.hWnd
    End If

    If Tag = 20 Then
        finalize
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If capture Then
        GetCursorPos pt
        Move pt.x * Screen.TwipsPerPixelX - XX, pt.Y * Screen.TwipsPerPixelY - YY
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

    capture = False
    Timer1.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
PNG.UnloadPNGForm
End Sub

Private Sub Timer1_Timer()

  Dim pos As POINTAPI, w As Long, h As Long, x As String, i As Integer, t As Single

    Me.Move Left, Top + 100
    
    If Top > Screen.Height - Height / 2 Then
        sndPlaySound App.Path & "\Splash.wav", 1
        For i = 0 To 7
            x = Right$("00" + Trim$(Str(Int(Rnd * 7))), 2)
            PNG.MakeTrans App.Path & "\\Drops\Splash " & x & ".png", Me
            t = Timer + 0.02
            Do While t > Timer
                DoEvents
            Loop
        Next
        
        Top = -Height
        Timer1.Interval = Int(Rnd * Height)
        PNG.MakeTrans App.Path & "\\Drops\Drop " & x & ".png", Me
    End If
   
    If GetAsyncKeyState(27) And GetAsyncKeyState(16) Then
        finalize
    End If

    If Int(Rnd * 100) = 1 Then Timer1.Interval = 1

End Sub



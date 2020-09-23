VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LCD-Display"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command15 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2160
      TabIndex        =   24
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Multi Count"
      Height          =   1935
      Left            =   2160
      TabIndex        =   18
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton Command14 
         Caption         =   "Stop"
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Start"
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   1575
      End
      Begin VB.PictureBox picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   1200
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   21
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   840
         Picture         =   "Form1.frx":01F6
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   480
         Picture         =   "Form1.frx":03EC
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   120
         Top             =   240
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Count"
      Height          =   1815
      Left            =   0
      TabIndex        =   13
      Top             =   1920
      Width           =   2055
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   240
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Stop"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Start"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   1575
      End
      Begin VB.PictureBox picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   840
         Picture         =   "Form1.frx":05E2
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Click Buttons"
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2055
      Begin VB.PictureBox picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   840
         Picture         =   "Form1.frx":07D8
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command10 
         Caption         =   "0"
         Height          =   495
         Left            =   1560
         TabIndex        =   3
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton Command8 
         Caption         =   "8"
         Height          =   495
         Left            =   1200
         TabIndex        =   5
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Caption         =   "6"
         Height          =   495
         Left            =   840
         TabIndex        =   7
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "4"
         Height          =   495
         Left            =   480
         TabIndex        =   9
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "2"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton Command9 
         Caption         =   "9"
         Height          =   495
         Left            =   1560
         TabIndex        =   4
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         Caption         =   "7"
         Height          =   495
         Left            =   1200
         TabIndex        =   6
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Caption         =   "5"
         Height          =   495
         Left            =   840
         TabIndex        =   8
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "3"
         Height          =   495
         Left            =   480
         TabIndex        =   10
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   525
      Index           =   0
      Left            =   4560
      Picture         =   "Form1.frx":09CE
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   252
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   525
      Index           =   1
      Left            =   4560
      Picture         =   "Form1.frx":19D0
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   252
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Frame Frame4 
      Caption         =   "You can't see this under runtime."
      Height          =   1455
      Left            =   4440
      TabIndex        =   25
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************************
'<Liquit Cristal Display> project by <Peter Hebels> Website "www.grworld.com/megagsite/peterspagina.html  *
'Iam not responsible for any damages may caused by this project                                           *
'PS: I have also included the bitmap files for the numbers so you can use them in your own projects       *
'Commented in English and Dutch.                                                                          *
'**********************************************************************************************************

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const MERGEPAINT = &HBB0226 'For the mask picture.
                                    'Voor het afdek plaatje.

Private Const SRCAND = &H8800C6  'For the front picture.
                                 'Voor het bovenste plaatje.

Private Const SRCCOPY = &HCC0020 'Not used but maybe handy to know.
                                 'Wordt niet gebruikt maar misschien handig om te weten.

Dim I As Integer 'Used to count.
                 'Gebruikt om te tellen.

Dim K As Integer 'Used to count.
Dim NextCount, NextCount1 As Integer 'Used to count.

Private Sub Command1_Click()
picture1(1).Cls 'First clear the picturebox.
                'Eerst de PictureBox leegmaken.

BitBlt picture1(1).hDC, -24, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT 'Draw the mask.
                                                                             'Teken het afdek plaatje.
BitBlt picture1(1).hDC, -24, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND     'Draw the front.
                                                                             'Teken het voorste plaatje.
End Sub

Private Sub Command10_Click()
picture1(1).Cls
BitBlt picture1(1).hDC, -0, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(1).hDC, -0, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND
End Sub

Private Sub Command11_Click()
Timer1.Enabled = True 'Enable the Timer so it will begin to count.
                      'Start de Timer Control om met tellen te beginnen.
End Sub

Private Sub Command12_Click()
Timer1.Enabled = False 'Dissable the timer so count stops.
                       'Stop de Timer Control om met tellen te stoppen.
picture1(0).Cls
BitBlt picture1(0).hDC, 0, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(0).hDC, 0, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND
I = 0 'set count value to zero.
      'Zet de boel terug naar 0.
End Sub

Private Sub Command13_Click()
Timer2.Enabled = True
End Sub

Private Sub Command14_Click()
'This sub will reset the whole timer.
'Deze sub verwijdert alle <tel> gegevens.
Timer2.Enabled = False

picture1(5).Cls
BitBlt picture1(5).hDC, 0, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(5).hDC, 0, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND

picture1(4).Cls
BitBlt picture1(4).hDC, 0, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(4).hDC, 0, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND

picture1(3).Cls
BitBlt picture1(3).hDC, 0, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(3).hDC, 0, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND

K = 0
NextCount = 0
NextCount1 = 0
End Sub

Private Sub Command15_Click()
'Make sure all the timers are stopped.
'Zorg dat alle Timers gestopt zijn.
Timer1.Enabled = False
Timer2.Enabled = False
'Then unload the form.
'Haal de Form weg
Unload Me
'Then end the app.
'Nu kan het programma veilig sluiten.
End
End Sub

'----------------------------------------------------------------------------
'-----------------------Begin For the click buttons--------------------------
Private Sub Command2_Click()
picture1(1).Cls
BitBlt picture1(1).hDC, -48, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(1).hDC, -48, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND
End Sub

Private Sub Command3_Click()
picture1(1).Cls
BitBlt picture1(1).hDC, -72, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(1).hDC, -72, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND
End Sub

Private Sub Command4_Click()
picture1(1).Cls
BitBlt picture1(1).hDC, -96, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(1).hDC, -96, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND
End Sub

Private Sub Command5_Click()
picture1(1).Cls
BitBlt picture1(1).hDC, -120, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(1).hDC, -120, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND
End Sub

Private Sub Command6_Click()
picture1(1).Cls
BitBlt picture1(1).hDC, -144, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(1).hDC, -144, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND
End Sub

Private Sub Command7_Click()
picture1(1).Cls
BitBlt picture1(1).hDC, -168, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(1).hDC, -168, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND
End Sub

Private Sub Command8_Click()
picture1(1).Cls
BitBlt picture1(1).hDC, -192, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(1).hDC, -192, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND
End Sub

Private Sub Command9_Click()
picture1(1).Cls
BitBlt picture1(1).hDC, -216, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(1).hDC, -216, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND
End Sub
'-----------------------End of click buttons--------------------------
'---------------------------------------------------------------------

Private Sub Form_Load()
Me.Width = 4305

'Set values to 0.
'Zet alles weer op nul.
I = 0
K = 0
NextCount = 0

'Put <0> in every LCD-screen
'Zet <0> in elk LCD-Scherm

'----------For the Click-Buttons--------------
BitBlt picture1(1).hDC, 0, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(1).hDC, 0, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND

'----------For all the other count windows----
'You can identify them by the value after every picture call
'Like this <picture1(0).hDC> the (0) is the identifier of the picturebox
'With can be found back by clicking on the control and look at <Properties>

BitBlt picture1(0).hDC, 0, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(0).hDC, 0, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND

BitBlt picture1(5).hDC, 0, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(5).hDC, 0, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND

picture1(5).Cls
BitBlt picture1(5).hDC, 0, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(5).hDC, 0, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND

picture1(4).Cls
BitBlt picture1(4).hDC, 0, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(4).hDC, 0, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND

picture1(3).Cls
BitBlt picture1(3).hDC, 0, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(3).hDC, 0, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND


End Sub


Private Sub Form_Unload(Cancel As Integer)
'Make sure all the timers are stopped.
'Zorg dat alle Timers gestopt zijn.
Timer1.Enabled = False
Timer2.Enabled = False
'Then unload the form.
'Haal de Form weg
Unload Me
'Then end the app.
'Nu kan het programma veilig sluiten.
End
End Sub

Private Sub Timer1_Timer()
'Start counting
'Hervat het tellen

'Simple just slide 24 pixels left.
'Simpel schuif gewoon 24 pixels naar links.
I = I + 24

If I = 240 Then I = 0
picture1(0).Cls
'And show it in the Picture Box
'En laat het zien in de Picture Box
BitBlt picture1(0).hDC, -I, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(0).hDC, -I, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND

End Sub

Private Sub Timer2_Timer()
'The same way as the above sub.
'Ongeveer hetzelfde als de bovenstaande sub.
K = K + 24

'Make the second window show its number.
'Laat het tweede scherm zijn nummer zien.
If K >= 240 Then
K = 0
picture1(4).Cls
NextCount = NextCount + 24
BitBlt picture1(4).hDC, -NextCount, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(4).hDC, -NextCount, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND
End If

'Make the third window show its number.
'Laat het derde scherm zijn nummer zien.
If NextCount >= 240 Then
NextCount = 0
picture1(3).Cls
picture1(4).Cls
NextCount1 = NextCount1 + 24

BitBlt picture1(4).hDC, 0, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(4).hDC, 0, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND

BitBlt picture1(3).hDC, -NextCount1, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(3).hDC, -NextCount1, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND
End If

'Reset the whole thing.
'Zet alles weer op nul.
If NextCount1 = 240 Then
picture1(5).Cls
BitBlt picture1(5).hDC, 0, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(5).hDC, 0, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND

picture1(4).Cls
BitBlt picture1(4).hDC, 0, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(4).hDC, 0, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND

picture1(3).Cls
BitBlt picture1(3).hDC, 0, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(3).hDC, 0, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND
K = 0
NextCount = 0
NextCount1 = 0
End If

'You have to start somewhere.
'Je moet toch ergens beginnen.
picture1(5).Cls
BitBlt picture1(5).hDC, -K, 0, 1000, 100, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1(5).hDC, -K, 0, 1000, 100, Picture3(0).hDC, 0, 0, SRCAND

End Sub

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H0080C0FF&
   Caption         =   "Ruble's Masking Utility"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic4 
      AutoRedraw      =   -1  'True
      Height          =   930
      Left            =   120
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SAVE MASK"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.PictureBox Dest 
      AutoRedraw      =   -1  'True
      Height          =   3855
      Left            =   120
      Picture         =   "Form1.frx":0ABA
      ScaleHeight     =   253
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   349
      TabIndex        =   1
      Top             =   2520
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TEST"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   2040
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.bmp|*.bmp"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MASK"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.PictureBox pic3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   2160
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   3
      Top             =   480
      Width           =   1020
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   120
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   0
      Top             =   480
      Width           =   1020
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Mask"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Original"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'* Made by Ruble
'* ruble_19@yahoo.com
'**********************
'*
'* You may use this code any way you like.
'*
'* Description:
'* This program takes an image and makes a mask image of it
'* It uses black as the masking color
'* Everything that is black in the original will be masked
'*
'* Usage:
'* 1) Double Click on the original picture box to load the picture
'* 2) Click the "MASK" Button
'* 3) Click the "TEST" Button to test if the mask works
'* 4) Click the "STOP" Button to stop the test
'* 5) Click the "SAVE" Button to save the mask image
'* 6) Repeat to make another mask
'*
'* Warning!!!
'* For very big pictures the form will not resize, however it will still work
'******************************************************************



'Define constants for BitBlt
'(You can also use vb constants  Ex. vbSrcCopy)
Const srcand = &H8800C6
Const srccopy = &HCC0020
Const SRCERASE = &H440328
Const srcinvert = &H660046

'This will be the step that determines how fast the test image moves
Dim stp As Integer

Private Sub Command1_Click()
'Change the TEST state
If Timer1.Enabled = True Then
    Timer1.Enabled = False
    Command1.Caption = "TEST"
Else
    Timer1.Enabled = True
    Command1.Caption = "STOP"
End If
End Sub

Private Sub Command2_Click()
pic3.Cls
'Initialize the progress bar
pBar.Value = 0
pBar.Visible = True
pBar.Max = pic1.ScaleWidth

'Loop through each pixel of original image
For i = 0 To pic1.ScaleWidth
    DoEvents
    For k = 0 To pic1.ScaleHeight
        'Get color of the pixel(i,k) of original image
        clr = GetPixel(pic1.hdc, i, k)
        'Check if color is masking color (black)
        If clr = vbBlack Then
            'Set mask image pixel(i,k) to inverse of mask color (white)
            SetPixel pic3.hdc, i, k, vbWhite
        Else
            'Else set to masking color (black)
            SetPixel pic3.hdc, i, k, vbBlack
        End If
    Next k
    pic3.Refresh
    'Increment progress bar
    If pBar.Value < pic1.ScaleWidth Then pBar.Value = pBar.Value + 1
Next i

'Hide progress bar
pBar.Visible = False
End Sub

Private Sub Command3_Click()
'Clear name
CommonDialog1.FileName = ""

'Show Dialog box
CommonDialog1.ShowSave

If CommonDialog1.FileName <> "" Then
    'Save to selected location
    SavePicture pic3.Image, CommonDialog1.FileName
End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub pic1_DblClick()
'Set the mode of the dialog box
CommonDialog1.Flags = cdlOFNCreatePrompt
'Display dialog box
CommonDialog1.Action = 1

'Check if cancel has been clicked
If CommonDialog1.FileName <> "" Then
    pic1.Picture = LoadPicture(CommonDialog1.FileName)
    
    'Resize original and mask picture boxes to fit picture
    pic3.Height = pic1.Height
    pic3.Width = pic1.Width
    pic4.Height = pic1.Height
    pic4.Width = pic1.Width
    pic3.Refresh
    pic4.Refresh
End If
End Sub

Private Sub Timer1_Timer()
    
    'Set step to 2 (Smaller number = faster)
    stp = 2
    'Clear Dest
    Dest.Cls
    'Find middle of Dest picture
    X1 = Int(Dest.ScaleWidth / 2 - pic3.ScaleWidth / 2)
    Y1 = Int(Dest.ScaleHeight / 2 - pic3.ScaleHeight / 2)
    oldX = X1
    oldY = Y1
    
    'Continueous loop until user stops
    Do
        DoEvents
        
        'Increase the loop count
        cont = cont + 1
        
        'If image moves too fast with stp = 1 then you can slow it down with this number (Greater number = slower)
        If cont = 2 Then
            Dest.Cls
            'BitBlt Dest.hdc, oldX, oldY, pic4.ScaleWidth, pic4.ScaleHeight, pic4.hdc, 0, 0, srccopy
            'Dest.Refresh
            'pic4.Refresh
            
            'Check if picture is offscreen
            If Y1 + pic3.ScaleHeight > Dest.ScaleHeight Then
                'Backup the current image at the location with coordinate correction for offTheScreen
                BitBlt pic4.hdc, 0, 0, pic3.ScaleWidth, Y1 - pic3.ScaleHeight, Dest.hdc, X1, Y1, srccopy
            Else
                'Backup the current image at the location
                BitBlt pic4.hdc, 0, 0, pic3.ScaleWidth, pic3.ScaleHeight, Dest.hdc, X1, Y1, srccopy
            End If

            'Draw Mask
            BitBlt Dest.hdc, X1, Y1, pic3.ScaleWidth, pic3.ScaleHeight, pic3.hdc, 0, 0, srcand
            'Draw Original
            BitBlt Dest.hdc, X1, Y1, pic1.ScaleWidth, pic1.ScaleHeight, pic1.hdc, 0, 0, srcinvert
            
            'Save current coordinates
            oldX = X1
            oldY = Y1
            
            'Increment coordinates
            X1 = X1
            Y1 = Y1 + stp
            
            'If image off screen then set Y accordingly
            If Y1 > Dest.ScaleHeight Then
                Y1 = 0 - pic1.ScaleHeight
            End If
            cont = 1
        End If
    Loop Until Timer1.Enabled = False
End Sub

VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Azimuth / Elevation"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   383
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   499
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnReset 
      Cancel          =   -1  'True
      Caption         =   "Reset"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   5340
      Width           =   1875
   End
   Begin VB.CommandButton btnDraw 
      Caption         =   "Draw"
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   4860
      Width           =   1875
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   5460
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "Form1.frx":000C
      Top             =   840
      Width           =   1995
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000005&
      Height          =   795
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   7425
      TabIndex        =   1
      Top             =   0
      Width           =   7485
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2005 Peter Wilson"
         Height          =   195
         Left            =   780
         TabIndex        =   6
         Top             =   420
         Width           =   2190
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Satellite Plotter v1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   780
         TabIndex        =   2
         Top             =   60
         Width           =   2835
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   180
         Picture         =   "Form1.frx":0090
         Top             =   120
         Width           =   480
      End
   End
   Begin SatellitePlotter.usSatellitePlot usSatellitePlot1 
      Height          =   4875
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8599
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnDraw_Click()

    Dim varData     As Variant
    Dim varPair     As Variant
    Dim intN        As Integer
    Dim Elevation   As Single
    Dim Azimuth     As Single
    
    Me.usSatellitePlot1.Reset
    
    varData = Split(Me.Text1.Text, vbCrLf)
        
    For intN = 1 To UBound(varData) - 1
    
        varPair = Split(varData(intN), ",")
        
        Elevation = varPair(0)
        Azimuth = varPair(1)
        
        Call Me.usSatellitePlot1.PlotSatellite(Azimuth, Elevation)
        
    Next intN
    
End Sub

Private Sub btnReset_Click()

    Me.usSatellitePlot1.Reset
    
End Sub

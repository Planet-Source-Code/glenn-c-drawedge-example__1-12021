VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCustom 
      Caption         =   "Custom Edge"
      Height          =   360
      Left            =   1980
      TabIndex        =   17
      Top             =   2145
      Width           =   1110
   End
   Begin VB.CommandButton cmdDrawRing 
      Caption         =   "Draw Ring"
      Height          =   360
      Left            =   1980
      TabIndex        =   16
      Top             =   2625
      Width           =   1110
   End
   Begin VB.CommandButton cmdDrawEdge 
      Caption         =   "&Draw Edge"
      Default         =   -1  'True
      Height          =   360
      Left            =   3840
      TabIndex        =   15
      Top             =   2145
      Width           =   1110
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sides"
      Height          =   1875
      Left            =   3735
      TabIndex        =   2
      Top             =   195
      Width           =   1305
      Begin VB.CheckBox chkSides 
         Caption         =   "Bottom"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   7
         Top             =   1530
         Width           =   855
      End
      Begin VB.CheckBox chkSides 
         Caption         =   "Right"
         Height          =   240
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1209
         Width           =   855
      End
      Begin VB.CheckBox chkSides 
         Caption         =   "Top"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   936
         Width           =   855
      End
      Begin VB.CheckBox chkSides 
         Caption         =   "Left"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   648
         Width           =   855
      End
      Begin VB.CheckBox chkSides 
         Caption         =   "All"
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   330
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Style"
      Height          =   2550
      Left            =   150
      TabIndex        =   1
      Top             =   195
      Width           =   1320
      Begin VB.OptionButton optStyle 
         Caption         =   "Sunken"
         Height          =   210
         Index           =   6
         Left            =   165
         TabIndex        =   14
         Top             =   2190
         Width           =   945
      End
      Begin VB.OptionButton optStyle 
         Caption         =   "Raised"
         Height          =   240
         Index           =   5
         Left            =   165
         TabIndex        =   13
         Top             =   1860
         Width           =   945
      End
      Begin VB.OptionButton optStyle 
         Caption         =   "Mono"
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   12
         Top             =   1575
         Width           =   945
      End
      Begin VB.OptionButton optStyle 
         Caption         =   "Flat"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   11
         Top             =   1290
         Width           =   945
      End
      Begin VB.OptionButton optStyle 
         Caption         =   "Etched"
         Height          =   225
         Index           =   2
         Left            =   165
         TabIndex        =   10
         Top             =   990
         Width           =   945
      End
      Begin VB.OptionButton optStyle 
         Caption         =   "Bump"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   9
         Top             =   705
         Width           =   945
      End
      Begin VB.OptionButton optStyle 
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   165
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   360
      Left            =   3840
      TabIndex        =   0
      Top             =   2625
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   1260
      Left            =   1725
      TabIndex        =   18
      Top             =   285
      Width           =   1785
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' Below are the most commonly used constants.
' Look in the API Viewer for more.

' These constants define the style of border to draw.
Private Const BDR_RAISED = &H5
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKEN = &HA
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const BF_FLAT = &H4000
Private Const BF_MONO = &H8000

' These constants define which sides to draw.
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Dim bdrSides As Long    ' Holds the sides to draw.
Dim bdrStyle As Long    ' Holds the type of border to draw.
Dim rc As RECT          ' Holds the rect to draw in pixels.
Dim tStyle As Integer   ' Holds the style selection.
Dim tSide As Integer    ' Holds the sides selection.


Private Sub Form_Load()
    ' Set defaults
    bdrStyle = 0        ' None
    bdrSides = BF_RECT  ' All
    rc.Right = Me.ScaleWidth
    rc.Bottom = Me.ScaleHeight
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDrawEdge_Click()
    ' Clear the current border.
    Me.Cls
    
    ' Simply call the API.
    DrawEdge Me.hdc, rc, bdrStyle, bdrSides
End Sub

Private Sub cmdDrawRing_Click()
    ' This shows how you can create a custom
    ' border by combining differnt styles
    ' and changing the rect.
    
    ' Clear the border
    Me.Cls
    
    ' Setup the inner border.
    Dim trc As RECT
    trc.Left = 4
    trc.Top = 4
    trc.Right = rc.Right - 4
    trc.Bottom = rc.Bottom - 4
    DrawEdge Me.hdc, trc, BDR_SUNKEN, BF_RECT
    
    ' Now draw the outer border.
    DrawEdge Me.hdc, rc, BDR_RAISED, BF_RECT
End Sub

Private Sub cmdCustom_Click()
    ' Another custom border.
    
    ' Clear the border
    Me.Cls
    
    ' Setup the inner rect
    Dim trc As RECT
    trc.Left = 3
    trc.Top = 3
    trc.Right = rc.Right - 3
    trc.Bottom = rc.Bottom - 3
    
    DrawEdge Me.hdc, trc, BDR_RAISEDINNER Or BDR_SUNKENOUTER, BF_RECT
    DrawEdge Me.hdc, rc, BDR_RAISEDOUTER Or BDR_SUNKENINNER, BF_RECT
End Sub

Private Sub chkSides_Click(Index As Integer)
    CalcSides
End Sub

Private Sub optStyle_Click(Index As Integer)
    Select Case Index
        Case 0 ' None
            bdrStyle = 0
        Case 1 ' Bump
            bdrStyle = BDR_RAISEDOUTER Or BDR_SUNKENINNER
        Case 2 ' Etched
            bdrStyle = BDR_RAISEDINNER Or BDR_SUNKENOUTER
        Case 3 ' Flat
            bdrStyle = BDR_SUNKEN
        Case 4 ' Mono
            bdrStyle = BDR_SUNKEN
        Case 5 ' Raised
            bdrStyle = BDR_RAISED
        Case 6 ' Sunken
            bdrStyle = BDR_SUNKEN
    End Select
    
    ' Because the "Flat" and "Mono" style change
    ' the bdrSides, we have to call this.
    CalcSides
End Sub

Private Sub CalcSides()
    ' This sub gets the value for bdrSides.
    ' I do this here because the "Flat" and "Mono"
    ' styles change this value.
    
    bdrSides = 0 ' reset
    
    If chkSides(0).Value Then
        ' If "All" is selected ignore the others.
        bdrSides = BF_RECT
    Else
        ' Or the sides together.
        If chkSides(1).Value Then bdrSides = BF_LEFT
        If chkSides(2).Value Then bdrSides = bdrSides Or BF_TOP
        If chkSides(3).Value Then bdrSides = bdrSides Or BF_RIGHT
        If chkSides(4).Value Then bdrSides = bdrSides Or BF_BOTTOM
    End If
    
    ' Because the "Flat" and "Mono" styles change the bdrSides value,
    ' we need to check for those styles.
    If optStyle(3).Value Then bdrSides = bdrSides Or BF_FLAT
    If optStyle(4).Value Then bdrSides = bdrSides Or BF_MONO
End Sub

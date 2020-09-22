VERSION 5.00
Object = "*\A..\..\..\..\..\..\DOCUME~1\iRoN\MYDOCU~1\AERONF~1\CARDTR~1\Card ActiveX\Cards.vbp"
Begin VB.Form frmMain 
   BackColor       =   &H00FBAA62&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A Card Trick Game"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "RESET"
      Height          =   375
      Left            =   9120
      TabIndex        =   6
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FBAA62&
      Caption         =   "Guide:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   855
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   8535
      Begin VB.Label lblGuide 
         BackStyle       =   0  'Transparent
         Caption         =   "Please select one card and try to remeber. Select in which row your card belongs."
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   8175
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdSelRow 
      Caption         =   "Row 3"
      Height          =   495
      Index           =   3
      Left            =   9120
      TabIndex        =   3
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdSelRow 
      Caption         =   "Row 2"
      Height          =   495
      Index           =   2
      Left            =   9120
      TabIndex        =   2
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdSelRow 
      Caption         =   "Row 1"
      Height          =   495
      Index           =   1
      Left            =   9120
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin Cards.Card Card 
      Height          =   1335
      Index           =   1
      Left            =   7680
      TabIndex        =   0
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2355
      FaceMode        =   0
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   1815
      Index           =   4
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   5520
      Width           =   8655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   1815
      Index           =   1
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   8655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   1815
      Index           =   0
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   8655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Index           =   3
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   8655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Index           =   2
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Width           =   8655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Index           =   5
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   8655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////
'/////////////////////////////////////
' CARD TRICK GAME
' By Aeron Gregory Murilla
'
' Card activex control
' By Feuchtersoft
'
'/////////////////////////////////////
'/////////////////////////////////////

Dim nNumSelect As Byte
Dim nPtr As Byte

Private Sub cmdReset_Click()
   Form_Load
End Sub

Private Sub cmdSelRow_Click(Index As Integer)
   nNumSelect = nNumSelect + 1
   
   If nNumSelect < 3 Then
   
      Select Case Index
        Case 1:
          SwitchCardValues 2, 1, 3
        Case 2:
          SwitchCardValues 3, 2, 1
        Case 3:
          SwitchCardValues 1, 3, 2
      End Select
   Else
      MsgBox "I will take your card now!"
      Card(Index * 7 - nPtr).FaceMode = Blank
      If MsgBox("Retry again?", vbYesNo) = vbYes Then
         Card(Index * 7 - nPtr).FaceMode = FaceUp
         Form_Load
      Else
         End
      End If
   End If
   
   DispCards
   
   Select Case nNumSelect
       Case 1
          lblGuide.Caption = "Select again a row which your card belongs."
       Case 2
          SwitchCardsRandom
   End Select
   
End Sub

Private Sub SwitchCardsRandom()
   Dim nTmpVal As CardValues
   Dim nTmpTyp As CardTypes
   
   Randomize
   nPtr = Int(7 * Rnd)
   
   For i = 1 To 3
      nTmpVal = Card(i * 7 - 3).Cardvalue
      nTmpTyp = Card(i * 7 - 3).CardType
      
      Card(i * 7 - 3).Cardvalue = Card(i * 7 - nPtr).Cardvalue
      Card(i * 7 - 3).CardType = Card(i * 7 - nPtr).CardType
      Card(i * 7 - 3).RefreshCard
      
      Card(i * 7 - nPtr).Cardvalue = nTmpVal
      Card(i * 7 - nPtr).CardType = nTmpTyp
      Card(i * 7 - nPtr).RefreshCard
   Next i
   
End Sub


Public Sub SwitchCardValues(row1 As Integer, row2 As Integer, row3 As Integer)
  For i = 1 To 7
    'Switch Card Type
    CardDeck(i).cType = Card(row1 * 7 - 7 + i).CardType
    CardDeck(7 + i).cType = Card(row2 * 7 - 7 + i).CardType
    CardDeck(14 + i).cType = Card(row3 * 7 - 7 + i).CardType
    'Switch Card Values
    CardDeck(i).cValue = Card(row1 * 7 - 7 + i).Cardvalue
    CardDeck(7 + i).cValue = Card(row2 * 7 - 7 + i).Cardvalue
    CardDeck(14 + i).cValue = Card(row3 * 7 - 7 + i).Cardvalue
  Next i
End Sub

Private Sub Form_Load()
  nNumSelect = 0
  Shuffle
  DispCards
  MsgBox "Follow the guide carefully!"
End Sub

Public Sub DispCards()
On Error Resume Next 'Bypass the already loaded card
  'load card and display
  For i = 1 To 21
    'load card
    Load Card(i)
    Card(i).Visible = True
  Next i
  
  'put card values
  For y = 1 To 7
    For x = 1 To 3
      Card((x * 7 - 7) + y).CardType = CardDeck((y * 3 - 3) + x).cType
      Card((x * 7 - 7) + y).Cardvalue = CardDeck((y * 3 - 3) + x).cValue
    Next x
  Next y
  
  'set position
  For i = 1 To 7
    'X position
    Card(i).Left = Card(i - 1).Left - 1200
    Card(i + 7).Left = Card(i).Left
    Card(i + 14).Left = Card(i).Left
    'Y POsition
    Card(i + 7).Top = Card(i).Top + 2160
    Card(i + 14).Top = Card(i).Top + 4320
  Next i
  
End Sub


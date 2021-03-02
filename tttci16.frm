VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "X's  And O's By BBanfield"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   8850
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Play Computer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5760
      TabIndex        =   20
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton oneplayer 
         Caption         =   "Play Computer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   21
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton endgame 
      Caption         =   "End Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   19
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton zerooutscores 
      Caption         =   "Zero out scores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   18
      Top             =   6360
      Width           =   2895
   End
   Begin VB.CommandButton startnewgame 
      Caption         =   "Start Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   17
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "ScoreBoard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   5760
      TabIndex        =   1
      Top             =   3840
      Width           =   2655
      Begin VB.Label Olabel 
         Alignment       =   2  'Center
         Caption         =   " O    ="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Xlabel 
         Alignment       =   2  'Center
         Caption         =   "X    - "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label oscore 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   5
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label xscore 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "First to Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   5760
      TabIndex        =   0
      Top             =   1440
      Width           =   2655
      Begin VB.OptionButton Option2 
         Caption         =   "O-Player2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "X-Player1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2295
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   735
      Left            =   600
      TabIndex        =   22
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Line winline 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   7
      Visible         =   0   'False
      X1              =   1080
      X2              =   4560
      Y1              =   3960
      Y2              =   720
   End
   Begin VB.Line winline 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   6
      Visible         =   0   'False
      X1              =   1080
      X2              =   4560
      Y1              =   720
      Y2              =   3960
   End
   Begin VB.Line winline 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   5
      Visible         =   0   'False
      X1              =   4080
      X2              =   4080
      Y1              =   720
      Y2              =   3960
   End
   Begin VB.Line winline 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   4
      Visible         =   0   'False
      X1              =   2880
      X2              =   2880
      Y1              =   720
      Y2              =   3960
   End
   Begin VB.Line winline 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   3
      Visible         =   0   'False
      X1              =   1680
      X2              =   1680
      Y1              =   720
      Y2              =   3960
   End
   Begin VB.Line winline 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   2
      Visible         =   0   'False
      X1              =   1080
      X2              =   4560
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line winline 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   1
      Visible         =   0   'False
      X1              =   1080
      X2              =   4560
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line winline 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   1080
      X2              =   4560
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label square 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   8
      Left            =   3600
      TabIndex        =   16
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label square 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   7
      Left            =   2400
      TabIndex        =   15
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label square 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   6
      Left            =   1200
      TabIndex        =   14
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label square 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   5
      Left            =   3600
      TabIndex        =   13
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label square 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   2400
      TabIndex        =   12
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label square 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   1200
      TabIndex        =   11
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label square 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   3600
      TabIndex        =   10
      Top             =   840
      Width           =   855
   End
   Begin VB.Label square 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   2400
      TabIndex        =   9
      Top             =   840
      Width           =   855
   End
   Begin VB.Label square 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   1200
      TabIndex        =   8
      Top             =   840
      Width           =   855
   End
   Begin VB.Line Line4 
      X1              =   1200
      X2              =   4440
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line3 
      X1              =   1200
      X2              =   4440
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line2 
      X1              =   3480
      X2              =   3480
      Y1              =   840
      Y2              =   3840
   End
   Begin VB.Line Line1 
      X1              =   2160
      X2              =   2160
      Y1              =   840
      Y2              =   3840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit
Dim compmoved As Integer, winner As Boolean, playcomputer As Boolean
Dim xtotal As Integer, ototal As Integer, checkplayermove As Integer
Dim startarray(10) As Integer, finisharray(10) As Integer, stepsarray(10) As Integer
Dim locateclick As Integer, start As Integer, finish As Integer, steps As Integer, subtest As Integer
Dim clear_squares As Integer, redlineoff As Integer
Dim total As Integer, cap As String, xscore1 As Integer, oscore1 As Integer, totaltest As Integer

Private Sub endgame_Click()
Unload Form1
End
End Sub

Private Sub Form_Load()
Call larray
cap = ""
xscore1 = 0
oscore1 = 0
total = 0
xtotal = 0
ototal = 0
End Sub


Private Sub oneplayer_Click()
playcomputer = True
Option1.Caption = "X-Computer"
If Option1 = True Then
         cap = "X"
      Call computer_first
End If
End Sub

Private Sub square_Click(Index As Integer)

compmoved = 0
   'SET STARTING CAPTION X OR O
If Option1 = True Then

   cap = "X"
Else
   cap = "O"
   
End If
checkplayermove = Index
' DETERMINE WHAT SQUARE HAS BEEN CLICKED USING THE INDEX
For locateclick = 0 To 8

  If locateclick = Index Then
  
    If square(locateclick).Caption = "" Then
    
        square(locateclick).Caption = cap
        
    End If
  End If
Next locateclick

'CALL SUB TO TEST IF THERE IS A WINNER
  Call test_for_winner
  
  'ROTATE PLAYERS
If playcomputer = False Then
  If Option1 = True Then

   Option1 = False
   Option2 = True
 Else
     If Option2 = True Then
     
        Option2 = False
        Option1 = True
 End If
    End If
End If

 If playcomputer = True Then
  Label1.Caption = "test"
    cap = "X"
    Call computer_second
 End If
End Sub

Private Sub test_for_winner()

For totaltest = 0 To 7
 
 start = startarray(totaltest)
 finish = finisharray(totaltest)
 steps = stepsarray(totaltest)
 
 total = 0
 
   If Option1 = True Then
      cap = "X"
      
   Else
      cap = "O"
      
   End If
      
     For subtest = start To finish Step steps
         If square(subtest).Caption = cap Then
            total = total + 1
              If total = 3 Then
              Call declare_winner
              End If
         End If
     Next subtest
      
Next totaltest

End Sub

Private Sub declare_winner()
'UPDATE SCORE
If Option1 = True Then
   xscore1 = xscore + 1
   xscore.Caption = xscore1
End If

If Option2 = True Then
  oscore1 = oscore1 + 1
  oscore.Caption = oscore1
End If
'TURN ON REDLINE TO SHOW WINNER

winline(totaltest).Visible = True
Beep
redlineoff = totaltest
winner = True
End Sub

Private Sub startnewgame_Click()
'CLEAR SCREEN FOR NEXT GAME
For clear_squares = 0 To 8
square(clear_squares).Caption = ""
Next clear_squares
winline(redlineoff).Visible = False
xtotal = 0
ototal = 0
compmoved = 0
cap = ""
checkplayermove = 0
total = 0
winner = False
End Sub

Private Sub zerooutscores_Click()
'CLEAR SCORES FROM SCREEN AND VARIABLES
xscore.Caption = 0
oscore.Caption = 0
xscore1 = 0
oscore1 = 0
xtotal = 0
ototal = 0
End Sub
Private Sub computer_first()
Option1 = True
compmoved = 0
If square(4).Caption = "" Then    'Capture centre square
   square(4).Caption = cap
End If
Option2 = True
compmoved = 0
End Sub
Private Sub larray()
For totaltest = 0 To 7
  Select Case totaltest
    Case 0
      start = 0
      finish = 2
      steps = 1
      
    Case 1
      start = 3
      finish = 5
      steps = 1
    Case 2
      start = 6
      finish = 8
      steps = 1
      
    Case 3
      start = 0
      finish = 6
      steps = 3
      
    Case 4
      start = 1
      finish = 7
      steps = 3
      
    Case 5
      start = 2
      finish = 8
      steps = 3
      
    Case 6
      start = 0
      finish = 8
      steps = 4
      
    Case 7
      start = 2
      finish = 6
      steps = 2
      
 End Select
 startarray(totaltest) = start
 finisharray(totaltest) = finish
 stepsarray(totaltest) = steps
 Next totaltest
End Sub

Private Sub computer_second()
Option1 = True
compmoved = 0
 If square(4).Caption = "" Then   'CAPTURE CENTRE SQUARE
    square(4).Caption = "X"
    compmoved = 1
 End If
                                  'BLOCKING MOVE
 If compmoved = 0 Then
        For totaltest = 0 To 7
            start = startarray(totaltest)
            finish = finisharray(totaltest)
            steps = stepsarray(totaltest)
            ototal = 0
         For subtest = start To finish Step steps
             If square(subtest).Caption = "O" Then
                   ototal = ototal + 1
             End If
             
               If ototal > 1 Then
                  If square(start).Caption = "O" And square(start + steps).Caption = "O" And square(finish).Caption = "" Then
                     square(finish).Caption = "X"
                     compmoved = 1
                     Exit For
                   Else
                      If square(start).Caption = "" And square(start + steps).Caption = "O" And square(finish).Caption = "O" Then
                      square(start).Caption = "X"
                      compmoved = 1
                      Exit For
                   Else
                     If square(start).Caption = "O" And square(start + steps).Caption = "" And square(finish).Caption = "O" Then
                      square(start + steps).Caption = "X"
                      compmoved = 1
                      Exit For
                                    
                     
                End If
                End If
                 End If
                  End If
                   
                   
                  
                  Next subtest
     Next totaltest
  End If
      

     ' CAPTURE CORNER SQUARE
If compmoved = 0 Then
  If checkplayermove = 0 And square(6).Caption = "" Then
   square(6).Caption = cap
   compmoved = 1
   Else
        
   If checkplayermove = 6 And square(8).Caption = "" Then
      square(8).Caption = cap
      compmoved = 1
     Else
       If checkplayermove = 8 And square(2).Caption = "" Then
          square(2).Caption = cap
          compmoved = 1
      Else
        If checkplayermove = 2 And square(0).Caption = "" Then
           square(0).Caption = cap
           compmoved = 1
End If
  End If
   End If
    End If
End If
                                       'ATTEMPT TO GET 3 X IN A ROW
      If compmoved = 0 Then
        For totaltest = 0 To 7
            start = startarray(totaltest)
            finish = finisharray(totaltest)
            steps = stepsarray(totaltest)
            xtotal = 0
         For subtest = start To finish Step steps
             If square(subtest).Caption = cap Then
                   xtotal = xtotal + 1
             End If
               If xtotal > 1 Then
                  If square(start).Caption = cap And square(start + steps).Caption = cap And square(finish).Caption = "" Then
                     square(finish).Caption = cap
                     compmoved = 1
                     Call test_for_winner
                     Exit For
                   Else
                      If square(start).Caption = "" And square(start + steps).Caption = cap And square(finish).Caption = cap Then
                      square(start).Caption = cap
                      compmoved = 1
                      Call test_for_winner
                      Exit For
                   Else
                     If square(start).Caption = cap And square(start + steps).Caption = "" And square(finish).Caption = cap Then
                      square(start + steps).Caption = cap
                      compmoved = 1
                      Call test_for_winner
                      Exit For
                   Else
                       If square(start).Caption = cap And square(start + steps).Caption = cap And square(finish).Caption = cap Then
                      square(start + steps).Caption = cap
                      compmoved = 1
                      Call test_for_winner
                      Exit For
               End If
                End If
                 End If
                  End If
                   End If
                   
                 
                  Next subtest
                  If winner = True Then
                   Exit Sub
                  End If
      Next totaltest
    End If
    
                                   'IF PLAYER MAKES A WEAK MOVE TRY TO CAPTURE ANOTHER CORNER
  
     If compmoved = 0 Then
          
      If checkplayermove Mod 2 <> 0 Then
         If square(0).Caption = "" Then
            square(0).Caption = cap
            compmoved = 1
      Else
           If square(2).Caption = "" Then
              square(2).Caption = cap
              compmoved = 1
       Else
            If square(6).Caption = "" Then
               square(6).Caption = cap
               compmoved = 1
         Else
            If square(8).Caption = "" Then
               square(8).Caption = cap
               compmoved = 1
       End If
   End If
    End If
     End If
      End If
                                        'IF NOTHING ELSE WORKS CAPTURE THE FIRST EMPTY SQUARE
 End If
     If compmoved = 0 Then
        For totaltest = 0 To 8
           If square(totaltest).Caption = "" Then
              square(totaltest).Caption = cap
            Exit For
           End If
       Next totaltest
    End If
     
 
Option2 = True
compmoved = 0

End Sub

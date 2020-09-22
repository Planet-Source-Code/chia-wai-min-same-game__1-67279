VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Same"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   2520
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   2520
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picScore 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   2460
      TabIndex        =   1
      Top             =   840
      Width           =   2520
   End
   Begin VB.Label lblBox 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuMenuNewGame 
         Caption         =   "New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuMenuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenuUndo 
         Caption         =   "Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuMenuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenuViewScore 
         Caption         =   "View Score"
      End
      Begin VB.Menu mnuMenuClearScore 
         Caption         =   "Clear Score"
      End
      Begin VB.Menu mnuMenuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Option"
      Begin VB.Menu mnuOptionSize 
         Caption         =   "Size"
         Begin VB.Menu mnuOptionSizeNormal 
            Caption         =   "Normal Size"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuOptionSizeLarge 
            Caption         =   "Large Size"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name : Chia Wai Min
'Education Level : Higher Diploma First Semester
'Major Subject : Computer System
'Date Started : 10/02/2003 (Monday)
'Date Finished : 12/02/2003 (Wednesday)
'Version : 1.1

'Prevent Undeclared Variables
    Option Explicit

'Constant Variables
    Private Const SmallRowBox = 20
    Private Const BigRowBox = 30
    Private Const Color1 = vbRed
    Private Const Color2 = vbYellow
    Private Const Color3 = vbGreen
    Private Const Color4 = vbBlue
    Private Const MaxChar = 4
    Private Const StartASCII = 65
    Private Const BlankText = ""
    Private Const BoxVisible = True
    Private Const ScorePower = 2
    Private Const MaxScoreRow = 10
    Private Const MaxScoreRecords = 3

'Internal Process Variables
    Dim PreIndex As Integer
    Dim SelectedBox As Integer
    Dim TotalCount As Integer
    Dim Score As Double
    Dim LastScore As Double

'Array Variables
    Dim ScoreBox(1 To MaxScoreRow, 1 To MaxScoreRecords) As String * 15

Private Function AddNewRecord(Playername As String, Score As Double)
    '**Explaination : Add A New Record To The Score List**

    'Local Variables
    Dim x As Integer, y As Integer
    
    'If Found Greater Than Score, Replace and Move Other Record To Next
    For x = 1 To MaxScoreRow
        If Score >= Val(ScoreBox(x, 2)) Then
            For y = MaxScoreRow To x + 1 Step -1
                ScoreBox(y, 1) = ScoreBox(y - 1, 1)
                ScoreBox(y, 2) = ScoreBox(y - 1, 2)
                ScoreBox(y, 3) = ScoreBox(y - 1, 3)
            Next
            
            ScoreBox(x, 1) = Playername
            ScoreBox(x, 2) = Score
            ScoreBox(x, 3) = Format(Date, "dd/mm/yyyy")
            Exit For
        End If
    Next
    
    'Store Last / Smallest Score
    LastScore = Val(ScoreBox(MaxScoreRow, 2))
    
    'Store To Registry
    Call SaveScoreList
End Function

Private Function SaveScoreList()
    'Explaination : Save Local Score List To Registry

    'Local Variables
    Dim x As Integer
    
    'Store Local Score List To Registry
    For x = 1 To MaxScoreRow
        SaveSetting "Same", "Name", x, ScoreBox(x, 1)
        SaveSetting "Same", "Score", x, ScoreBox(x, 2)
        SaveSetting "Same", "Date", x, ScoreBox(x, 3)
    Next
    
    'Store Last / Smallest Score
    LastScore = Val(ScoreBox(MaxScoreRow, 2))
End Function

Private Function GetScoreList()
    '**Explaination : Get Score From Registry and Store In Local Score List

    'Local Variables
    Dim x As Integer
    
    'Retrieve Score and Put In Local Array Score List
    For x = 1 To MaxScoreRow
        ScoreBox(x, 1) = GetSetting("Same", "Name", x, "")
        ScoreBox(x, 2) = GetSetting("Same", "Score", x, "")
        ScoreBox(x, 3) = GetSetting("Same", "Date", x, "")
    Next
    
    'Store Last / Smallest Score
    LastScore = Val(ScoreBox(MaxScoreRow, 2))
End Function

Private Function NewGame(Small As Boolean)
    '**Explaination : Start New Game

    'Local Variables
    Dim x As Integer, y As Integer, z As Integer
    
    'Select Type Of Game (Normal / Large)
    SelectedBox = IIf(Small = True, SmallRowBox, BigRowBox)
    
    'Disable Undo In Menu
    mnuMenuUndo.Enabled = False
    
    'Unload Boxes Object
    For x = lblBox.LBound + 1 To lblBox.UBound
        Unload lblBox(x)
    Next
    
    'Prevent Same Number Random
    Call Randomize
    
    'Load Box Object and Random Text For Each Box
    For y = 1 To (SelectedBox / 2)
        z = z + 1
        Load lblBox(z)
        lblBox(z).Top = lblBox(z).Top + ((y - 1) * lblBox(0).Height)
        lblBox(z).Left = 0
        lblBox(z).Caption = Chr(StartASCII + Int(Rnd * MaxChar))
        lblBox(z).ForeColor = vbBlack
        lblBox(z).BackColor = getColor(lblBox(z).Caption)
        lblBox(z).Visible = True
        
        For x = 2 To SelectedBox
            z = z + 1
            Load lblBox(z)
            lblBox(z).Top = lblBox(z - 1).Top
            lblBox(z).Left = lblBox(z - 1).Left + lblBox(z).Width
            lblBox(z).Caption = Chr(StartASCII + Int(Rnd * MaxChar))
            lblBox(z).ForeColor = vbBlack
            lblBox(z).BackColor = getColor(lblBox(z).Caption)
            lblBox(z).Visible = True
        Next
    Next
    
    'Set The Form Width and Height
    Me.Width = (lblBox(0).Width) * (SelectedBox + 0.2)
    Me.Height = (lblBox(0).Height) * ((SelectedBox + 4.2) / 2)
    
    'Set The Form To Center Screen
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    'Clear Previous Selected Index
    PreIndex = 0
    
    'Clear The Player Score
    Score = 0
    
    'Print Score On Bottom Label
    Call PrintScore(False)
End Function

Private Function getColor(BoxText As String) As ColorConstants
    'Explaination : Return Different Color For Different Text
    Select Case BoxText
        Case Chr(StartASCII): getColor = Color1
        Case Chr(StartASCII + 1): getColor = Color2
        Case Chr(StartASCII + 2): getColor = Color3
        Case Chr(StartASCII + 3): getColor = Color4
    End Select
End Function
    
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Explaination : When Pressed Escape Key , Exit The Program
    If KeyCode = vbKeyEscape Then mnuMenuExit_Click
End Sub

Private Sub Form_Load()
    'Explaination : Intiation Commands
    Call GetScoreList
    Call NewGame(True)
End Sub

Private Sub lblBox_DblClick(Index As Integer)
    'Explaination : Box Double Click Handle
    If lblBox(Index).BackColor = vbWhite Then CancelBox
End Sub

Private Sub lblBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    'Explaination : Box Click Handle
    
    'Left Click
    If Button = 1 Then
        'Check If The Box Highlighted
        If lblBox(Index).BackColor <> vbWhite And lblBox(Index).Caption <> BlankText Then
            'Cancel Hightlight For Previous Box
            If PreIndex > 0 Then Call CheckBox(PreIndex, False)
            
            'Set Total Highlighted Box to 0
            TotalCount = 0
            
            'Highlighted Near Same Text Boxes
            Call CheckBox(Index, True)
            
            'Check If Only One Highlighted
            If TotalCount <= 1 Then
                'Cancel Highlight
                Call Highlight(Index, False)
                
                'Clear Previous Highlighted Box Index
                PreIndex = 0
            Else
                'More Than One Highlighted
                
                'Set Current Index To Previous Index
                PreIndex = Index
                
                'Print Box and Score Status To Bottom Label
                Call PrintScore(True)
            End If
        End If
    'Right Click
    ElseIf Button = 2 And lblBox(Index).BackColor = vbWhite Then
        'Cancel Highlighted
        Call CheckBox(Index, False)
        
        'Print Box and Score Status To Bottom Label
        Call PrintScore(False)
        
        'Clear Previous Highlighted Box Index
        PreIndex = 0
    End If
End Sub

Private Function CancelBox()
    'Explaination : Cancel Box For Double Click

    'Local Variables
    Dim x As Integer
        
    'Store Current Box Display For Undo Use
    Call Undo("Store")
    
    'Add Player Score
    Score = Score + TotalCount ^ ScorePower
    
    'Check For All Highlighted Boxes
    For x = lblBox.LBound + 1 To lblBox.UBound
        If lblBox(x).BackColor = vbWhite And _
            lblBox(x).Caption <> BlankText Then
                Do While lblBox(x).BackColor = vbWhite
                    Call DropBox(x)   'Arrange Columns
                    Call ArrangeColumn         'Arrange Rows
                Loop
        End If
    Next
    
    'Print Box and Score Status To Bottom Label
    Call PrintScore(False)
    
    'Clear Previous Highlighted Box Index
    PreIndex = 0
    
    'Check If The No Box Available To Double Click
    Call CheckGameOver
End Function

Private Function CheckGameOver()
    'Explaination : To Check For No Boxes Available To Double Click

    'Local Variables
    Dim x As Integer, Total As Integer
    
    'Start From First Box To Last Box
    For x = lblBox.LBound + 1 To lblBox.UBound
        'Set Total Highlighed Box To 0
        TotalCount = 0
        
        'If The Boxed Is Clickable
        If lblBox(x).Caption <> BlankText Then
            
            'Highlight Box
            Call CheckBox(x, True)
            
            'Get Total Highlighted Boxes
            Total = TotalCount
            
            'Cancel 'Highlighted Box
            Call CheckBox(x, False)
            
            'If Still Available The Exit This Function
            If Total > 1 Then Exit Function
        ElseIf x = lblBox.UBound Then
            'Finished Check
            
            'Display Message
            MsgBox "Congratulation, you have won the game!", vbInformation, "Congratulation"
            
            'If Player Score Greater Than Score List Last / Smallest Score
            If Val(Score) >= Val(LastScore) Then
                'Request For Player Name and Add New Record
                Call AddNewRecord(InputBox("Your name will placed in top 10 records" & vbCrLf & "Please insert your name...", "Top 10 Records"), Score)
                
                'Display All Score
                Call mnuMenuViewScore_Click
            End If
            
            'Start New Game
            Call mnuMenuNewGame_Click
            Exit Function
        End If
    Next
End Function

Private Function PrintScore(Active As Boolean)
    'Explaination : To Print Boxes Status and Player Score In The Bottom Label
    
    picScore.Cls
    picScore.ForeColor = vbWhite
    picScore.FontBold = True
    picScore.Print vbTab & "MARK   :   " & IIf(Active = True, TotalCount, "0") & vbTab & "( POINT   :   " & IIf(Active = True, TotalCount ^ ScorePower, "0") & " )"
    picScore.CurrentY = 0
    picScore.CurrentX = picScore.Width - 3000
    picScore.Print "SCORE   :   " & Score
End Function

Private Function DropBox(Index As Integer)
    'Explaination : Arrange Column When Boxes Cancelled
    Do While Index >= SelectedBox + 1
    
        If lblBox(Index - SelectedBox).Caption = BlankText Then
            Call KillBox(Index)
        Else
            lblBox(Index).ForeColor = vbBlack
            lblBox(Index).Caption = lblBox(Index - SelectedBox).Caption
            lblBox(Index).BackColor = getColor(lblBox(Index).Caption)
        End If
        Index = Index - SelectedBox
    Loop
    
    Call KillBox(Index)
End Function

Private Function KillBox(Index As Integer)
    'Explaination : Disable The Box
    lblBox(Index).Caption = BlankText
    lblBox(Index).ForeColor = vbWhite
    lblBox(Index).BackColor = vbBlack
    lblBox(Index).Visible = BoxVisible
End Function

Private Function ArrangeColumn()
    'Explaination : When Entire Column Blank Found, Move Other Back Column To Front
    
    'Local Variables
    Dim x As Integer, y As Integer, z As Integer
        
    'Box Index For Inhenritate
    Dim Index As Integer
    
    'Move Column
    For x = 1 To SelectedBox
        If lblBox(SelectedBox * SelectedBox / 2 - SelectedBox + x).Caption = BlankText Then
            For y = x To SelectedBox - 1
                For z = 1 To SelectedBox / 2
                    Index = ((z - 1) * SelectedBox + y)
                    lblBox(Index).Caption = lblBox(Index + 1).Caption
                    lblBox(Index).ForeColor = lblBox(Index + 1).ForeColor
                    lblBox(Index).BackColor = lblBox(Index + 1).BackColor
                    lblBox(Index).Visible = lblBox(Index + 1).Visible
                Next
            Next
                    
            'Disable Entire Last Column
            For y = 1 To SelectedBox / 2
                KillBox ((y - 1) * SelectedBox + SelectedBox)
            Next
                    
            'Quit This Function
            Exit Function
        End If
    Next
End Function

Private Function CheckBox(Index As Integer, Active As Boolean)
    'Explaination : To Made Sure The Box Is In Same Row Or Column
    Call Highlight(Index, Active)
    If ((Index - 1) Mod SelectedBox <> 0) Then Call CheckOther(Index - 1, Index, Active)
    If ((Index) Mod SelectedBox) <> 0 Then Call CheckOther(Index + 1, Index, Active)
    If Index + SelectedBox <= SelectedBox * SelectedBox / 2 Then Call CheckOther(Index + SelectedBox, Index, Active)
    If Index - SelectedBox >= 0 Then Call CheckOther(Index - SelectedBox, Index, Active)
End Function

Private Function CheckOther(Index As Integer, Index2 As Integer, Active As Boolean)
    'Explaination : Check and Hightligh Near Same Text Box
    If Active = True Then
        If lblBox(Index).BackColor <> vbWhite And _
                lblBox(Index).Caption = lblBox(Index2).Caption Then Call CheckBox(Index, Active)
    Else
        If lblBox(Index).BackColor = vbWhite And _
            lblBox(Index).Caption = lblBox(Index2).Caption Then Call CheckBox(Index, Active)
    End If
End Function

Private Function Highlight(Index As Integer, Active As Boolean)
    'Explaination : To Highlight Box
    If Active = True Then
        'Highlight Box
        TotalCount = TotalCount + 1
        lblBox(Index).ForeColor = lblBox(Index).BackColor
        lblBox(Index).BackColor = vbWhite
    Else
        'Cancel Highlight Box
        lblBox(Index).BackColor = lblBox(Index).ForeColor
        lblBox(Index).ForeColor = vbBlack
    End If
End Function

Private Sub mnuAbout_Click()
    'Explaination : Prompt The Game Designer Information
    MsgBox "Presented by: Chia Wai Min" & vbCrLf & "Email: cwaimin@hotmail.com", vbInformation, "About"
End Sub

Private Sub mnuMenuClearScore_Click()
    'Explaination : To Clear Score List
    
    'Local Variables
    Dim x As Integer, y As Integer
    
    'Clear Local Array Score List
    For x = 1 To MaxScoreRow
        For y = 1 To MaxScoreRecords
            ScoreBox(x, y) = ""
        Next
    Next
    
    'Update Registry
    Call SaveScoreList
    
    'Prompt Message
    MsgBox "Score List had been successfully clear!", vbInformation, "Clear Score"
End Sub

Private Sub mnuMenuExit_Click()
    'Explaination : End The Game
    Unload Me
End Sub

Private Sub mnuMenuNewGame_Click()
    'Explaination : Start New Game According Settings (Normal or Large Game)
    NewGame IIf(mnuOptionSizeNormal.Checked = True, True, False)
End Sub

Private Function LimitText(Text As String, Total As Integer) As String
    'Explantion : Return Format Fixed Length
    If Len(Text) < Total Then
        LimitText = Text & String(Total - Len(Text), " ")
    Else
        LimitText = Left(Text, Total)
    End If
End Function

Private Sub mnuMenuUndo_Click()
    'Explaination : Undo Previous Box Model
    Call Undo("Restore")
End Sub

Private Sub mnuMenuViewScore_Click()
    'Explantion : Generate Score Box In Message Box
    Dim DisplayRecords As String, x As Integer
    DisplayRecords = "No." & vbTab & "Name" & vbTab & vbTab & "Score" & vbTab & "Date"
    For x = 1 To MaxScoreRow
        DisplayRecords = DisplayRecords & vbCrLf & x & vbTab & _
            LimitText(ScoreBox(x, 1), 15) & vbTab & _
            LimitText(ScoreBox(x, 2), 10) & vbTab & _
            ScoreBox(x, 3)
    Next
    MsgBox DisplayRecords, vbInformation, "Score"
End Sub

Private Sub mnuOptionSizeLarge_Click()
    'Explaination : Large Size Selected
    mnuOptionSizeLarge.Checked = True
    mnuOptionSizeNormal.Checked = False
    
    'Start New Game
    Call mnuMenuNewGame_Click
End Sub

Private Sub mnuOptionSizeNormal_Click()
    'Explaination : Normal Size Selected
    mnuOptionSizeNormal.Checked = True
    mnuOptionSizeLarge.Checked = False
    
    'Start New Game
    Call mnuMenuNewGame_Click
End Sub


Private Function Undo(Action As String)
    'Explaination : For Player To Undo
    
    'Local Variables
    Dim x As Integer
    
    'Static Variables
    Static Box(1 To BigRowBox * BigRowBox / 2) As String
    Static PreScore As Double

    Select Case Action
        'Store Current Box Status
        Case "Store"
            PreScore = Score
            For x = lblBox.LBound + 1 To lblBox.UBound
                 Box(x) = IIf(lblBox(x).Caption <> BlankText, lblBox(x).Caption, BlankText)
            Next
            mnuMenuUndo.Enabled = True
        'Restore Previous Box Status
        Case "Restore"
            If mnuMenuUndo.Enabled = True Then
                Score = PreScore
                For x = lblBox.LBound + 1 To lblBox.UBound
                    lblBox(x).Caption = Box(x)
                    lblBox(x).ForeColor = vbBlack
                    lblBox(x).BackColor = getColor(lblBox(x).Caption)
                Next
                mnuMenuUndo.Enabled = False
                Call PrintScore(False)
            End If
    End Select
End Function

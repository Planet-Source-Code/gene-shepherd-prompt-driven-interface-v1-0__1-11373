VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2835
   ClientLeft      =   3315
   ClientTop       =   4530
   ClientWidth     =   7905
   BeginProperty Font 
      Name            =   "OCR A Extended"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Console 
      BackColor       =   &H80000006&
      ForeColor       =   &H80000005&
      Height          =   2535
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Main.frx":0000
      Top             =   360
      Width           =   7935
   End
   Begin VB.Label Label1 
      Caption         =   "Hello...Welocome to OddBall's Prompt Driven Interface example...            created, September 6, 2000"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim User As String

Private Sub Console_KeyPress(KeyAscii As Integer)
Dim Command As String, CommandUndec As String, ParamOne As String, ParamTwo As String
Dim CommandStrt As Integer, CommandEnd As Integer, ParLen As Integer, CoorOne As Integer, CoorTwo As Integer
Dim Com As Boolean

Com = False

If User$ = "" Then
    User$ = "Anonymous"
End If

If KeyAscii% = 8 Then
    If Right(Console.Text, 2) = ">>" Then
        KeyAscii% = 0
    End If
End If

If KeyAscii% = 13 Then
    
    KeyAscii% = 0

    CommandStrt% = InStrRev(Console.Text, ">>")
    CommandUndec$ = Right(Console.Text, Len(Console.Text) - (CommandStrt% + 1))
    CommandEnd% = InStr(CommandUndec$, "::")

        If CommandEnd% = 0 Then
            CommandEnd% = Len(CommandUndec$) + 1
        End If

    Command$ = Left(CommandUndec$, CommandEnd% - 1)
    
        Select Case Command$
        
            Case "exit"
                End
                
            Case "hello"
                Console.Text = Console.Text & vbCrLf & _
                    "Hello " & User$ & "!"
                Com = True
                
            Case "user"
                ParamOne$ = Right(CommandUndec$, Len(CommandUndec$) - 6)
                User$ = ParamOne$
                Console.Text = Console.Text & vbCrLf & _
                "Changed user to :  " & User$
                Com = True
             
            Case "repos"
                ParamOne$ = Right(CommandUndec$, Len(CommandUndec$) - 7)
                ParLen% = InStr(ParamOne$, "::")
                CoorOne% = Left(ParamOne$, ParLen% - 1)
                CoorTwo% = Right(ParamOne$, ParLen% - 1)
                Form1.Left = CoorOne%
                Form1.Top = CoorTwo%
                Console.Text = Console.Text & vbCrLf & _
                    "Window Position changed to :  X-" & CoorOne% & "Y-" & CoorTwo%
                Com = True
                
                
        End Select
        
        If Com = False Then
            Console.Text = Console.Text & vbCrLf & _
                "Error, Incorrect Syntax or no command by that name!"
        End If
        
    Console.Text = Console.Text & vbCrLf & "OR~>>"
    Console.SelStart = Len(Console.Text)
End If

End Sub


Private Sub Form_Load()
Console.Text = "Odd Realism" & vbCrLf & vbCrLf
Console.Text = Console.Text & "This is version 1.0 of my prompt driven interface example. I hope " & vbCrLf
Console.Text = Console.Text & "this is very useful to you.  The drag method in the label is odd and " & vbCrLf
Console.Text = Console.Text & "original.  I didn't mess with it much so it does some odd things." & vbCrLf
Console.Text = Console.Text & "There are only three commands: hello, user::<param>, exit.  When you use " & vbCrLf
Console.Text = Console.Text & "user use the correct syntax : user::john  user::Odd Realism etc... The others you type in." & vbCrLf
Console.Text = Console.Text & vbCrLf & "~Odd Realism, geneshepherd@msn.com"
Console.Text = Console.Text & vbCrLf & vbCrLf & "OR~>>"
Console.SelStart = Len(Console.Text)

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim DifferenceX As Integer, DifferenceY As Integer

If Button = 1 Then

    If DifferenceX% = 0 Or DifferenceY% = 0 Then
        DifferenceX% = X
        DifferenceY% = Y
    Else
        DifferenceX% = X + DifferenceX%
        DifferenceY% = Y + DifferenceY%
    End If

  
   
        Form1.Left = Form1.Left + DifferenceX%
        Form1.Top = Form1.Top + DifferenceY%
    
End If
    



        
End Sub



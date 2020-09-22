VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmIDE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Intellisense Project"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   6465
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option3 
      Caption         =   "Combo of both"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4785
      TabIndex        =   12
      Top             =   4785
      Width           =   1470
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Visual Basic 6 Style"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3075
      TabIndex        =   9
      Top             =   4785
      Width           =   1770
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Office XP Style"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1515
      TabIndex        =   8
      Top             =   4785
      Value           =   -1  'True
      Width           =   1470
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Important Notes"
      Height          =   330
      Left            =   4785
      TabIndex        =   6
      Top             =   4320
      Width           =   1620
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&View Objects File"
      Height          =   330
      Left            =   4785
      TabIndex        =   4
      Top             =   3945
      Width           =   1620
   End
   Begin VB.PictureBox picIntellisense 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   1845
      ScaleHeight     =   1605
      ScaleWidth      =   2640
      TabIndex        =   1
      Top             =   2070
      Visible         =   0   'False
      Width           =   2640
      Begin VB.PictureBox picIListCont 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   60
         ScaleHeight     =   1500
         ScaleWidth      =   2520
         TabIndex        =   11
         Top             =   45
         Width           =   2520
         Begin VB.ListBox lstIntellisense 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   975
            IntegralHeight  =   0   'False
            Left            =   105
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   135
            Width           =   2265
         End
      End
      Begin VB.CommandButton cmdVBBorder 
         Height          =   1605
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   2640
      End
   End
   Begin VB.PictureBox picIntellisenseShadow 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   1920
      ScaleHeight     =   1605
      ScaleWidth      =   2640
      TabIndex        =   2
      Top             =   2145
      Visible         =   0   'False
      Width           =   2640
   End
   Begin RichTextLib.RichTextBox RTF1 
      Height          =   3840
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   6773
      _Version        =   393217
      BackColor       =   16777215
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      RightMargin     =   1e7
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Select ""Look"":"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   10
      Top             =   4785
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":00C4
      Height          =   900
      Left            =   75
      TabIndex        =   3
      Top             =   3885
      Width           =   4665
   End
End
Attribute VB_Name = "frmIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// /////////////////////////////////////////////////////////////// //
'//                                                                 //
'// INTELLISENSE PROJECT                                            //
'// REVISION 2 -- AUGUST 9, 2001                                    //
'//                                                                 //
'// BY: ANTHONY DUNLEAVY                                            //
'// ad@atrixware.com                                                //
'//                                                                 //
'// /////////////////////////////////////////////////////////////// //

'// ALL VARIABLES MUST BE DEFINED                                   //
    Option Explicit
        
Private Sub Command1_Click()

'// CALL NOTEPAD TO SHOW OBJECTS FILE                               //
    Dim Dummy As Long
    Dummy = ShellExecute(Me.hwnd, vbNullString, CurDir & "\obj.txt", _
                         vbNullString, "c:\", 1)
                       
End Sub

Private Sub Command2_Click()

'// CALL NOTEPAD TO SHOW OBJECTS FILE                               //
    Dim Dummy As Long
    Dummy = ShellExecute(Me.hwnd, vbNullString, CurDir & "\notes.txt", _
                         vbNullString, "c:\", 1)
                       
End Sub



Private Sub RTF1_Click()

'// HIDE INTELLISENSE BOX AND SHADOW                                //
    doShowIntellisense False
    
End Sub

Private Sub RTF1_KeyPress(KeyAscii As Integer)

'// RECEIVE KEYPRESS FROM INTELLISENSE BOX                          //
'// DO NOT REMOVE THIS COMMENT                                      //

End Sub

Private Sub RTF1_KeyUp(KeyCode As Integer, Shift As Integer)
       
'// WAS A "DOT" TYPED? IF NOT, HIDE INTELLISENSE BOX AND SHADOW     //
    If KeyCode <> 190 Then
        doShowIntellisense False
        Exit Sub
    End If
    
'// A "DOT" WAS TYPED, SO BRING UP INTELLISENSE BOX                 //
    Dim tmpEnd As Integer
    Dim tmpBeg As Integer
    Dim iEval As Integer
    Dim tmpStr As String
    Dim tmp2 As String
    Dim X As Integer
    Dim eChar As String
    Dim codeObj As String   'LOCAL "gObjName" Temp String
    Dim F As String
    Dim tmpLeft As Integer
    Dim tmpTop As Integer
    Dim tmpCount As Integer
    
'// "iEval" = # CHARACTERS TO GO BACKWARDS TO CHECK FOR OBJECT NAME //
'// CHANGE TO = THE # CHARACTERS OF YOUR LARGEST OBJECT. FOR EXAMPLE//
'// "FORM1." IS AN OBJECT, YOU MAY HAVE AN OBJECT IN YOUR SCRIPTING //
'// LANGUAGE NAMED "SOMEVERYLONGOBJECTNAME.", AND YOU WILL NEED     //
'// TO CHANGE THE VALUE OF THIS TO ACCOUNT FOR THE LONG NAMES       //
    iEval = 20
    
'// GRAB "OBJECT NAME" FROM CHARACTERS BEFORE "DOT". THERE MAY      //
'// BE CHARACTERS OR WORDS BEFORE THE "OBJECT" NAME, SO WE NEED     //
'// TO LOOK FOR SPACES, DOTS, OR LINE FEEDS                         //
    tmpEnd = RTF1.SelStart
    tmpBeg = tmpEnd - iEval
    If tmpBeg < 0 Then tmpBeg = 0
    RTF1.SelStart = tmpBeg
    RTF1.SelLength = tmpEnd - tmpBeg
    
'// MAKE SURE LAST CHARACTER IS A "DOT"                             //
'// THIS WILL HANDLE MOST OCCASIONS, AND OTHERWISE WILL GRACEFULLY  //
'// CONTINUE (AND NOT CAUSE A RUNTIME) IF IT DOES NOT RESOLVE       //
'// (NOTE: ORIG SUBMISSION WOULD FREEZE HERE. -- THIS FIXES PROB    //
    If Right$(RTF1.SelText, 1) <> "." Then
    Do Until Right$(RTF1.SelText, 1) = "."
       tmpCount = tmpCount + 1
       If tmpCount = 5 Then
        tmpCount = 0
        Exit Do
       End If
       RTF1.SelLength = RTF1.SelLength + 1
    Loop
    End If
    
'// WE HAVE OUR EVALUATION STRING. NOW WE NEED TO SEPARATE THE      //
'// OBJECT NAME AS BEST AS WE CAN DETERMINE. THIS MAY HAVE BEEN     //
'// A NICE CLEAN ENTRY "LIKE AT A BEGINNING OF A LINE", BUT WE      //
'// NEED TO ALSO LOOK FOR STUFF LIKE ENTERING A DOT IN THE MIDDLE   //
'// OF A LINE.                                                      //
    tmpStr$ = RTF1.SelText
    RTF1.SelStart = tmpEnd
    RTF1.SelLength = 0
    tmp2$ = ""
    
    For X = 1 To (Len(tmpStr$) - 1)
    eChar$ = Mid$(tmpStr$, X, 1)

'// WE WILL QUALIFY UNLESS VBCR, VBLF, SPACE, OR DOT                //
    If eChar <> Chr$(10) And eChar <> Chr$(13) _
       And eChar <> " " And eChar <> "." Then
        tmp2$ = tmp2$ & eChar$
    Else
        tmp2$ = " "
    End If
    
    Next
    
'// ONCE AGAIN, LETS MAKE SURE ENDING CHAR IS A "DOT"               //
    tmp2$ = Trim$(tmp2)
    If Right$(tmp2, 1) <> "." Then tmp2 = tmp2 & "."
    
'// PLACE OBJ NAME TEMPORARILY INTO LOCAL VAR codeObj               //
    codeObj = tmp2$
    
'// HERE IS THE NAME OF THE FILE USED FOR THE INTELLISENSE          //
'// YOU CAN CHANGE TO WHATEVER FITS YOUR TASTES OR ENVIRONMENT      //
'// AS LONG AS FORMAT IS FOLLOWED                                   //
    Dim iFile$
    iFile$ = CurDir & "\obj.txt"
    
'// THIS ROUTINE FILLS INTELLISENSE, AND THEN SHOWS IT IF IT        //
'// CONTAINS AT LEAST ONE ITEM IN THE LIST                          //
    doFillIntellisense iFile$, codeObj$
        
End Sub

Private Sub lstIntellisense_DblClick()

'// INITIATE THE PRESSING OF THE ENTER KEY                         //
    lstIntellisense_KeyPress (13)
    
End Sub

Private Sub lstIntellisense_KeyDown(KeyCode As Integer, Shift As Integer)

'// 'ESCAPE KEY HIDES INTELLISENSE BOX                              //
    If KeyCode = 27 Then
        doShowIntellisense False
        RTF1.SelStart = RTF1.SelStart + RTF1.SelLength
        RTF1.SelLength = 0
        RTF1.SetFocus
    End If
    
End Sub

Private Sub lstIntellisense_KeyPress(KeyAscii As Integer)
    
'// REMOVE THIS IF YOU WANT TO HANDLE EXTREME OCCASIONS ON AN           //
'// INCIDENT BY INCIDENT BASIS. FOR NOW, I USE THIS ERROR RESUME        //
'// TO GRACEFULLY CONTINUE WITHOUT RUNTIMES -- YOUR SCRIPTING           //
'// LANGUAGE MAY HAVE SITUATIONS YOU WANT TO HANDLE MANUALLY
    On Error Resume Next
    
'// SET UP LOCAL VARIABLES                                              //
    Dim origStart As Integer
    Dim origLen As Integer
    Dim tmp As String
    Dim tmpPad As String
    Dim tmpLoop As Integer
    Dim oldStart As Integer
    Dim oldLen As Integer
    Dim tmpDel As Integer
    
'// DETERMINE WHAT KEY WAS PRESSED                                      //
    Select Case KeyAscii
    
'// A-Z, a-z, ".", 0-9                                                  //
'// IF YOUR PROPERTY OR OBJECT NAMES CONTAIN OTHER CHARACTERS           //
'// MAKE SURE YOU INCLUDE THEM IN THIS CASE STATEMENT                   //
    Case 65 To 90, 97 To 122, 46, 48 To 57
    
    origStart = RTF1.SelStart
    origLen = RTF1.SelLength
    RTF1.SelText = RTF1.SelText & Chr$(KeyAscii)
    RTF1.SelStart = origStart
    RTF1.SelLength = origLen + 1
    
'// WE WANT THE LIST TO REFLECT CHARACTER WE HAVE TYPED, AND HIGHLIGHT  //
'// AN ITEM THAT CONTAINS THE CHARACTERS THAT WE HAVE TYPED, SO         //
'// CHANGE LIST INDEX IF MATCH                                          //
    tmp$ = Right$(RTF1.SelText, Len(RTF1.SelText) _
         - InStr(1, RTF1.SelText, "."))
    tmpPad = Space(Len(tmp$))
    For tmpLoop% = 0 To (lstIntellisense.ListCount - 1)
    If LCase$(Trim$(tmp$)) = _
       LCase$(Trim$(Left$(lstIntellisense.List(tmpLoop) _
       & tmpPad, Len(tmp$)))) Then
        lstIntellisense.ListIndex = tmpLoop
        Exit For
    End If
    Next
    
    End Select
    
'// WAS ANOTHER DOT TYPED WHILE WE ARE ALREADY IN AN OBJECT             //
'// EXAMPLE:  Form1.Font. << SECOND DOT NEEDS TO BRING UP               //
'// A DIFFERENT INTELLISENSE LIST                                       //
    If KeyAscii = 46 Then
        Dim newCmd
        newCmd = Right$(RTF1.SelText, Len(RTF1.SelText) - Len(gObjName))
        RTF1.SelStart = RTF1.SelStart + Len(gObjName) + Len(newCmd)
        RTF1.SelLength = 0
        doShowIntellisense False
        RTF1_KeyUp 190, 0
        RTF1.SetFocus
        RTF1.SelStart = RTF1.SelStart + Len(newCmd)
        RTF1.SelLength = 0
    End If
    
'// WAS THE "ENTER" KEY PRESSED? IF SO, SELECT HIGHLIGHTED ITEM IN LIST //
    If KeyAscii = 13 Then
        If Right$(RTF1.SelText, 1) = "." Then
            RTF1.SelStart = RTF1.SelStart + RTF1.SelLength
            RTF1.SelLength = 0
        Else
            oldLen = RTF1.SelLength
            tmpDel = InStr(1, RTF1.SelText, ".")
            RTF1.SelStart = RTF1.SelStart + tmpDel
            RTF1.SelLength = oldLen - tmpDel
        End If
        RTF1.SelText = lstIntellisense.List(lstIntellisense.ListIndex)
        doShowIntellisense False
        RTF1.SetFocus
    End If
    
'// WAS "BACKSPACE" KEY PRESSED? IF SO, WE NEED TO ADJUST THE TEXT      //
'// THAT APPEARS IN THE RTF BOX, AND THEN RE-SEARCH THE INTELLISENSE    //
'// LIST AND FIND THE CLOSEST MATCH AGAIN IF THERE WAS OTHER TEXT       //
'// ENTERED, OR IF WE ARE DELETING THE "DOT", THEN REMOVE THE           //
'// INTELLISENSE BOX                                                    //
    If KeyAscii = 8 Then

'// WAS ANY TEXT ENTERED (OR ARE WE DELETING A "DOT") WHICH MEANS       //
'// GET RID OF THE INTELLISENSE LIST                                    //
        If Right$(RTF1.SelText, 1) = "." Then
            RTF1.SelStart = RTF1.SelStart + (RTF1.SelLength - 1)
            RTF1.SelLength = 1
            RTF1.SelText = ""
            RTF1_KeyPress (KeyAscii)
            RTF1.SetFocus
'// IT WASNT A DOT, SO WE NEED TO DELETE ONE CHARACTER ON THE RTF BOX,  //
'// AND THEN SEARCH THE LIST FOR CLOSEST MATCH AGAIN                    //
        Else
            oldStart = RTF1.SelStart
            oldLen = RTF1.SelLength
            RTF1.SelStart = RTF1.SelStart + (RTF1.SelLength - 1)
            RTF1.SelLength = 1
            RTF1.SelText = ""
            RTF1_KeyPress (KeyAscii)
            
            If oldStart > 0 Then RTF1.SelStart = oldStart
            If oldLen > 0 Then RTF1.SelLength = oldLen - 1
            tmp$ = Right$(RTF1.SelText, Len(RTF1.SelText) _
                 - InStr(1, RTF1.SelText, "."))
            tmpPad = Space(Len(tmp$))
            For tmpLoop% = 0 To (lstIntellisense.ListCount - 1)
            If LCase$(Trim$(tmp$)) = _
               LCase$(Trim$(Left$(lstIntellisense.List(tmpLoop) _
               & tmpPad, Len(tmp$)))) Then
            lstIntellisense.ListIndex = tmpLoop
            Exit For
            End If
            Next
        End If
    End If
    
End Sub

Private Sub lstIntellisense_LostFocus()

'// IF LIST LOSES FOCUS, THEN ACT AS IF THE TAB KEY WAS PRESSED     //

    Dim oldLen As Integer
    Dim tmpDel As Integer

'// ACCOUNT FOR "TAB"                                               //
    If picIntellisense.Visible = True And _
       Me.ActiveControl <> Me.RTF1 Then
         If Right$(RTF1.SelText, 1) = "." Then
            RTF1.SelStart = RTF1.SelStart + RTF1.SelLength
            RTF1.SelLength = 0
        Else
            oldLen = RTF1.SelLength
            tmpDel = InStr(1, RTF1.SelText, ".")
            RTF1.SelStart = RTF1.SelStart + tmpDel
            RTF1.SelLength = oldLen - tmpDel
        End If
        If lstIntellisense.ListIndex > -1 Then
            RTF1.SelText = lstIntellisense.List(lstIntellisense.ListIndex)
        Else
            RTF1.SelText = lstIntellisense.List(0)
        End If
        doShowIntellisense False
        RTF1.SetFocus
    End If
   
End Sub

Private Sub doFillIntellisense(iFile As String, codeObj As String)
        
'// CHECK FILE FOR OBJECT NAME, AND FILL LIST IF MATCH FOUND        //
    
    Dim A As String
    Dim FF As Integer
    Dim lngStart As Integer
    Dim tmpLeft As Integer
    Dim tmpTop As Integer
    Dim PT As POINTAPI
    
'// CLEAR OUT ANY OLD ENTRIES IN INTELLISENSE LIST                  //
    lstIntellisense.Clear
    
'// INITIALIZE THE GLOBAL VARIABLE gObjName to Nothing              //
    gObjName = ""
    
'// OPEN FILE AND CHECK FOR OBJECT NAME AND PROPERTIES              //
    FF = FreeFile
    Open iFile For Input As FF
    
'// FIRST, FIND OBJECT NAME IF IT EXISTS                            //
    Do Until LCase$(Trim$(A$)) = LCase$(codeObj$)
    If EOF(FF) Then Exit Do
    Line Input #FF, A$
    Loop

'// NOW LOOK FOR PROPERTIES OF OBJECT                               //
    Do Until Trim$(A$) = "!"
    Line Input #FF, A$
    If Trim$(A$) <> "!" And Trim$(A$) <> "" Then
        lstIntellisense.AddItem Right$(A$, Len(A$) - Len(codeObj))
    End If
    Loop
    
    Close FF
    
'// IF NOTHING IN LIST, NO NEED TO DISPLAYY IT                      //
    If lstIntellisense.ListCount = 0 Then Exit Sub
    
'// MATCH FOUND, SO SET GLOBAL VAR gOBJNAME                         //
    gObjName = codeObj$

'// SETUP SELECTED TEXT IN RTF BOX SO WHEN WE BEGIN                 //
'// TYPING WHEN INTELLISENSE IS THE ACTIVE CONTROL, WE CAN SAFELY   //
'// PASS THE CHARACTERS TO THE RTF BOX AND HAVE THEM APPEAR WHERE   //
'// THEY BELONG                                                     //
    lngStart = RTF1.SelStart
    If RTF1.SelStart - Len(codeObj) > 0 Then
        RTF1.SelStart = RTF1.SelStart - Len(codeObj$)
    Else
        RTF1.SelStart = 0
    End If
    
    RTF1.SelLength = Len(codeObj$)
    
'// FIGURE OUT WHERE INTELLISENSE BOX SHOULD APPEAR                 //

'// WHAT ARE CURSOR COORDINATES                                     //
    GetCaretPos PT
    
'// PREFERENCE IS BOTTOM RIGHT, BUT IF IT WONT FIT THERE, WE WILL   //
'// PLACE TO TOP RIGHT OR TOP LEFT OR BOTTOM LEFT                   //
    If (PT.X * 15) + (RTF1.SelFontSize * 15) + _
       picIntellisense.Width + 45 < ScaleWidth Then
        tmpLeft = (PT.X * 15) + (RTF1.SelFontSize * 15) + 45
    Else
        tmpLeft = (PT.X * 15) + (RTF1.SelFontSize * 15) - _
        picIntellisense.Width - 45
    End If
    
    If (PT.Y * 15) + (RTF1.SelFontSize * 25) + _
       picIntellisense.Height + 45 < ScaleHeight Then
        tmpTop = (PT.Y * 15) + (RTF1.SelFontSize * 25)
    Else
        tmpTop = (PT.Y * 15) + (RTF1.SelFontSize * 25) - _
        picIntellisense.Height - RTF1.SelFontSize * 25
    End If
    
'// SHOW AND LOCATE INTELLISENSE BOX                                //
    
'// OFFICE XP "SHADOWED" STYLE                                      //
    If Option1.Value = True Then
        doShowIntellisense True, tmpLeft, tmpTop

'// VB6 (3D-Windows 95/98 look)
    ElseIf Option2.Value = True Then
        doShowIntellisense True, tmpLeft, tmpTop, "VB"
    
'// COMBINATION -- VB 3D LOOK WITH XP SHADOW
    Else
        doShowIntellisense True, tmpLeft, tmpTop, "MIX"
    End If
    
End Sub

Private Sub doShowIntellisense(fVisible As Boolean, _
                               Optional vLeft%, Optional vTop%, _
                               Optional iStyle As String)
                               
    
'// CALLS OR HIDES INTELLISENSE                                     //

'// IF fvisible = false, then lets just hide the intellsense        //
    If Not fVisible Then
        picIntellisense.Visible = False
        picIntellisenseShadow.Visible = False
        Exit Sub
    End If
    
'// fvisible must be true, so lets position intellisense and then   //
'// set visible to true                                             //
    With picIntellisense
        .Move vLeft, vTop
        .Visible = True
    
    '// ASSUME "OFFICE XP STYLE"
        picIListCont.Move 0, 0, .Width - 15, .Height - 15
        lstIntellisense.Move 0, 0, .Width + 15, .Height + 15
        
    '// SHADOW                              //
        picIntellisenseShadow.Move .Left + 45, .Top + 45, _
                                   .Width, .Height
    End With
    
'// STYLE OF INTELLISENSE                                           //
    If iStyle = "" Then
        picIntellisenseShadow.Visible = True
    ElseIf iStyle = "MIX" Then
        picIntellisenseShadow.Visible = True
        With picIntellisense
            picIListCont.Move 30, 30, .Width - 60, .Height - 60
            lstIntellisense.Move -15, -15, _
                                 picIListCont.Width + 30, _
                                 picIListCont.Height + 30
        End With
            
    Else
    '// MODIFY LOOK TO VISUAL BASIC 6 LOOK
        With picIntellisense
            picIListCont.Move 30, 30, .Width - 60, .Height - 60
            lstIntellisense.Move -15, -15, _
                                 picIListCont.Width + 30, _
                                 picIListCont.Height + 30
        End With
    End If
    
    lstIntellisense.SetFocus

End Sub
    


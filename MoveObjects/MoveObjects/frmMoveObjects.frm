VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMoveObjects 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Mike's Move Object's Example - [Default]"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5250
   Icon            =   "frmMoveObjects.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog DataView 
      Left            =   4680
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Load Moved Objects:"
   End
   Begin VB.CheckBox Checkobject10 
      Caption         =   "Checkobject10"
      Height          =   255
      Left            =   3720
      TabIndex        =   34
      Top             =   2280
      Visible         =   0   'False
      Width           =   1400
   End
   Begin VB.CheckBox Checkobject9 
      Caption         =   "Checkobject9"
      Height          =   255
      Left            =   3720
      TabIndex        =   33
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Checkobject8 
      Caption         =   "Checkobject8"
      Height          =   255
      Left            =   3720
      TabIndex        =   32
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Checkobject7 
      Caption         =   "Checkobject7"
      Height          =   255
      Left            =   3720
      TabIndex        =   31
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Checkobject6 
      Caption         =   "Checkobject6"
      Height          =   255
      Left            =   3720
      TabIndex        =   30
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Checkobject5 
      Caption         =   "Checkobject5"
      Height          =   255
      Left            =   2160
      TabIndex        =   29
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Checkobject4 
      Caption         =   "Checkobject4"
      Height          =   255
      Left            =   2160
      TabIndex        =   28
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Checkobject3 
      Caption         =   "Checkobject3"
      Height          =   255
      Left            =   2160
      TabIndex        =   27
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Checkobject2 
      Caption         =   "Checkobject2"
      Height          =   255
      Left            =   2160
      TabIndex        =   26
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Checkobject1 
      Caption         =   "Checkobject1"
      Height          =   255
      Left            =   2160
      TabIndex        =   25
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton OBObject10 
      Caption         =   "OBObject10"
      Height          =   255
      Left            =   3480
      TabIndex        =   24
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton OBObject9 
      Caption         =   "OBObject9"
      Height          =   255
      Left            =   3480
      TabIndex        =   23
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton OBObject8 
      Caption         =   "OBObject8"
      Height          =   255
      Left            =   3480
      TabIndex        =   22
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton OBObject7 
      Caption         =   "OBObject7"
      Height          =   255
      Left            =   3480
      TabIndex        =   21
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton OBObject6 
      Caption         =   "OBObject6"
      Height          =   255
      Left            =   3480
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton OBObject5 
      Caption         =   "OBObject5"
      Height          =   255
      Left            =   1920
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton OBObject4 
      Caption         =   "OBObject4"
      Height          =   255
      Left            =   1920
      TabIndex        =   18
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton OBObject3 
      Caption         =   "OBObject3"
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton OBObject2 
      Caption         =   "OBObject2"
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton OBObject1 
      Caption         =   "OBObject1"
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CBObject1 
      Caption         =   "CBObject1"
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton CBObject10 
      Caption         =   "CBObject10"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton CBObject9 
      Caption         =   "CBObject9"
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton CBObject8 
      Caption         =   "CBObject8"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton CBObject7 
      Caption         =   "CBObject7"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton CBObject6 
      Caption         =   "CBObject6"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton CBObject5 
      Caption         =   "CBObject5"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton CBObject4 
      Caption         =   "CBObject4"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton CBObject3 
      Caption         =   "CBObject3"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton CBObject2 
      Caption         =   "CBObject2"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "VB-Objects"
      Height          =   2775
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      Begin VB.Frame Frame2 
         Caption         =   "Options"
         Height          =   1095
         Left            =   0
         TabIndex        =   35
         Top             =   1680
         Width           =   1695
         Begin VB.CommandButton Command4 
            Caption         =   "&Help"
            Height          =   300
            Left            =   90
            TabIndex        =   39
            Top             =   660
            Width           =   750
         End
         Begin VB.CommandButton Command3 
            Caption         =   "E&xit"
            Height          =   300
            Left            =   840
            TabIndex        =   38
            Top             =   660
            Width           =   750
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Save"
            Height          =   300
            Left            =   840
            TabIndex        =   37
            Top             =   360
            Width           =   750
         End
         Begin VB.CommandButton Command1 
            Caption         =   "L&oad"
            Height          =   300
            Left            =   90
            TabIndex        =   36
            Top             =   360
            Width           =   750
         End
      End
      Begin VB.Label AddCheckB 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Check Box"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label MSGLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Right Click an Object"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   40
         TabIndex        =   12
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label AddOB 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Option Button"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label AddCB 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Command Button"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu menuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu menuRename 
         Caption         =   "Re&name"
      End
      Begin VB.Menu Line2 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "frmMoveObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'                      -_________________________-

'                       * -Mike's Move Objects- *'
'                      _-------------------------_
'
'
'            -April 22, 2000-
'
'   Thanks for downloading my Move objects
'example. First I'd like to say...this is a
'long piece of coding if you noticed. So
'I recommend you use Arrays if your planning
'on doing something like this. I tried to use
'arrays but I quit after I finished half of this
'project and re-written it all over. (Problems with PointAPI)
'Anyway, here is my example. It lets you Add objects
'to the form and lets the User move the object anywhere
'on the form, I also made it so that they couldn't move
'the object off of the form. You can also Rename or Delete
'the object by Right Clicking them. And the hardest part of this
'Was being able to detect and Save all of the open Objects
'I mean..the object's position, caption and if it's visible
'And then be able to load all that Info on each object!
'Well here is my loooong project. But it's a perfect example
'For you programmers who want to make a program that lets
'The user move objects around on the Form.
'
'Here's an example on how this is useful..
'
'Visual HTML making programs...Like FrontPage or
'Macromedia DreamWeaver...They let the user add objects
'and lets them move it around. (Input Button, Option Button..Checkbutton...etc.)
'Well you could also use a PictureBox and insert a Picture into it so
'you can create different looking objects to move...

'Well here is my project, thanks for reading this And
'I hope You find this usefull.. If you do please
'E-mail me at: Mike@dev-center.com
'AIM: Mike3dd
'Because to know that somebody likes
'what I create is cool... You know?

'And also RATE this Please..I spent
'long Hours making this so take the time
'To rate it. Thanks Again....

'             VB Programmer,
'
'                -Mike Canejo-
'               (Dev-Center.com)






Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Sets a window to a position on the user's screen
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Gets the user's cursor coordinates
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Gets a value from an INI file
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'Writes a value to an INI file
Private Const HWND_TOPMOST = -1
'Makes a Window OnTop of other windows
Private Const HWND_NOTOPMOST = -2
'Makes a Window NotOnTop of other windows
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
'These tells the window to not be moved or resized
Private LastPoint As POINTAPI
'I put this here to subtract the current objects
'X and Y coordinates from LastPoint to make
'The object moveable.
Private TheTracker As Boolean
'If True then let the object be moveable
'If False then not let the object be moveable
Private HoldNumber As Integer
'I put this here to detect the current
'Object thats being manipulated.
Private thepointX As Long
'Holds the X cursor coordinate
Private thepointY As Long
'Holds the Y cursor coordinate
Private HoldButton As String
'Holds the current objects name thats being Manipulated
Private Type POINTAPI
x As Long: y As Long
End Type
'This is probobly thr most inportant in this project
'This gets the current X and Y coordinates if
'The user's mouse using the GetCursirPos Delare
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE 'Flags for my StayOnTop Function
Private Function StayOnTop(TheForm As Form)
    Dim SetWinOnTop As Long
    SetWinOnTop = SetWindowPos(TheForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
'This makes a form OnTop of other windows
End Function
Private Function NotOnTop(TheForm As Form)
Dim SetWinOnTop As Long
SetWinOnTop = SetWindowPos(TheForm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
'This makes a form Not OnTop of other windows
End Function
Private Function CenterForm(TheForm As Form)
    TheForm.Move (Screen.Width) / 2 - (TheForm.Width) / 2, (Screen.Height) / 2 - (TheForm.Height) / 2
'This centers the form based on the actual screen size and divides it by 2
End Function
Private Function WriteToINI(Section As String, key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, key$, KeyValue$, Directory$)
'This writes a value to a key in a INI file
End Function
Public Function GetFromINI(Section As String, key As String, Directory As String) As String
    Dim strBuffer As String: strBuffer = String(750, Chr(0))
    GetFromINI$ = Left(strBuffer, GetPrivateProfileString(LCase$(Section$), ByVal LCase$(key$), "", strBuffer, Len(strBuffer), Directory$))
'This gets a value from a key in a INI file.
End Function
Private Sub AddCB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim POINT As POINTAPI, z As Integer
If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
If Button = 1 Then
    If CBObject1.Visible = False Then CBObject1.Visible = True: CBObject1.SetFocus: CBObject1.Left = Frame1.Left + Frame1.Width: CBObject1.Top = Frame1.Top + AddCB.Top: HoldNumber = 1: Exit Sub
    If CBObject2.Visible = False Then CBObject2.Visible = True: CBObject2.SetFocus: CBObject2.Left = Frame1.Left + Frame1.Width: CBObject2.Top = Frame1.Top + AddCB.Top: HoldNumber = 2: Exit Sub
    If CBObject3.Visible = False Then CBObject3.Visible = True: CBObject3.SetFocus: CBObject3.Left = Frame1.Left + Frame1.Width: CBObject3.Top = Frame1.Top + AddCB.Top: HoldNumber = 3: Exit Sub
    If CBObject4.Visible = False Then CBObject4.Visible = True: CBObject4.SetFocus: CBObject4.Left = Frame1.Left + Frame1.Width: CBObject4.Top = Frame1.Top + AddCB.Top: HoldNumber = 4: Exit Sub
    If CBObject5.Visible = False Then CBObject5.Visible = True: CBObject5.SetFocus: CBObject5.Left = Frame1.Left + Frame1.Width: CBObject5.Top = Frame1.Top + AddCB.Top: HoldNumber = 5: Exit Sub
    If CBObject6.Visible = False Then CBObject6.Visible = True: CBObject6.SetFocus: CBObject6.Left = Frame1.Left + Frame1.Width: CBObject6.Top = Frame1.Top + AddCB.Top: HoldNumber = 6: Exit Sub
    If CBObject7.Visible = False Then CBObject7.Visible = True: CBObject7.SetFocus: CBObject7.Left = Frame1.Left + Frame1.Width: CBObject7.Top = Frame1.Top + AddCB.Top: HoldNumber = 7: Exit Sub
    If CBObject8.Visible = False Then CBObject8.Visible = True: CBObject8.SetFocus: CBObject8.Left = Frame1.Left + Frame1.Width: CBObject8.Top = Frame1.Top + AddCB.Top: HoldNumber = 8: Exit Sub
    If CBObject9.Visible = False Then CBObject9.Visible = True: CBObject9.SetFocus: CBObject9.Left = Frame1.Left + Frame1.Width: CBObject9.Top = Frame1.Top + AddCB.Top: HoldNumber = 9: Exit Sub
    If CBObject10.Visible = False Then CBObject10.Visible = True: CBObject10.SetFocus: CBObject10.Left = Frame1.Left + Frame1.Width: CBObject10.Top = Frame1.Top + AddCB.Top: HoldNumber = 10: Exit Sub
End If
'If the left mouse button is down on the add command button label,
'Make the Button movable by your cursor using PointAPI
End Sub

Private Sub AddCB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
AddCB.BorderStyle = 1: AddOB.BorderStyle = 0: AddCheckB.BorderStyle = 0
If Not TheTracker Then Exit Sub
If Button = 1 Then
    If HoldNumber = 1 Then CBMove CBObject1: Exit Sub
    If HoldNumber = 2 Then CBMove CBObject2: Exit Sub
    If HoldNumber = 3 Then CBMove CBObject3: Exit Sub
    If HoldNumber = 4 Then CBMove CBObject4: Exit Sub
    If HoldNumber = 5 Then CBMove CBObject5: Exit Sub
    If HoldNumber = 6 Then CBMove CBObject6: Exit Sub
    If HoldNumber = 7 Then CBMove CBObject7: Exit Sub
    If HoldNumber = 8 Then CBMove CBObject8: Exit Sub
    If HoldNumber = 9 Then CBMove CBObject9: Exit Sub
    If HoldNumber = 10 Then CBMove CBObject10: Exit Sub
End If
'If the left mouse button is down on the add command button label,
'Use PointAPI and let the user move the object
End Sub
Private Function CBMove(TheButton As Object)
Dim POINT As POINTAPI
    If TheButton.Left > Me.Width - TheButton.Width - 80 Then TheButton.Left = Me.Width - TheButton.Width - 100: Exit Function
    If TheButton.Top < 0 Then TheButton.Top = 0: Exit Function
    If TheButton.Top > Me.Height - TheButton.Height - 340 Then TheButton.Top = Me.Height - TheButton.Height - 340: Exit Function
    If TheButton.Left < Frame1.Left + Frame1.Width Then TheButton.Left = Frame1.Left + Frame1.Width: Exit Function
GetCursorPos POINT: thepointX& = (POINT.x - LastPoint.x) * Screen.TwipsPerPixelX: thepointY& = (POINT.y - LastPoint.y) * Screen.TwipsPerPixelY: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheButton.Move TheButton.Left + thepointX&, TheButton.Top + thepointY&: TheButton.Visible = True
'Lets User move an object and detect if its moved off
'of the form..then move it back onto the form
End Function

Private Sub AddCheckB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim POINT As POINTAPI, z As Integer
If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
If Button = 1 Then
    If Checkobject1.Visible = False Then Checkobject1.Visible = True: Checkobject1.Left = Frame1.Left + Frame1.Width: Checkobject1.Top = Frame1.Top + AddCheckB.Top: HoldNumber = 1: Exit Sub
    If Checkobject2.Visible = False Then Checkobject2.Visible = True: Checkobject2.Left = Frame1.Left + Frame1.Width: Checkobject2.Top = Frame1.Top + AddCheckB.Top: HoldNumber = 2: Exit Sub
    If Checkobject3.Visible = False Then Checkobject3.Visible = True: Checkobject3.Left = Frame1.Left + Frame1.Width: Checkobject3.Top = Frame1.Top + AddCheckB.Top: HoldNumber = 3: Exit Sub
    If Checkobject4.Visible = False Then Checkobject4.Visible = True: Checkobject4.Left = Frame1.Left + Frame1.Width: Checkobject4.Top = Frame1.Top + AddCheckB.Top: HoldNumber = 4: Exit Sub
    If Checkobject5.Visible = False Then Checkobject5.Visible = True: Checkobject5.Left = Frame1.Left + Frame1.Width: Checkobject5.Top = Frame1.Top + AddCheckB.Top: HoldNumber = 5: Exit Sub
    If Checkobject6.Visible = False Then Checkobject6.Visible = True: Checkobject6.Left = Frame1.Left + Frame1.Width: Checkobject6.Top = Frame1.Top + AddCheckB.Top: HoldNumber = 6: Exit Sub
    If Checkobject7.Visible = False Then Checkobject7.Visible = True: Checkobject7.Left = Frame1.Left + Frame1.Width: Checkobject7.Top = Frame1.Top + AddCheckB.Top: HoldNumber = 7: Exit Sub
    If Checkobject8.Visible = False Then Checkobject8.Visible = True: Checkobject8.Left = Frame1.Left + Frame1.Width: Checkobject8.Top = Frame1.Top + AddCheckB.Top: HoldNumber = 8: Exit Sub
    If Checkobject9.Visible = False Then Checkobject9.Visible = True: Checkobject9.Left = Frame1.Left + Frame1.Width: Checkobject9.Top = Frame1.Top + AddCheckB.Top: HoldNumber = 9: Exit Sub
    If Checkobject10.Visible = False Then Checkobject10.Visible = True: Checkobject10.Left = Frame1.Left + Frame1.Width: Checkobject10.Top = Frame1.Top + AddCheckB.Top: HoldNumber = 10: Exit Sub
End If
'If the left mouse button is down on the add Checkbox label,
'Make the Button movable by your cursor using PointAPI
End Sub

Private Sub AddCheckB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
AddCB.BorderStyle = 0: AddOB.BorderStyle = 0: AddCheckB.BorderStyle = 1
Dim POINT As POINTAPI, z As Integer
If Not TheTracker Then Exit Sub
If Button = 1 Then
    If HoldNumber = 1 Then CBMove Checkobject1: Exit Sub
    If HoldNumber = 2 Then CBMove Checkobject2: Exit Sub
    If HoldNumber = 3 Then CBMove Checkobject3: Exit Sub
    If HoldNumber = 4 Then CBMove Checkobject4: Exit Sub
    If HoldNumber = 5 Then CBMove Checkobject5: Exit Sub
    If HoldNumber = 6 Then CBMove Checkobject6: Exit Sub
    If HoldNumber = 7 Then CBMove Checkobject7: Exit Sub
    If HoldNumber = 8 Then CBMove Checkobject8: Exit Sub
    If HoldNumber = 9 Then CBMove Checkobject9: Exit Sub
    If HoldNumber = 10 Then CBMove Checkobject10: Exit Sub
End If
'If the left mouse button is down on the add checkbox label,
'Use PointAPI and let the user move the object
End Sub

Private Sub AddLabel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
AddCB.BorderStyle = 0: AddOB.BorderStyle = 0: AddCheckB.BorderStyle = 0
End Sub

Private Sub AddOB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim POINT As POINTAPI, z As Integer
If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
If Button = 1 Then
    If OBObject1.Visible = False Then OBObject1.Visible = True: OBObject1.SetFocus: OBObject1.Left = Frame1.Left + Frame1.Width: OBObject1.Top = Frame1.Top + AddOB.Top: HoldNumber = 1: Exit Sub
    If OBObject2.Visible = False Then OBObject2.Visible = True: OBObject2.SetFocus: OBObject2.Left = Frame1.Left + Frame1.Width: OBObject2.Top = Frame1.Top + AddOB.Top: HoldNumber = 2: Exit Sub
    If OBObject3.Visible = False Then OBObject3.Visible = True: OBObject3.SetFocus: OBObject3.Left = Frame1.Left + Frame1.Width: OBObject3.Top = Frame1.Top + AddOB.Top: HoldNumber = 3: Exit Sub
    If OBObject4.Visible = False Then OBObject4.Visible = True: OBObject4.SetFocus: OBObject4.Left = Frame1.Left + Frame1.Width: OBObject4.Top = Frame1.Top + AddOB.Top: HoldNumber = 4: Exit Sub
    If OBObject5.Visible = False Then OBObject5.Visible = True: OBObject5.SetFocus: OBObject5.Left = Frame1.Left + Frame1.Width: OBObject5.Top = Frame1.Top + AddOB.Top: HoldNumber = 5: Exit Sub
    If OBObject6.Visible = False Then OBObject6.Visible = True: OBObject6.SetFocus: OBObject6.Left = Frame1.Left + Frame1.Width: OBObject6.Top = Frame1.Top + AddOB.Top: HoldNumber = 6: Exit Sub
    If OBObject7.Visible = False Then OBObject7.Visible = True: OBObject7.SetFocus: OBObject7.Left = Frame1.Left + Frame1.Width: OBObject7.Top = Frame1.Top + AddOB.Top: HoldNumber = 7: Exit Sub
    If OBObject8.Visible = False Then OBObject8.Visible = True: OBObject8.SetFocus: OBObject8.Left = Frame1.Left + Frame1.Width: OBObject8.Top = Frame1.Top + AddOB.Top: HoldNumber = 8: Exit Sub
    If OBObject9.Visible = False Then OBObject9.Visible = True: OBObject9.SetFocus: OBObject9.Left = Frame1.Left + Frame1.Width: OBObject9.Top = Frame1.Top + AddOB.Top: HoldNumber = 9: Exit Sub
    If OBObject10.Visible = False Then OBObject10.Visible = True: OBObject10.SetFocus: OBObject10.Left = Frame1.Left + Frame1.Width: OBObject10.Top = Frame1.Top + AddOB.Top: HoldNumber = 10: Exit Sub
End If
'If the left mouse button is down on the add option button label,
'Make the Option Button movable by your cursor using PointAPI
End Sub

Private Sub AddOB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
AddCB.BorderStyle = 0: AddOB.BorderStyle = 1: AddCheckB.BorderStyle = 0
If Not TheTracker Then Exit Sub
If Button = 1 Then
    If HoldNumber = 1 Then CBMove OBObject1: Exit Sub
    If HoldNumber = 2 Then CBMove OBObject2: Exit Sub
    If HoldNumber = 3 Then CBMove OBObject3: Exit Sub
    If HoldNumber = 4 Then CBMove OBObject4: Exit Sub
    If HoldNumber = 5 Then CBMove OBObject5: Exit Sub
    If HoldNumber = 6 Then CBMove OBObject6: Exit Sub
    If HoldNumber = 7 Then CBMove OBObject7: Exit Sub
    If HoldNumber = 8 Then CBMove OBObject8: Exit Sub
    If HoldNumber = 9 Then CBMove OBObject9: Exit Sub
    If HoldNumber = 10 Then CBMove OBObject10: Exit Sub
End If
'If the left mouse button is down on the add optionbutton label,
'Use PointAPI and let the user move the object
End Sub

'The Below lets the User move the Object freely
'With their mouse when the left mousebutton is down
'It also lets the user to be able to right click the
'object to bring up the Menu1 menu so they can
'Rename or Delete the right-clicked object.

Private Sub CBObject1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub

Private Sub CBObject1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove CBObject1
End Sub

Private Sub CBObject1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "CBObject1": CBObject1.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub CBObject2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "CBObject2": CBObject2.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub CBObject3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "CBObject3": CBObject3.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub CBObject4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "CBObject4": CBObject4.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub CBObject5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "CBObject5": CBObject5.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub CBObject6_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "CBObject6": CBObject6.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub CBObject7_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "CBObject7": CBObject7.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub CBObject8_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "CBObject8": CBObject8.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub CBObject9_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "CBObject9": CBObject9.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub CBObject10_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "CBObject10": CBObject10.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub

Private Sub CBObject10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub

Private Sub CBObject10_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove CBObject10
End Sub

Private Sub CBObject2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub

Private Sub CBObject2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove CBObject2
End Sub

Private Sub CBObject3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub

Private Sub CBObject3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove CBObject3
End Sub

Private Sub CBObject4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub

Private Sub CBObject4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove CBObject4
End Sub

Private Sub CBObject5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub

Private Sub CBObject5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove CBObject5
End Sub

Private Sub CBObject6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub

Private Sub CBObject6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove CBObject6
End Sub

Private Sub CBObject7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub

Private Sub CBObject7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove CBObject7
End Sub

Private Sub CBObject8_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub

Private Sub CBObject8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove CBObject8
End Sub

Private Sub CBObject9_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub

Private Sub CBObject9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove CBObject9
End Sub

'Command1_Click lets you Load a saved "Project"
'That contains the Objects: Caption, Left, Top and Visibility
'it looks messed up but its just basically 1-10 or each thing
'This wouldn't be possible if it wern't for VB's Replace Feature...LoL

Private Sub Command1_Click()
On Error GoTo ExitonError
    DataView.FileName = ""
    DataView.Filter = "Moved Objects(*.dat)|*.dat"
    DataView.ShowOpen: DataView.DialogTitle = "Load Moved Objects:"
If DataView.FileName = "" Then Exit Sub
    CBObject1.Visible = False: CBObject1.Caption = "CBObject1": CBObject2.Visible = False: CBObject2.Caption = "CBObject2"
    CBObject3.Visible = False: CBObject3.Caption = "CBObject3": CBObject4.Visible = False: CBObject4.Caption = "CBObject4"
    CBObject5.Visible = False: CBObject5.Caption = "CBObject5": CBObject6.Visible = False: CBObject6.Caption = "CBObject6"
    CBObject7.Visible = False: CBObject7.Caption = "CBObject7": CBObject8.Visible = False: CBObject8.Caption = "CBObject8"
    CBObject9.Visible = False: CBObject9.Caption = "CBObject9": CBObject10.Visible = False: CBObject10.Caption = "CBObject10"
    OBObject1.Visible = False: OBObject1.Caption = "OBObject1": OBObject2.Visible = False: OBObject2.Caption = "OBObject2"
    OBObject3.Visible = False: OBObject3.Caption = "OBObject3": OBObject4.Visible = False: OBObject4.Caption = "OBObject4"
    OBObject5.Visible = False: OBObject5.Caption = "OBObject5": OBObject6.Visible = False: OBObject6.Caption = "OBObject6"
    OBObject7.Visible = False: OBObject7.Caption = "OBObject7": OBObject8.Visible = False: OBObject8.Caption = "OBObject8"
    OBObject9.Visible = False: OBObject9.Caption = "OBObject9": OBObject10.Visible = False: OBObject10.Caption = "OBObject10"
    Checkobject1.Visible = False: Checkobject1.Caption = "Checkobject1": Checkobject2.Visible = False: Checkobject2.Caption = "Checkobject2"
    Checkobject3.Visible = False: Checkobject3.Caption = "Checkobject3": Checkobject4.Visible = False: Checkobject4.Caption = "Checkobject4"
    Checkobject5.Visible = False: Checkobject5.Caption = "Checkobject5": Checkobject6.Visible = False: Checkobject6.Caption = "Checkobject6"
    Checkobject7.Visible = False: Checkobject7.Caption = "Checkobject7": Checkobject8.Visible = False: Checkobject8.Caption = "Checkobject8"
    Checkobject9.Visible = False: Checkobject9.Caption = "Checkobject9": Checkobject10.Visible = False: Checkobject10.Caption = "Checkobject10"
    If Len(GetFromINI("CBObject1", "Left", DataView.FileName)) > 0 Then CBObject1.Left = GetFromINI("CBObject1", "Left", DataView.FileName):
    If Len(GetFromINI("CBObject1", "Top", DataView.FileName)) > 0 Then CBObject1.Top = GetFromINI("CBObject1", "Top", DataView.FileName)
    If GetFromINI("CBObject1", "Visible", DataView.FileName) = "True" Then CBObject1.Visible = True: If Len(GetFromINI("CBObject1", "Caption", DataView.FileName)) > 0 Then CBObject1.Caption = GetFromINI("CBObject1", "Caption", DataView.FileName)
    If Len(GetFromINI("CBObject2", "Left", DataView.FileName)) > 0 Then CBObject2.Left = GetFromINI("CBObject2", "Left", DataView.FileName)
    If Len(GetFromINI("CBObject2", "Top", DataView.FileName)) > 0 Then CBObject2.Top = GetFromINI("CBObject2", "Top", DataView.FileName)
    If GetFromINI("CBObject2", "Visible", DataView.FileName) = "True" Then CBObject2.Visible = True: If Len(GetFromINI("CBObject2", "Caption", DataView.FileName)) > 0 Then CBObject2.Caption = GetFromINI("CBObject2", "Caption", DataView.FileName)
    If Len(GetFromINI("CBObject3", "Left", DataView.FileName)) > 0 Then CBObject3.Left = GetFromINI("CBObject3", "Left", DataView.FileName)
    If Len(GetFromINI("CBObject3", "Top", DataView.FileName)) > 0 Then CBObject3.Top = GetFromINI("CBObject3", "Top", DataView.FileName)
    If GetFromINI("CBObject3", "Visible", DataView.FileName) = "True" Then CBObject3.Visible = True: If Len(GetFromINI("CBObject3", "Caption", DataView.FileName)) > 0 Then CBObject3.Caption = GetFromINI("CBObject3", "Caption", DataView.FileName)
    If Len(GetFromINI("CBObject4", "Left", DataView.FileName)) > 0 Then CBObject4.Left = GetFromINI("CBObject4", "Left", DataView.FileName)
    If Len(GetFromINI("CBObject4", "Top", DataView.FileName)) > 0 Then CBObject4.Top = GetFromINI("CBObject4", "Top", DataView.FileName)
    If GetFromINI("CBObject4", "Visible", DataView.FileName) = "True" Then CBObject4.Visible = True: If Len(GetFromINI("CBObject4", "Caption", DataView.FileName)) > 0 Then CBObject4.Caption = GetFromINI("CBObject4", "Caption", DataView.FileName)
    If Len(GetFromINI("CBObject5", "Left", DataView.FileName)) > 0 Then CBObject5.Left = GetFromINI("CBObject5", "Left", DataView.FileName)
    If Len(GetFromINI("CBObject5", "Top", DataView.FileName)) > 0 Then CBObject5.Top = GetFromINI("CBObject5", "Top", DataView.FileName)
    If GetFromINI("CBObject5", "Visible", DataView.FileName) = "True" Then CBObject5.Visible = True: If Len(GetFromINI("CBObject5", "Caption", DataView.FileName)) > 0 Then CBObject5.Caption = GetFromINI("CBObject5", "Caption", DataView.FileName)
    If Len(GetFromINI("CBObject6", "Left", DataView.FileName)) > 0 Then CBObject6.Left = GetFromINI("CBObject6", "Left", DataView.FileName)
    If Len(GetFromINI("CBObject6", "Top", DataView.FileName)) > 0 Then CBObject6.Top = GetFromINI("CBObject6", "Top", DataView.FileName)
    If GetFromINI("CBObject6", "Visible", DataView.FileName) = "True" Then CBObject6.Visible = True: If Len(GetFromINI("CBObject6", "Caption", DataView.FileName)) > 0 Then CBObject6.Caption = GetFromINI("CBObject6", "Caption", DataView.FileName)
    If Len(GetFromINI("CBObject7", "Left", DataView.FileName)) > 0 Then CBObject7.Left = GetFromINI("CBObject7", "Left", DataView.FileName)
    If Len(GetFromINI("CBObject7", "Top", DataView.FileName)) > 0 Then CBObject7.Top = GetFromINI("CBObject7", "Top", DataView.FileName)
    If GetFromINI("CBObject7", "Visible", DataView.FileName) = "True" Then CBObject7.Visible = True: If Len(GetFromINI("CBObject7", "Caption", DataView.FileName)) > 0 Then CBObject7.Caption = GetFromINI("CBObject7", "Caption", DataView.FileName)
    If Len(GetFromINI("CBObject8", "Left", DataView.FileName)) > 0 Then CBObject8.Left = GetFromINI("CBObject8", "Left", DataView.FileName)
    If Len(GetFromINI("CBObject8", "Top", DataView.FileName)) > 0 Then CBObject8.Top = GetFromINI("CBObject8", "Top", DataView.FileName)
    If GetFromINI("CBObject8", "Visible", DataView.FileName) = "True" Then CBObject8.Visible = True: If Len(GetFromINI("CBObject8", "Caption", DataView.FileName)) > 0 Then CBObject8.Caption = GetFromINI("CBObject8", "Caption", DataView.FileName)
    If Len(GetFromINI("CBObject9", "Left", DataView.FileName)) > 0 Then CBObject9.Left = GetFromINI("CBObject9", "Left", DataView.FileName)
    If Len(GetFromINI("CBObject9", "Top", DataView.FileName)) > 0 Then CBObject9.Top = GetFromINI("CBObject9", "Top", DataView.FileName)
    If GetFromINI("CBObject9", "Visible", DataView.FileName) = "True" Then CBObject9.Visible = True: If Len(GetFromINI("CBObject9", "Caption", DataView.FileName)) > 0 Then CBObject9.Caption = GetFromINI("CBObject9", "Caption", DataView.FileName)
    If Len(GetFromINI("CBObject10", "Left", DataView.FileName)) > 0 Then CBObject10.Left = GetFromINI("CBObject10", "Left", DataView.FileName)
    If Len(GetFromINI("CBObject10", "Top", DataView.FileName)) > 0 Then CBObject10.Top = GetFromINI("CBObject10", "Top", DataView.FileName)
    If GetFromINI("CBObject10", "Visible", DataView.FileName) = "True" Then CBObject10.Visible = True: If Len(GetFromINI("CBObject10", "Caption", DataView.FileName)) > 0 Then CBObject10.Caption = GetFromINI("CBObject10", "Caption", DataView.FileName)
    If Len(GetFromINI("OBObject1", "Left", DataView.FileName)) > 0 Then OBObject1.Left = GetFromINI("OBObject1", "Left", DataView.FileName)
    If Len(GetFromINI("OBObject1", "Top", DataView.FileName)) > 0 Then OBObject1.Top = GetFromINI("OBObject1", "Top", DataView.FileName)
    If GetFromINI("OBObject1", "Visible", DataView.FileName) = "True" Then OBObject1.Visible = True: If Len(GetFromINI("OBObject1", "Caption", DataView.FileName)) > 0 Then OBObject1.Caption = GetFromINI("OBObject1", "Caption", DataView.FileName)
    If Len(GetFromINI("OBObject2", "Left", DataView.FileName)) > 0 Then OBObject2.Left = GetFromINI("OBObject2", "Left", DataView.FileName)
    If Len(GetFromINI("OBObject2", "Top", DataView.FileName)) > 0 Then OBObject2.Top = GetFromINI("OBObject2", "Top", DataView.FileName)
    If GetFromINI("OBObject2", "Visible", DataView.FileName) = "True" Then OBObject2.Visible = True: If Len(GetFromINI("OBObject2", "Caption", DataView.FileName)) > 0 Then OBObject2.Caption = GetFromINI("OBObject2", "Caption", DataView.FileName)
    If Len(GetFromINI("OBObject3", "Left", DataView.FileName)) > 0 Then OBObject3.Left = GetFromINI("OBObject3", "Left", DataView.FileName)
    If Len(GetFromINI("OBObject3", "Top", DataView.FileName)) > 0 Then OBObject3.Top = GetFromINI("OBObject3", "Top", DataView.FileName)
    If GetFromINI("OBObject3", "Visible", DataView.FileName) = "True" Then OBObject3.Visible = True: If Len(GetFromINI("OBObject3", "Caption", DataView.FileName)) > 0 Then OBObject3.Caption = GetFromINI("OBObject3", "Caption", DataView.FileName)
    If Len(GetFromINI("OBObject4", "Left", DataView.FileName)) > 0 Then OBObject4.Left = GetFromINI("OBObject4", "Left", DataView.FileName)
    If Len(GetFromINI("OBObject4", "Top", DataView.FileName)) > 0 Then OBObject4.Top = GetFromINI("OBObject4", "Top", DataView.FileName)
    If GetFromINI("OBObject4", "Visible", DataView.FileName) = "True" Then OBObject4.Visible = True: If Len(GetFromINI("OBObject4", "Caption", DataView.FileName)) > 0 Then OBObject4.Caption = GetFromINI("OBObject4", "Caption", DataView.FileName)
    If Len(GetFromINI("OBObject5", "Left", DataView.FileName)) > 0 Then OBObject5.Left = GetFromINI("OBObject5", "Left", DataView.FileName)
    If Len(GetFromINI("OBObject5", "Top", DataView.FileName)) > 0 Then OBObject5.Top = GetFromINI("OBObject5", "Top", DataView.FileName)
    If GetFromINI("OBObject5", "Visible", DataView.FileName) = "True" Then OBObject5.Visible = True: If Len(GetFromINI("OBObject5", "Caption", DataView.FileName)) > 0 Then OBObject5.Caption = GetFromINI("OBObject5", "Caption", DataView.FileName)
    If Len(GetFromINI("OBObject6", "Left", DataView.FileName)) > 0 Then OBObject6.Left = GetFromINI("OBObject6", "Left", DataView.FileName)
    If Len(GetFromINI("OBObject6", "Top", DataView.FileName)) > 0 Then OBObject6.Top = GetFromINI("OBObject6", "Top", DataView.FileName)
    If GetFromINI("OBObject6", "Visible", DataView.FileName) = "True" Then OBObject6.Visible = True: If Len(GetFromINI("OBObject6", "Caption", DataView.FileName)) > 0 Then OBObject6.Caption = GetFromINI("OBObject6", "Caption", DataView.FileName)
    If Len(GetFromINI("OBObject7", "Left", DataView.FileName)) > 0 Then OBObject7.Left = GetFromINI("OBObject7", "Left", DataView.FileName)
    If Len(GetFromINI("OBObject7", "Top", DataView.FileName)) > 0 Then OBObject7.Top = GetFromINI("OBObject7", "Top", DataView.FileName)
    If GetFromINI("OBObject7", "Visible", DataView.FileName) = "True" Then OBObject7.Visible = True: If Len(GetFromINI("OBObject7", "Caption", DataView.FileName)) > 0 Then OBObject7.Caption = GetFromINI("OBObject7", "Caption", DataView.FileName)
    If Len(GetFromINI("OBObject8", "Left", DataView.FileName)) > 0 Then OBObject8.Left = GetFromINI("OBObject8", "Left", DataView.FileName)
    If Len(GetFromINI("OBObject8", "Top", DataView.FileName)) > 0 Then OBObject8.Top = GetFromINI("OBObject8", "Top", DataView.FileName)
    If GetFromINI("OBObject8", "Visible", DataView.FileName) = "True" Then OBObject8.Visible = True: If Len(GetFromINI("OBObject8", "Caption", DataView.FileName)) > 0 Then OBObject8.Caption = GetFromINI("OBObject8", "Caption", DataView.FileName)
    If Len(GetFromINI("OBObject9", "Left", DataView.FileName)) > 0 Then OBObject9.Left = GetFromINI("OBObject9", "Left", DataView.FileName)
    If Len(GetFromINI("OBObject9", "Top", DataView.FileName)) > 0 Then OBObject9.Top = GetFromINI("OBObject9", "Top", DataView.FileName)
    If GetFromINI("OBObject9", "Visible", DataView.FileName) = "True" Then OBObject9.Visible = True: If Len(GetFromINI("OBObject9", "Caption", DataView.FileName)) > 0 Then OBObject9.Caption = GetFromINI("OBObject9", "Caption", DataView.FileName)
    If Len(GetFromINI("OBObject10", "Left", DataView.FileName)) > 0 Then OBObject10.Left = GetFromINI("OBObject10", "Left", DataView.FileName)
    If Len(GetFromINI("OBObject10", "Top", DataView.FileName)) > 0 Then OBObject10.Top = GetFromINI("OBObject10", "Top", DataView.FileName)
    If GetFromINI("OBObject10", "Visible", DataView.FileName) = "True" Then OBObject10.Visible = True: If Len(GetFromINI("OBObject10", "Caption", DataView.FileName)) > 0 Then OBObject10.Caption = GetFromINI("OBObject10", "Caption", DataView.FileName)
    If Len(GetFromINI("Checkobject1", "Left", DataView.FileName)) > 0 Then Checkobject1.Left = GetFromINI("Checkobject1", "Left", DataView.FileName)
    If Len(GetFromINI("Checkobject1", "Top", DataView.FileName)) > 0 Then Checkobject1.Top = GetFromINI("Checkobject1", "Top", DataView.FileName)
    If GetFromINI("Checkobject1", "Visible", DataView.FileName) = "True" Then Checkobject1.Visible = True: If Len(GetFromINI("Checkobject1", "Caption", DataView.FileName)) > 0 Then Checkobject1.Caption = GetFromINI("Checkobject1", "Caption", DataView.FileName)
    If Len(GetFromINI("Checkobject2", "Left", DataView.FileName)) > 0 Then Checkobject2.Left = GetFromINI("Checkobject2", "Left", DataView.FileName)
    If Len(GetFromINI("Checkobject2", "Top", DataView.FileName)) > 0 Then Checkobject2.Top = GetFromINI("Checkobject2", "Top", DataView.FileName)
    If GetFromINI("Checkobject2", "Visible", DataView.FileName) = "True" Then Checkobject2.Visible = True: If Len(GetFromINI("Checkobject2", "Caption", DataView.FileName)) > 0 Then Checkobject2.Caption = GetFromINI("Checkobject2", "Caption", DataView.FileName)
    If Len(GetFromINI("Checkobject3", "Left", DataView.FileName)) > 0 Then Checkobject3.Left = GetFromINI("Checkobject3", "Left", DataView.FileName)
    If Len(GetFromINI("Checkobject3", "Top", DataView.FileName)) > 0 Then Checkobject3.Top = GetFromINI("Checkobject3", "Top", DataView.FileName)
    If GetFromINI("Checkobject3", "Visible", DataView.FileName) = "True" Then Checkobject3.Visible = True: If Len(GetFromINI("Checkobject3", "Caption", DataView.FileName)) > 0 Then Checkobject3.Caption = GetFromINI("Checkobject3", "Caption", DataView.FileName)
    If Len(GetFromINI("Checkobject4", "Left", DataView.FileName)) > 0 Then Checkobject4.Left = GetFromINI("Checkobject4", "Left", DataView.FileName)
    If Len(GetFromINI("Checkobject4", "Top", DataView.FileName)) > 0 Then Checkobject4.Top = GetFromINI("Checkobject4", "Top", DataView.FileName)
    If GetFromINI("Checkobject4", "Visible", DataView.FileName) = "True" Then Checkobject4.Visible = True: If Len(GetFromINI("Checkobject4", "Caption", DataView.FileName)) > 0 Then Checkobject4.Caption = GetFromINI("Checkobject4", "Caption", DataView.FileName)
    If Len(GetFromINI("Checkobject5", "Left", DataView.FileName)) > 0 Then Checkobject5.Left = GetFromINI("Checkobject5", "Left", DataView.FileName)
    If Len(GetFromINI("Checkobject5", "Top", DataView.FileName)) > 0 Then Checkobject5.Top = GetFromINI("Checkobject5", "Top", DataView.FileName)
    If GetFromINI("Checkobject5", "Visible", DataView.FileName) = "True" Then Checkobject5.Visible = True: If Len(GetFromINI("Checkobject5", "Caption", DataView.FileName)) > 0 Then Checkobject5.Caption = GetFromINI("Checkobject5", "Caption", DataView.FileName)
    If Len(GetFromINI("Checkobject6", "Left", DataView.FileName)) > 0 Then Checkobject6.Left = GetFromINI("Checkobject6", "Left", DataView.FileName)
    If Len(GetFromINI("Checkobject6", "Top", DataView.FileName)) > 0 Then Checkobject6.Top = GetFromINI("Checkobject6", "Top", DataView.FileName)
    If GetFromINI("Checkobject6", "Visible", DataView.FileName) = "True" Then Checkobject6.Visible = True: If Len(GetFromINI("Checkobject6", "Caption", DataView.FileName)) > 0 Then Checkobject6.Caption = GetFromINI("Checkobject6", "Caption", DataView.FileName)
    If Len(GetFromINI("Checkobject7", "Left", DataView.FileName)) > 0 Then Checkobject7.Left = GetFromINI("Checkobject7", "Left", DataView.FileName)
    If Len(GetFromINI("Checkobject7", "Top", DataView.FileName)) > 0 Then Checkobject7.Top = GetFromINI("Checkobject7", "Top", DataView.FileName)
    If GetFromINI("Checkobject7", "Visible", DataView.FileName) = "True" Then Checkobject7.Visible = True: If Len(GetFromINI("Checkobject7", "Caption", DataView.FileName)) > 0 Then Checkobject7.Caption = GetFromINI("Checkobject7", "Caption", DataView.FileName)
    If Len(GetFromINI("Checkobject8", "Left", DataView.FileName)) > 0 Then Checkobject8.Left = GetFromINI("Checkobject8", "Left", DataView.FileName)
    If Len(GetFromINI("Checkobject8", "Top", DataView.FileName)) > 0 Then Checkobject8.Top = GetFromINI("Checkobject8", "Top", DataView.FileName)
    If GetFromINI("Checkobject8", "Visible", DataView.FileName) = "True" Then Checkobject8.Visible = True: If Len(GetFromINI("Checkobject8", "Caption", DataView.FileName)) > 0 Then Checkobject8.Caption = GetFromINI("Checkobject8", "Caption", DataView.FileName)
    If Len(GetFromINI("Checkobject9", "Left", DataView.FileName)) > 0 Then Checkobject9.Left = GetFromINI("Checkobject9", "Left", DataView.FileName)
    If Len(GetFromINI("Checkobject9", "Top", DataView.FileName)) > 0 Then Checkobject9.Top = GetFromINI("Checkobject9", "Top", DataView.FileName)
    If GetFromINI("Checkobject9", "Visible", DataView.FileName) = "True" Then Checkobject9.Visible = True: If Len(GetFromINI("Checkobject9", "Caption", DataView.FileName)) > 0 Then Checkobject9.Caption = GetFromINI("Checkobject9", "Caption", DataView.FileName)
    If Len(GetFromINI("Checkobject10", "Left", DataView.FileName)) > 0 Then Checkobject10.Left = GetFromINI("Checkobject10", "Left", DataView.FileName)
    If Len(GetFromINI("Checkobject10", "Top", DataView.FileName)) > 0 Then Checkobject10.Top = GetFromINI("Checkobject10", "Top", DataView.FileName)
    If GetFromINI("Checkobject10", "Visible", DataView.FileName) = "True" Then Checkobject10.Visible = True: If Len(GetFromINI("Checkobject10", "Caption", DataView.FileName)) > 0 Then Checkobject10.Caption = GetFromINI("Checkobject10", "Caption", DataView.FileName)
    If Len(GetFromINI("Form", "Width", DataView.FileName)) > 0 Then Me.Width = GetFromINI("Form", "Width", DataView.FileName)
    If Len(GetFromINI("Form", "Height", DataView.FileName)) > 0 Then Me.Height = GetFromINI("Form", "Height", DataView.FileName)
    Me.Caption = "Mike's Move Object's Example - [" & DataView.FileName & "]"
ExitonError: Exit Sub
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Command1.SetFocus
'Set's focus on command1
End Sub

'Command2_Click lets the user save his Project
'This saves all the visible objects's: Caption,Left,Top and Visibility
'To a .dat file using CommonDialog (DataView)
'It also saves the Form's Width and Height incase the user
'Resizes the form to move the objects elsewhere on the form

Private Sub Command2_Click()
On Error GoTo ExitonError
    DataView.FileName = ""
    DataView.Filter = "Moved Objects(*.dat)|*.dat": DataView.InitDir = App.Path & "\"
    DataView.ShowSave: DataView.DialogTitle = "Save Moved Objects:"
    If DataView.FileName = "" Then Exit Sub
    If CBObject1.Visible = True Then WriteToINI "CBObject1", "Visible", "True", DataView.FileName: WriteToINI "CBObject1", "Left", CBObject1.Left, DataView.FileName: WriteToINI "CBObject1", "Top", CBObject1.Top, DataView.FileName: WriteToINI "CBObject1", "Caption", CBObject1.Caption, DataView.FileName
    If CBObject2.Visible = True Then WriteToINI "CBObject2", "Visible", "True", DataView.FileName: WriteToINI "CBObject2", "Left", CBObject2.Left, DataView.FileName: WriteToINI "CBObject2", "Top", CBObject2.Top, DataView.FileName: WriteToINI "CBObject2", "Caption", CBObject2.Caption, DataView.FileName
    If CBObject3.Visible = True Then WriteToINI "CBObject3", "Visible", "True", DataView.FileName: WriteToINI "CBObject3", "Left", CBObject3.Left, DataView.FileName: WriteToINI "CBObject3", "Top", CBObject3.Top, DataView.FileName: WriteToINI "CBObject3", "Caption", CBObject3.Caption, DataView.FileName
    If CBObject4.Visible = True Then WriteToINI "CBObject4", "Visible", "True", DataView.FileName: WriteToINI "CBObject4", "Left", CBObject4.Left, DataView.FileName: WriteToINI "CBObject4", "Top", CBObject4.Top, DataView.FileName: WriteToINI "CBObject4", "Caption", CBObject4.Caption, DataView.FileName
    If CBObject5.Visible = True Then WriteToINI "CBObject5", "Visible", "True", DataView.FileName: WriteToINI "CBObject5", "Left", CBObject5.Left, DataView.FileName: WriteToINI "CBObject5", "Top", CBObject5.Top, DataView.FileName: WriteToINI "CBObject5", "Caption", CBObject5.Caption, DataView.FileName
    If CBObject6.Visible = True Then WriteToINI "CBObject6", "Visible", "True", DataView.FileName: WriteToINI "CBObject6", "Left", CBObject6.Left, DataView.FileName: WriteToINI "CBObject6", "Top", CBObject6.Top, DataView.FileName: WriteToINI "CBObject6", "Caption", CBObject6.Caption, DataView.FileName
    If CBObject7.Visible = True Then WriteToINI "CBObject7", "Visible", "True", DataView.FileName: WriteToINI "CBObject7", "Left", CBObject7.Left, DataView.FileName: WriteToINI "CBObject7", "Top", CBObject7.Top, DataView.FileName: WriteToINI "CBObject7", "Caption", CBObject7.Caption, DataView.FileName
    If CBObject8.Visible = True Then WriteToINI "CBObject8", "Visible", "True", DataView.FileName: WriteToINI "CBObject8", "Left", CBObject8.Left, DataView.FileName: WriteToINI "CBObject8", "Top", CBObject8.Top, DataView.FileName: WriteToINI "CBObject8", "Caption", CBObject8.Caption, DataView.FileName
    If CBObject9.Visible = True Then WriteToINI "CBObject9", "Visible", "True", DataView.FileName: WriteToINI "CBObject9", "Left", CBObject9.Left, DataView.FileName: WriteToINI "CBObject9", "Top", CBObject9.Top, DataView.FileName: WriteToINI "CBObject9", "Caption", CBObject9.Caption, DataView.FileName
    If CBObject10.Visible = True Then WriteToINI "CBObject10", "Visible", "True", DataView.FileName: WriteToINI "CBObject10", "Left", CBObject10.Left, DataView.FileName: WriteToINI "CBObject10", "Top", CBObject10.Top, DataView.FileName: WriteToINI "CBObject10", "Caption", CBObject10.Caption, DataView.FileName
    If OBObject1.Visible = True Then WriteToINI "OBObject1", "Visible", "True", DataView.FileName: WriteToINI "OBObject1", "Left", OBObject1.Left, DataView.FileName: WriteToINI "OBObject1", "Top", OBObject1.Top, DataView.FileName: WriteToINI "OBObject1", "Caption", OBObject1.Caption, DataView.FileName
    If OBObject2.Visible = True Then WriteToINI "OBObject2", "Visible", "True", DataView.FileName: WriteToINI "OBObject2", "Left", OBObject2.Left, DataView.FileName: WriteToINI "OBObject2", "Top", OBObject2.Top, DataView.FileName: WriteToINI "OBObject2", "Caption", OBObject2.Caption, DataView.FileName
    If OBObject3.Visible = True Then WriteToINI "OBObject3", "Visible", "True", DataView.FileName: WriteToINI "OBObject3", "Left", OBObject3.Left, DataView.FileName: WriteToINI "OBObject3", "Top", OBObject3.Top, DataView.FileName: WriteToINI "OBObject3", "Caption", OBObject3.Caption, DataView.FileName
    If OBObject4.Visible = True Then WriteToINI "OBObject4", "Visible", "True", DataView.FileName: WriteToINI "OBObject4", "Left", OBObject4.Left, DataView.FileName: WriteToINI "OBObject4", "Top", OBObject4.Top, DataView.FileName: WriteToINI "OBObject4", "Caption", OBObject4.Caption, DataView.FileName
    If OBObject5.Visible = True Then WriteToINI "OBObject5", "Visible", "True", DataView.FileName: WriteToINI "OBObject5", "Left", OBObject5.Left, DataView.FileName: WriteToINI "OBObject5", "Top", OBObject5.Top, DataView.FileName: WriteToINI "OBObject5", "Caption", OBObject5.Caption, DataView.FileName
    If OBObject6.Visible = True Then WriteToINI "OBObject6", "Visible", "True", DataView.FileName: WriteToINI "OBObject6", "Left", OBObject6.Left, DataView.FileName: WriteToINI "OBObject6", "Top", OBObject6.Top, DataView.FileName: WriteToINI "OBObject6", "Caption", OBObject6.Caption, DataView.FileName
    If OBObject7.Visible = True Then WriteToINI "OBObject7", "Visible", "True", DataView.FileName: WriteToINI "OBObject7", "Left", OBObject7.Left, DataView.FileName: WriteToINI "OBObject7", "Top", OBObject7.Top, DataView.FileName: WriteToINI "OBObject7", "Caption", OBObject7.Caption, DataView.FileName
    If OBObject8.Visible = True Then WriteToINI "OBObject8", "Visible", "True", DataView.FileName: WriteToINI "OBObject8", "Left", OBObject8.Left, DataView.FileName: WriteToINI "OBObject8", "Top", OBObject8.Top, DataView.FileName: WriteToINI "OBObject8", "Caption", OBObject8.Caption, DataView.FileName
    If OBObject9.Visible = True Then WriteToINI "OBObject9", "Visible", "True", DataView.FileName: WriteToINI "OBObject9", "Left", OBObject9.Left, DataView.FileName: WriteToINI "OBObject9", "Top", OBObject9.Top, DataView.FileName: WriteToINI "OBObject9", "Caption", OBObject9.Caption, DataView.FileName
    If OBObject10.Visible = True Then WriteToINI "OBObject10", "Visible", "True", DataView.FileName: WriteToINI "OBObject10", "Left", OBObject10.Left, DataView.FileName: WriteToINI "OBObject10", "Top", OBObject10.Top, DataView.FileName: WriteToINI "OBObject10", "Caption", OBObject10.Caption, DataView.FileName
    If Checkobject1.Visible = True Then WriteToINI "Checkobject1", "Visible", "True", DataView.FileName: WriteToINI "Checkobject1", "Left", Checkobject1.Left, DataView.FileName: WriteToINI "Checkobject1", "Top", Checkobject1.Top, DataView.FileName: WriteToINI "Checkobject1", "Caption", Checkobject1.Caption, DataView.FileName
    If Checkobject2.Visible = True Then WriteToINI "Checkobject2", "Visible", "True", DataView.FileName: WriteToINI "Checkobject2", "Left", Checkobject2.Left, DataView.FileName: WriteToINI "Checkobject2", "Top", Checkobject2.Top, DataView.FileName: WriteToINI "Checkobject2", "Caption", Checkobject2.Caption, DataView.FileName
    If Checkobject3.Visible = True Then WriteToINI "Checkobject3", "Visible", "True", DataView.FileName: WriteToINI "Checkobject3", "Left", Checkobject3.Left, DataView.FileName: WriteToINI "Checkobject3", "Top", Checkobject3.Top, DataView.FileName: WriteToINI "Checkobject3", "Caption", Checkobject3.Caption, DataView.FileName
    If Checkobject4.Visible = True Then WriteToINI "Checkobject4", "Visible", "True", DataView.FileName: WriteToINI "Checkobject4", "Left", Checkobject4.Left, DataView.FileName: WriteToINI "Checkobject4", "Top", Checkobject4.Top, DataView.FileName: WriteToINI "Checkobject4", "Caption", Checkobject4.Caption, DataView.FileName
    If Checkobject5.Visible = True Then WriteToINI "Checkobject5", "Visible", "True", DataView.FileName: WriteToINI "Checkobject5", "Left", Checkobject5.Left, DataView.FileName: WriteToINI "Checkobject5", "Top", Checkobject5.Top, DataView.FileName: WriteToINI "Checkobject5", "Caption", Checkobject5.Caption, DataView.FileName
    If Checkobject6.Visible = True Then WriteToINI "Checkobject6", "Visible", "True", DataView.FileName: WriteToINI "Checkobject6", "Left", Checkobject6.Left, DataView.FileName: WriteToINI "Checkobject6", "Top", Checkobject6.Top, DataView.FileName: WriteToINI "Checkobject6", "Caption", Checkobject6.Caption, DataView.FileName
    If Checkobject7.Visible = True Then WriteToINI "Checkobject7", "Visible", "True", DataView.FileName: WriteToINI "Checkobject7", "Left", Checkobject7.Left, DataView.FileName: WriteToINI "Checkobject7", "Top", Checkobject7.Top, DataView.FileName: WriteToINI "Checkobject7", "Caption", Checkobject7.Caption, DataView.FileName
    If Checkobject8.Visible = True Then WriteToINI "Checkobject8", "Visible", "True", DataView.FileName: WriteToINI "Checkobject8", "Left", Checkobject8.Left, DataView.FileName: WriteToINI "Checkobject8", "Top", Checkobject8.Top, DataView.FileName: WriteToINI "Checkobject8", "Caption", Checkobject8.Caption, DataView.FileName
    If Checkobject9.Visible = True Then WriteToINI "Checkobject9", "Visible", "True", DataView.FileName: WriteToINI "Checkobject9", "Left", Checkobject9.Left, DataView.FileName: WriteToINI "Checkobject9", "Top", Checkobject9.Top, DataView.FileName: WriteToINI "Checkobject9", "Caption", Checkobject9.Caption, DataView.FileName
    If Checkobject10.Visible = True Then WriteToINI "Checkobject10", "Visible", "True", DataView.FileName: WriteToINI "Checkobject10", "Left", Checkobject10.Left, DataView.FileName: WriteToINI "Checkobject10", "Top", Checkobject10.Top, DataView.FileName: WriteToINI "Checkobject10", "Caption", Checkobject10.Caption, DataView.FileName
  WriteToINI "Form", "Width", Me.Width, DataView.FileName
WriteToINI "Form", "Height", Me.Height, DataView.FileName
Exit Sub
ExitonError: MsgBox "There was an error saving the selected file.", vbSystemModal + vbCritical, "Error:"
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Command2.SetFocus
'Set's focus on command2
End Sub

Private Sub Command3_Click()
    Dim x As Long
    x& = MsgBox("Are you sure you want to exit with out saving?", vbSystemModal + vbInformation + vbYesNo, "Confirm:"): If x& = vbYes Then End
'Confirms Exit
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Command3.SetFocus
'Set's focus on Command3
End Sub

Private Sub Command4_Click()
    MsgBox "To Add an object, hold your left mouse button down On the a object in the " & """" & "VB-Objects" & """" & " Frame" & vbCrLf & vbCrLf & "To Move an object, hold your left mouse button down on the object and move your mouse." & vbCrLf & vbCrLf & "To Rename/Delete an object, Right Click the object and the Rename/Delete menu will appear." & vbCrLf & vbCrLf & "Questions/Comments?" & vbCrLf & "E-mail: Mike@dev-center.com", vbSystemModal + vbInformation, "From Mike Canejo:"
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Command4.SetFocus
'Set's focus on Command4
End Sub

Private Sub Form_Load()
    StayOnTop Me
    CenterForm Me
    HoldNumber = 0
'centers form1 and keeps it on top
'And resets HoldNumber to 0
    MsgBox "Thanks for downloading my Move Object example." & vbCrLf & vbCrLf & "To Add an object, hold your left mouse button down On the a object in the " & """" & "VB-Objects" & """" & " Frame" & vbCrLf & vbCrLf & "To Move an object, hold your left mouse button down on the object and move your mouse." & vbCrLf & vbCrLf & "To Rename/Delete an object, Right Click the object and the Rename/Delete menu will appear." & vbCrLf & vbCrLf & "Thanks again for downloading and please RATE this :)" & vbCrLf & vbCrLf & "       -Mike Canejo" & vbCrLf & "[Mike@dev-center.com]", vbInformation, "From Mike Canejo:"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    AddCB.BorderStyle = 0: AddOB.BorderStyle = 0: AddCheckB.BorderStyle = 0
'Resets the Labels borderstyle to 0
'To give it the Flat Feel
End Sub

Private Sub Form_Resize()
If Me.Height < 3210 Then Me.Height = 3180
If Me.Width < 5340 Then Me.Width = 5370
    Frame1.Height = Me.Height - 400
    Frame2.Top = Me.Height - Frame2.Height - 400
    MSGLabel.Top = Me.Height - 1700
'This makes sure the Frame 1 and 2 and the msglabel
'Stays in location incase the users resizes the form
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    AddCB.BorderStyle = 0: AddOB.BorderStyle = 0: AddCheckB.BorderStyle = 0
'Resets the Labels borderstyle to 0
'To give it the Flat Feel
End Sub

Private Sub menuDelete_Click()
If HoldButton = "OBObject1" Then OBObject1.Visible = False: Exit Sub
    If HoldButton = "OBObject2" Then OBObject2.Visible = False: Exit Sub
        If HoldButton = "OBObject3" Then OBObject3.Visible = False: Exit Sub
        If HoldButton = "OBObject4" Then OBObject4.Visible = False: Exit Sub
        If HoldButton = "OBObject5" Then OBObject5.Visible = False: Exit Sub
        If HoldButton = "OBObject6" Then OBObject6.Visible = False: Exit Sub
        If HoldButton = "OBObject7" Then OBObject7.Visible = False: Exit Sub
        If HoldButton = "OBObject8" Then OBObject8.Visible = False: Exit Sub
        If HoldButton = "OBObject9" Then OBObject9.Visible = False: Exit Sub
        If HoldButton = "OBObject10" Then OBObject10.Visible = False: Exit Sub
If HoldButton = "CBObject1" Then CBObject1.Visible = False: Exit Sub
    If HoldButton = "CBObject2" Then CBObject2.Visible = False: Exit Sub
        If HoldButton = "CBObject3" Then CBObject3.Visible = False: Exit Sub
        If HoldButton = "CBObject4" Then CBObject4.Visible = False: Exit Sub
        If HoldButton = "CBObject5" Then CBObject5.Visible = False: Exit Sub
        If HoldButton = "CBObject6" Then CBObject6.Visible = False: Exit Sub
        If HoldButton = "CBObject7" Then CBObject7.Visible = False: Exit Sub
        If HoldButton = "CBObject8" Then CBObject8.Visible = False: Exit Sub
        If HoldButton = "CBObject9" Then CBObject9.Visible = False: Exit Sub
        If HoldButton = "CBObject10" Then CBObject10.Visible = False: Exit Sub
If HoldButton = "Checkobject1" Then Checkobject1.Visible = False: Exit Sub
    If HoldButton = "Checkobject2" Then Checkobject2.Visible = False: Exit Sub
        If HoldButton = "Checkobject3" Then Checkobject3.Visible = False: Exit Sub
        If HoldButton = "Checkobject4" Then Checkobject4.Visible = False: Exit Sub
        If HoldButton = "Checkobject5" Then Checkobject5.Visible = False: Exit Sub
        If HoldButton = "Checkobject6" Then Checkobject6.Visible = False: Exit Sub
        If HoldButton = "Checkobject7" Then Checkobject7.Visible = False: Exit Sub
        If HoldButton = "Checkobject8" Then Checkobject8.Visible = False: Exit Sub
    If HoldButton = "Checkobject9" Then Checkobject9.Visible = False: Exit Sub
If HoldButton = "Checkobject10" Then Checkobject10.Visible = False: Exit Sub

'This detects which Object was right clicked and deletes it
'by making the objects Visibility false
End Sub

Private Sub menuRename_Click()
On Error Resume Next
NotOnTop Me
Dim x As String
If HoldButton = "OBObject1" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", OBObject1.Caption): OBObject1.Caption = x
    If HoldButton = "OBObject2" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", OBObject2.Caption): OBObject2.Caption = x
        If HoldButton = "OBObject3" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", OBObject3.Caption): OBObject3.Caption = x
        If HoldButton = "OBObject4" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", OBObject4.Caption): OBObject4.Caption = x
        If HoldButton = "OBObject5" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", OBObject5.Caption): OBObject5.Caption = x
        If HoldButton = "OBObject6" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", OBObject6.Caption): OBObject6.Caption = x
        If HoldButton = "OBObject7" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", OBObject7.Caption): OBObject7.Caption = x
        If HoldButton = "OBObject8" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", OBObject8.Caption): OBObject8.Caption = x
    If HoldButton = "OBObject9" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", OBObject9.Caption): OBObject9.Caption = x
If HoldButton = "OBObject10" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", OBObject10.Caption): OBObject10.Caption = x
If HoldButton = "CBObject1" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", CBObject1.Caption): CBObject1.Caption = x
    If HoldButton = "CBObject2" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", CBObject2.Caption): CBObject2.Caption = x
        If HoldButton = "CBObject3" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", CBObject3.Caption): CBObject3.Caption = x
        If HoldButton = "CBObject4" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", CBObject4.Caption): CBObject4.Caption = x
        If HoldButton = "CBObject5" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", CBObject5.Caption): CBObject5.Caption = x
        If HoldButton = "CBObject6" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", CBObject6.Caption): CBObject6.Caption = x
        If HoldButton = "CBObject7" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", CBObject7.Caption): CBObject7.Caption = x
        If HoldButton = "CBObject8" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", CBObject8.Caption): CBObject8.Caption = x
    If HoldButton = "CBObject9" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", CBObject9.Caption): CBObject9.Caption = x
If HoldButton = "CBObject10" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", CBObject10.Caption): CBObject10.Caption = x
If HoldButton = "Checkobject1" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", Checkobject1.Caption): Checkobject1.Caption = x
    If HoldButton = "Checkobject2" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", Checkobject2.Caption): Checkobject2.Caption = x
        If HoldButton = "Checkobject3" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", Checkobject3.Caption): Checkobject3.Caption = x
        If HoldButton = "Checkobject4" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", Checkobject4.Caption): Checkobject4.Caption = x
        If HoldButton = "Checkobject5" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", Checkobject5.Caption): Checkobject5.Caption = x
        If HoldButton = "Checkobject6" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", Checkobject6.Caption): Checkobject6.Caption = x
        If HoldButton = "Checkobject7" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", Checkobject7.Caption): Checkobject7.Caption = x
        If HoldButton = "Checkobject8" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", Checkobject8.Caption): Checkobject8.Caption = x
    If HoldButton = "Checkobject9" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", Checkobject9.Caption): Checkobject9.Caption = x
If HoldButton = "Checkobject10" Then x = InputBox("Rename " & HoldButton & " to what?", "Object Renamer", Checkobject10.Caption): Checkobject10.Caption = x
StayOnTop Me
'This detects which Object was right clicked and Opens the
'VB's InputBox dialog to Change the objects Caption
End Sub

'The below lets the user move the objects wih their
'Mouse freely on the form but doesn't allow it to be
'Moved off of the form. This also detects if the
'Right mouse button was clicked to bring up
'The Menu1 popupmenu so the user can
'Rename or Delete the Right-Clicked Object.

Private Sub OBObject1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub
Private Sub OBObject1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove OBObject1
End Sub

Private Sub OBObject2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub
Private Sub OBObject2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove OBObject2
End Sub

Private Sub OBObject3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub
Private Sub OBObject3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove OBObject3
End Sub

Private Sub OBObject4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub
Private Sub OBObject4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove OBObject4
End Sub

Private Sub OBObject5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub
Private Sub OBObject5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove OBObject5
End Sub

Private Sub OBObject6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub
Private Sub OBObject6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove OBObject6
End Sub

Private Sub OBObject7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub
Private Sub OBObject7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove OBObject7
End Sub

Private Sub OBObject8_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub
Private Sub OBObject8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove OBObject8
End Sub

Private Sub OBObject9_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub
Private Sub OBObject9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove OBObject9
End Sub

Private Sub OBObject10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub
Private Sub OBObject10_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove OBObject10
End Sub
Private Sub checkobject1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub
Private Sub checkobject1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove Checkobject1
End Sub

Private Sub checkobject2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub
Private Sub checkobject2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove Checkobject2
End Sub

Private Sub checkobject3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub
Private Sub checkobject3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove Checkobject3
End Sub

Private Sub checkobject4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub
Private Sub checkobject4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove Checkobject4
End Sub

Private Sub checkobject5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub
Private Sub checkobject5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove Checkobject5
End Sub

Private Sub checkobject6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub
Private Sub checkobject6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove Checkobject6
End Sub

Private Sub checkobject7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub
Private Sub checkobject7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove Checkobject7
End Sub

Private Sub checkobject8_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub
Private Sub checkobject8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove Checkobject8
End Sub
Private Sub checkobject9_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub
Private Sub checkobject9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove Checkobject9
End Sub

Private Sub CheckObject10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim POINT As POINTAPI, z As Integer
    If Button = 1 Then GetCursorPos POINT: LastPoint.x = POINT.x: LastPoint.y = POINT.y: TheTracker = True
End Sub
Private Sub checkobject10_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then CBMove Checkobject10
End Sub
Private Sub Checkobject1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "Checkobject1": Checkobject1.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub Checkobject2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "Checkobject2": Checkobject2.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub Checkobject3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "Checkobject3": Checkobject3.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub Checkobject4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "Checkobject4": Checkobject4.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub Checkobject5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "Checkobject5": Checkobject5.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub Checkobject6_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "Checkobject6": Checkobject6.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub Checkobject7_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "Checkobject7": Checkobject7.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub Checkobject8_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "Checkobject8": Checkobject8.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub Checkobject9_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "Checkobject9": Checkobject9.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub Checkobject10_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "Checkobject10": Checkobject10.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub OBObject1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "OBObject1": OBObject1.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub OBObject2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "OBObject2": OBObject2.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub OBObject3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "OBObject3": OBObject3.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub OBObject4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "OBObject4": OBObject4.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub OBObject5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "OBObject5": OBObject5.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub OBObject6_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "OBObject6": OBObject6.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub OBObject7_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "OBObject7": OBObject7.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub OBObject8_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "OBObject8": OBObject8.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub OBObject9_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "OBObject9": OBObject9.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub
Private Sub OBObject10_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HoldButton = "OBObject10": OBObject10.SetFocus
    If Button = 2 Then Me.PopupMenu Menu1
End Sub

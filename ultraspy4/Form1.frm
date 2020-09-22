VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Form1.frx":08CA
   ScaleHeight     =   7020
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Window Styles"
      Height          =   1905
      Left            =   90
      TabIndex        =   34
      Top             =   5040
      Width           =   4245
      Begin VB.TextBox lblStyle 
         Height          =   1455
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   270
         Width           =   3885
      End
   End
   Begin VB.CommandButton cmdSaveIcon 
      Caption         =   "Sa&ve Icon"
      Height          =   555
      Left            =   90
      TabIndex        =   33
      Top             =   3690
      Width           =   1185
   End
   Begin VB.Frame Frame4 
      Caption         =   "Extract Icon"
      Height          =   1275
      Left            =   1440
      TabIndex        =   30
      Top             =   3600
      Width           =   2985
      Begin VB.PictureBox picIcon 
         AutoRedraw      =   -1  'True
         Height          =   645
         Left            =   180
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   585
         ScaleWidth      =   585
         TabIndex        =   31
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label13 
         Caption         =   "To extract an icon, simply drag the desired application or library over the picturebox to the left."
         Height          =   825
         Left            =   900
         TabIndex        =   32
         Top             =   270
         Width           =   1995
      End
   End
   Begin VB.CommandButton cmdMakeSmall 
      Caption         =   "<--"
      Height          =   555
      Left            =   8280
      TabIndex        =   26
      Top             =   2700
      Width           =   555
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   555
      Left            =   7020
      TabIndex        =   25
      Top             =   2700
      Width           =   1095
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   555
      Left            =   5760
      TabIndex        =   24
      Top             =   2700
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4050
      Top             =   1260
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save Picture"
      Filter          =   "Bitmap Files|*.bmp|Icon Files|*.ico"
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Picture"
      Height          =   555
      Left            =   90
      TabIndex        =   22
      Top             =   4320
      Width           =   1185
   End
   Begin VB.Frame Frame3 
      Caption         =   "Screenshot"
      Height          =   3435
      Left            =   4500
      TabIndex        =   21
      Top             =   3510
      Width           =   4425
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         Height          =   2985
         Left            =   180
         ScaleHeight     =   2925
         ScaleWidth      =   4005
         TabIndex        =   23
         Top             =   270
         Width           =   4065
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Altering Information"
      Height          =   2445
      Left            =   4410
      TabIndex        =   2
      Top             =   90
      Width           =   4425
      Begin VB.TextBox txtStatic 
         Height          =   285
         Left            =   1710
         TabIndex        =   18
         Text            =   "<null>"
         Top             =   990
         Width           =   2445
      End
      Begin VB.TextBox txtDis 
         Height          =   285
         Left            =   1710
         TabIndex        =   16
         Text            =   "<null>"
         Top             =   630
         Width           =   2445
      End
      Begin VB.TextBox txtKill 
         Height          =   285
         Left            =   1710
         TabIndex        =   14
         Text            =   "<null>"
         Top             =   1710
         Width           =   2445
      End
      Begin VB.TextBox txtClassText 
         Height          =   285
         Left            =   1710
         TabIndex        =   12
         Text            =   "<null>"
         Top             =   1350
         Width           =   2445
      End
      Begin VB.TextBox txtEnable 
         Height          =   285
         Left            =   1710
         TabIndex        =   11
         Text            =   "<null>"
         Top             =   270
         Width           =   2445
      End
      Begin VB.Label Label10 
         Caption         =   "Window to change:"
         Height          =   285
         Left            =   270
         TabIndex        =   17
         Top             =   1080
         Width           =   2265
      End
      Begin VB.Label Label9 
         Caption         =   "Disable Window:"
         Height          =   195
         Left            =   270
         TabIndex        =   15
         Top             =   720
         Width           =   2085
      End
      Begin VB.Label Label8 
         Caption         =   "Minimize Window:"
         Height          =   195
         Left            =   270
         TabIndex        =   13
         Top             =   1800
         Width           =   1545
      End
      Begin VB.Label Label7 
         Caption         =   "Class to change:"
         Height          =   330
         Left            =   270
         TabIndex        =   10
         Top             =   1440
         Width           =   2325
      End
      Begin VB.Label Label3 
         Caption         =   "Enable Window:"
         Height          =   240
         Left            =   270
         TabIndex        =   3
         Top             =   360
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "UltraSpy Information"
      Height          =   2445
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3975
      Begin VB.TextBox txtHandle 
         Height          =   285
         Left            =   810
         TabIndex        =   20
         Top             =   360
         Width           =   2265
      End
      Begin VB.TextBox txtParent 
         Height          =   285
         Left            =   810
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1440
         Width           =   2265
      End
      Begin VB.TextBox txtClass 
         Height          =   285
         Left            =   810
         TabIndex        =   5
         Top             =   720
         Width           =   2265
      End
      Begin VB.TextBox txtUM 
         Height          =   285
         Left            =   810
         TabIndex        =   4
         Top             =   1080
         Width           =   2265
      End
      Begin VB.PictureBox Picture1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   1
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RGB"
         Height          =   375
         Left            =   1620
         TabIndex        =   29
         Top             =   1890
         Width           =   1455
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   810
         TabIndex        =   28
         Top             =   1890
         Width           =   645
      End
      Begin VB.Label Label5 
         Caption         =   "Color:"
         Height          =   285
         Left            =   180
         TabIndex        =   27
         Top             =   1890
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Handle:"
         Height          =   285
         Left            =   180
         TabIndex        =   19
         Top             =   450
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Parent:"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   1530
         Width           =   570
      End
      Begin VB.Label Label2 
         Caption         =   "Text:"
         Height          =   180
         Left            =   180
         TabIndex        =   8
         Top             =   1170
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Class:"
         Height          =   270
         Left            =   180
         TabIndex        =   7
         Top             =   810
         Width           =   450
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim verylong As String * 100
Dim gParent As String * 100
Dim SndMsg As String * 100
Dim windowname As String * 100
Dim sztext As String * 100
Dim mousemove As Boolean
Dim Pic01 As Boolean
Dim SmallBL As Boolean









Private Sub cmdAbout_Click()
'displays messagbox about this program
MsgBox "UltraSpy v4.0 by Shimoon Technologies." & vbCrLf & vbCrLf & "Compiled on *INSERT DATE HERE*" & vbCrLf & "Coded by Armen Shimoon." & vbCrLf & "©2001 Shimoon Technologies.", vbOKOnly, "About"


End Sub

Private Sub cmdExit_Click()

'unloads the forms from memory
Unload frmSplash
Unload Form1



End Sub

Private Sub cmdMakeSmall_Click()


'code to contract and expand the size of the window
If SmallBL = True Then
    Form1.Height = 7395
    cmdMakeSmall.Caption = "<--"
    SmallBL = False
    
Else
    Form1.Height = 3800
    cmdMakeSmall.Caption = "-->"
    SmallBL = True
End If

End Sub

Private Sub cmdSave_Click()
 Dim CheckFile As Boolean
 Dim strFileName As String
 
'saves the screenshot
 CommonDialog1.ShowSave
 
 
 strFileName = CommonDialog1.FileName

 DoEvents
     On Error GoTo 20
    SavePicture Picture2.Image, strFileName
 DoEvents
MsgBox "Saved to " & strFileName, vbInformation, "Saved"

'error handler when user clicks Cancel in the commondialog
20: Exit Sub
End Sub

Private Sub cmdSaveIcon_Click()
'saves our icon we extracted
CommonDialog1.ShowSave
On Error GoTo 30
Call SavePicture(picIcon.Image, CommonDialog1.FileName)
MsgBox "Saved to " & CommonDialog1.FileName, vbInformation, "Saved"

30: Exit Sub
End Sub

Private Sub Form_Load()
Picture1.Picture = LoadResPicture(102, vbResIcon) ' loads the pointer
Form1.Icon = LoadResPicture(101, vbResIcon) ' loads the window icon
Form1.Caption = LoadResString(101) & " " & LoadResString(102) & " v4.0" & "  -  " & LoadResString(103) ' sets the window title

mousemove = False ' makes sure program knows we arent dragging the pointer

'sets the textboxes to read only
TextRO txtClass
TextRO txtParent
TextRO txtUM

Pic01 = True

'makes the window big and not small
SmallBL = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

'unloads all the forms from memory
Unload frmSplash
Unload Form1

End Sub

Private Sub picIcon_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim FileName
Dim FileIcon As Long

'gets the filename of the file dropped onto the picturebox then loads the icon
If (Data.GetFormat(vbCFFiles) = True) Then
    For Each FileName In Data.Files
        FileIcon = ExtractIcon(App.hInstance, FileName, 0)
        picIcon.Picture = Nothing
        Call DrawIcon(picIcon.hdc, 1, 1, FileIcon)
    Next FileName
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'gives the effect that we are actually dragging the pointer
Picture1.Picture = Nothing
Form1.MouseIcon = LoadResPicture(102, vbResIcon)
Form1.MousePointer = 99
'tells program to get information
mousemove = True

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim cursorPos1 As POINTAPI
   Dim wintext As String
   Dim garmon As String
   Dim gIcon As Image
   Dim OldX As Integer
   Dim OldY As Integer
   Dim ttxt As String
   Dim abc As String
   Dim WndRECT As RECT
   Dim Width1 As Integer, Height1 As Integer
   Dim sWnd As Long
   Dim wStyles As String
 If mousemove = True Then
 
    'gets the cursor position
    r = GetCursorPos(cursorPos1)
    'various functions to get information about the window under the cursor
    hWnd1 = WindowFromPoint(cursorPos1.X, cursorPos1.Y)
    r = GetClassName(hWnd1, sztext, 100)
    hWnd2 = WindowFromPoint(cursorPos1.X, cursorPos1.Y)
    p = GetWindowText(hWnd2, windowname, 100)
    hwnd3 = WindowFromPoint(cursorPos1.X, cursorPos1.Y)
    q = GetParent(hwnd3)
    Call GetWindowRect(hWnd1, WndRECT)
    v& = GetDC(hWnd1)
    Width1 = WndRECT.Left + WndRECT.Right
    Height1 = WndRECT.Top + WndRECT.Bottom
    
    'function to get the pixel color under the mouse
    Call GetPixel1(Label6, Label12)
    
    sWnd = WindowFromPoint(cursorPos1.X, cursorPos1.Y)
    
    'function to get the window styles
    Call GetStyles(sWnd, wStyles)
    lblStyle.Text = wStyles
    
    'clears the screenshot box so a new image can be put in
    Picture2.Picture = Nothing
    'paint screenshot into picturebox
    Call BitBlt(Picture2.hdc, 0, 0, 600, 500, v&, 0, 0, vbSrcCopy)
    'release screenshot from memory
    Call ReleaseDC(hWnd1, v&)
    
              'function to get unmasked text
              ttxt = Space(100)
              errval = GetCursorPos(cursorPos1)
              thwnd = WindowFromPoint(cursorPos1.X, cursorPos1.Y)
              errval = SendMessage(thwnd, WM_GETTEXT, ByVal TXT_LEN, ByVal ttxt)
              ttxt = RTrim(ttxt)
              
    'set all the textboxes
    txtUM.Text = ttxt
    txtHandle.Text = hWnd1
    txtParent.Text = q
    txtClass.Text = sztext
    


'all these functions test to see if it should alter the information of the hwnd
'as we entered it into the edit boxes in the frame "Altering Info"
If txtClass.Text = txtClassText.Text Then
    a = InputBox("New string for " & txtClass & ":", "New string")
    b = SetWindowText(hWnd1, a)
    txtClassText.Text = "<null>"
    Picture1.Picture = LoadResPicture(102, vbResIcon)
    Form1.MousePointer = 0
ElseIf txtUM.Text = Frame2.Caption Then
    Picture1.Picture = LoadResPicture(102, vbResIcon)
    Form1.MousePointer = 0
    mousemove = False
ElseIf txtKill.Text = txtUM.Text Then
    a = CloseWindow(hWnd2)
    txtKill.Text = "<null>"
    Picture1.Picture = LoadResPicture(102, vbResIcon)
    Form1.MousePointer = 0
ElseIf txtDis.Text = txtUM.Text Then
    a = EnableWindow(hWnd2, 0)
    txtDis.Text = "<null>"
    Picture1.Picture = LoadResPicture(102, vbResIcon)
    Form1.MousePointer = 0
ElseIf txtUM.Text = txtStatic.Text Then
    abc = InputBox("New string for static " & txtUM.Text, "New string")
    Call SendMessage(hWnd2, WM_SETTEXT, 0&, ByVal abc)
    txtStatic.Text = "<null>"
    Picture1.Picture = LoadResPicture(102, vbResIcon)
    Form1.MousePointer = 0
ElseIf txtUM.Text = txtEnable.Text Then
    Call EnableWindow(hWnd2, 1&)
    txtEnable.Text = "<null>"
    Picture1.Picture = LoadResPicture(102, vbResIcon)
    Form1.MousePointer = 1
End If
 
 
    
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'when the mouse is released after dragging we tell the program to simulate the pointer going back to the box
Picture1.Picture = LoadResPicture(102, vbResIcon)
Form1.MousePointer = 0
'tell the program to stop getting information about the windows
mousemove = False
End Sub


Private Function TextRO(textbx As TextBox)
'function to set the textboxes to readonly
a = SendMessage(textbx.hwnd, EM_SETREADONLY, 1, 0)
End Function






'function to get the pixel color under the mouse
Public Function GetPixel1(PaintLabel As Label, DisplayLabel As Label)

Dim cursorPos1 As POINTAPI ' variable for the cursor position
Dim hwndWindow As String ' variable for the cursor positions hwnd
Dim wndDC1 As String ' variable for the hwnd's DC
Dim dcPixel As String ' variable for the DC's pixel



a = GetCursorPos(cursorPos1) ' get the current cursor position
hwndWindow = WindowFromPoint(cursorPos1.X, cursorPos1.Y) ' from the cursor position, get the hwnd under it
wndDC1 = GetDC(hwndWindow) ' get DC from the currents hwnd
dcPixel = GetPixel(wndDC1, cursorPos1.X, cursorPos1.Y) ' get the pixel under the mouse from the DC



Dim blue&, green&, red&, colour& ' variabled to change decimal format into RGB




colour& = dcPixel ' set color variable to decimal format

If Len(Str(dcPixel)) = 2 Then ' checks if dcpixel is incorrect format
    Exit Function ' if so, exit the sub without drawing color
Else ' if format is valid, continue

blue& = Int(colour& / 65536) ' function to get the blue
green& = Int((colour& - (65536 * blue&)) / 256) ' function to get the green
red& = colour& - (blue& * 65536) - (green& * 256) ' function to get the red

colour& = RGB(red&, green&, blue&) ' set final RGB format

PaintLabel.BackColor = colour& ' paint dcPixel in RGB format to the color label
DisplayLabel.Caption = red & "," & green & "," & blue

End If ' stop asking questions, lol

Call DeleteDC(wndDC)


End Function


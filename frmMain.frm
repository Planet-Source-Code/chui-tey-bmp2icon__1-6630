VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BMP2ICON"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSelectBMP 
      Caption         =   "Choose bitmap to convert"
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   810
      Width           =   2265
   End
   Begin MSComDlg.CommonDialog cdlFileOpen 
      Left            =   2790
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.BMP"
      Filter          =   "Bitmap (*.bmp)|*.bmp"
   End
   Begin VB.Image Image1 
      Height          =   1275
      Left            =   3570
      Top             =   105
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "BMP2ICON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   2175
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2700
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   16777215
      _Version        =   327682
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=================================================================================
'   BMP2ICON
'
'   It eats its own dogfood: the icon for this program is created using BMP2ICON
'   (c) Chui Tey 2000
'
'   Email: teyc@bigfoot.com
'
'   Requires:
'       Common Dialog Control
'       Windows Common Controls-5.0 (SP2)
'
'=================================================================================

Private Sub cmdSelectBMP_Click()

    On Error GoTo ErrHandler
    cdlFileOpen.ShowOpen
        
    Dim lsICOFilename As String
    Dim lsBMPFilename As String
    lsBMPFilename = cdlFileOpen.filename
    lsICOFilename = Left(lsBMPFilename, Len(lsBMPFilename) - 3) & "ico"
    
    '   Responding to Alessandro's request:
    '   1. Show the Bitmap in the image control
    '   2. Let the user decide whether to convert
    '      to icon format
    '
    Image1.Picture = LoadPicture(lsBMPFilename)
    Dim liChoice As VbMsgBoxResult
    liChoice = MsgBox("Convert this image to icon?", vbYesNo)
    If liChoice = vbYes Then
    
        '   Call the actual conversion routine
        '
        Convert lsBMPFilename, lsICOFilename
        
        MsgBox "BMP file converted to " & lsICOFilename, vbInformation
        
    End If
    
    Exit Sub
    
ErrHandler:

    Select Case Err.Number
        Case 32755
            'Do nothing
        
        Case Else
            MsgBox "ERROR " & Err.Number & vbNewLine & Err.Description, vbExclamation
            
    End Select
    
End Sub

Private Sub Convert(ByVal asBMPFilename As String, ByVal asICOFilename As String)

    Dim a As IPictureDisp
    Set a = LoadPicture(asBMPFilename)
    
    ImageList1.ListImages.Add Picture:=a
    Set a = ImageList1.ListImages(1).ExtractIcon
    
    On Error GoTo ErrHandler
    SavePicture a, asICOFilename
    
    ImageList1.ListImages.Remove (1)
        
    Exit Sub
    
ErrHandler:

    ImageList1.ListImages.Remove (1)
    
    Select Case Err.Number
        Case 380
            Err.Raise Err.Number, Err.Source, Err.Description & vbNewLine & "The bitmap may be too big to convert to an icon format"
        Case Else
            Err.Raise Err.Number, Err.Source, Err.Description
    End Select
    
End Sub


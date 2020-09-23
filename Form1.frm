VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   6990
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "HOW??"
      Height          =   495
      Left            =   6870
      TabIndex        =   1
      Top             =   6315
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TRY DROPPING DRAGGED PICTURE ANYWHERE IN THIS FORM..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2250
      Left            =   3285
      TabIndex        =   0
      Top             =   4140
      Width           =   4755
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is an example of a code which enables your application to use the OLEDragDrop method in loading files faster and a lot easier.
'Codes by: Ronald Borla email: rdb_dvo@yahoo.com

'Note that you have to set the OLEDropMode Property of an object to 1-Manual, to be
'able to perform this code and use the OLEDragDrop Event of the object

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 1
'This part gets the filename of the file dropped
Me.Picture = LoadPicture(Data.Files.Item(1))
Exit Sub
1:
'Because the code above loads a valid picture only, it errors when the selected file is not a supported image file
MsgBox "The dropped file is not a picture or is not supported in this application!", vbCritical, "Error"
End Sub

'It does not mean that you can only use a Form object for this code, there are also other
'default objects in VB which supports the OLEDropMode property and OLEDragDrop Event.

'To drop multiple files at once, use the following codes instead:

'------------------------------------
'Dim i As Integer, Files() As String
'
'For i = 1 To Data.Files.Count
'    ReDim Preserve Files(i)
'    Files(i) = Data.Files.Item(i)
'Next i
'------------------------------------

'The filename of each file dropped is stored in the array "Files()"
'You can then load any of the files dropped anyway you want.

'Note: The minimum index number for "Data.Files.Item" is 1. It does not start with 0 like any other arrays

'Another thing: To preempt the user that the selected file is not a supported file, use the following code in the OLEDragOver Event:

'-------------------------------------------------------------------------------------------
'If Right(Data.Files.Item(1), 4) <> ".jpg" And Right(Data.Files.Item(1), 5) <> ".jpeg" Then
'    Effect = 0
'    'When effect is = 0, the mouse icon will change into the "NO" icon, which will stop the drop event, and avoid further file loading
'End If
'-------------------------------------------------------------------------------------------

Private Sub Command1_Click()
MsgBox "While application running, select any image file from any opened folder then drag and drop it on the form, the selected picture will be loaded in the form", vbInformation, "OLE Sample"
MsgBox "Note that even if you select multiple files, it will only load the first file selected", vbInformation, "OLE Sample"
MsgBox "To drop multiple files at once, see code in Form1", vbInformation, "OLE Sample"
End Sub

'I hope this code will help you in most of your applications especially in applications having output files...

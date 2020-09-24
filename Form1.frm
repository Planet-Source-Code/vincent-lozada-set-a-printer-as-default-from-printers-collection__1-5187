VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Default Printer"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Make Default"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Select a printer:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'To verify that the printer has been changed to the default, open your "Printers" folder
'in your control panel and select a printer and press the right mouse button to access
'the context menu. Take notice of "Set As Default" check mark for the selected printer
'when your run this app...

Private cSetPrinter As New cSetDfltPrinter

Private Sub Command1_Click()
    Dim sMsg As String
    Dim DeviceName As String
    
    If List1.SelCount = 1 Then
        DeviceName = List1.List(List1.ListIndex)
        If cSetPrinter.SetPrinterAsDefault(DeviceName) Then
            sMsg = DeviceName & " has successfully been set as the default printer."
        Else
            sMsg = DeviceName & " has failed to be set as the default printer."
        End If
        MsgBox sMsg, vbExclamation, App.Title
    Else
        MsgBox "Please select a printer from the list.", vbInformation, App.Title
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    For i = 0 To Printers.Count - 1
        List1.AddItem Printers(i).DeviceName
    Next i
End Sub

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Membuka Form dari Variable String"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  'Deklarasi variabel bertipe Form
  Dim frm As Form
  'Ambil nama form yang akan dibuka
  'Dimisalkan data "Form2" berasal dari database
  DataForm = "Form2"
  'Buat/tambahkan acuan nama form ke variabel frm
  Set frm = Forms.Add(DataForm)
  'Tampilkan form
  frm.Show
End Sub


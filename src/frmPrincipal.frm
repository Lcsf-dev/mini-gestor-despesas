VERSION 5.00
Begin VB.Form frmPrincipal 
   Caption         =   "Mini Gestor de Despesas"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   21015
   ScaleWidth      =   38160
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If AbrirConexao() Then
        MsgBox "Conexão com DB_MiniDespesas realizada com sucesso!", vbInformation, "Sucesso"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FecharConexao
End Sub

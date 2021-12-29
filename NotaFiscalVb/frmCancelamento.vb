Imports System.Threading
Imports System.Data.SqlClient
Imports System.ComponentModel

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.IO
Imports Microsoft.VisualBasic
Imports System.Security.Cryptography.Xml
Imports System.Security.Cryptography.X509Certificates
Imports System.ServiceModel
Imports System.Xml
Public Class frmCancelamento
    Private sChave As String
    Private sProtocolo As String

    Private Function CancelarNota() As Boolean
        Dim Canc As New Cancelamento
        Dim Lote As String = ""

        Try
            Canc.Versao = VersaoNota
            Canc.chNFe = sChave
            Canc.dhEvento = txtData.Text
            Canc.nProt = sProtocolo
            Canc.xJust = txtJustificativa.Text

            'Fazer uma consulta na tabela de notas para trazer dados da nota


            If Canc.CancelarNota(Lote, "F") Then

                'Gravar retorno do Cancelamento

            End If


        Catch ex As Exception

        End Try
        Return True
    End Function

    Private Function BuscarLoteNota() As Boolean
        Dim sql As String
        Dim dr As SqlDataReader

        Try

            'Banco.Conecxao

            sql = ""

            sql = "select * from LoteNota where fkNumNota =  " & txtNota.Text & " and status =  100"
            'dr =  Banco.CriaDataReader(Banco.bd,sql)

            If dr.Read() Then
                sChave = dr("ChaveNfe").ToString()
                sProtocolo = dr("NumProtocolo").ToString()
            Else
                MessageBox.Show("Nenhuma nota foi encontrada")
                Return False
            End If

            dr.Close()
            'banco.bd.close()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        End Try
        Return True
    End Function

    Private Sub LimpaCampos()
        txtNota.Text = ""
        txtData.Text = ""
        txtJustificativa.Text = ""
        txtNota.Focus()
    End Sub
    Private Sub btnNovo_Click(sender As Object, e As EventArgs) Handles btnNovo.Click
        Call LimpaCampos()
    End Sub

    Private Sub btnCancelar_Click(sender As Object, e As EventArgs) Handles btnCancelar.Click
        If BuscarLoteNota() Then
            If CancelarNota() Then
                MessageBox.Show("Nota Cancelada com Sucesso!!!")
                Call LimpaCampos()
            End If
        End If
    End Sub
End Class
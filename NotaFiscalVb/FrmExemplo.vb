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

Public Class FrmExemplo
    Private Nota As Integer
    Private Numerolote As Integer


    Private cStat As String
    Private xMotivo As String
    Private dhRecibo As String
    Private nRecibo As String
    Private Protocolo As String

    Private chNfe As String
    Private digVal As String
    Private cStatNfe As String
    Private xMotivoNfe As String
    Private VerAplic As String

    Private XmlNota As New AssinarXml()
    'primeira função a ser chamada depois de gerar o xml da nota
    Private Function GerarLote(NumLote As String, NomeCert As String, sSicrona As Boolean) As Boolean
        Dim bLote As Boolean
        Dim sql As String


        Try
            bLote = True

            Dim cls As New AutorizaNFe()

            Dim Caminho As String = sNota

            If Not cls.AutorizandoNota(Caminho, NumLote, NomeCert, True) Then
                bLote = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            bLote = False
        End Try
        Return bLote
    End Function
    'segunda chamada
    Private Function Consultalote(NumLote As Integer, NomeCert As String, NumNota As Integer) As Boolean
        Dim sql, sql2 As String
        Dim Dr As SqlDataReader
        Dim Recibo As String
        Dim cls As New RetAutorizaNFe()
        Dim Status As String
        Dim consLote As Boolean
        Try
            consLote = True
            'Fazer essa consulta na tabela Lote
            'Lembrando que ja deve estar Gravada o retorno da consult a do lote
            'sua conecxão
            sql = String.Empty
            sql = " select * from Lote where NumLote = " + NumLote + " "
            sql += " and Recibo is not null "
            ' aqui fica o datareader

            If Dr.Read() Then

                'Recibo = Dr["Recibo"].ToString().Trim()

                Thread.Sleep(15000)
                While Status = "105" Or Status = ""
                    If Not cls.RetornandoNFe(Recibo, NomeCert, NumLote, Status) Then
                        consLote = False
                    End If
                    Thread.Sleep(15000)
                End While

                Dr.Close()
            Else
                Dr.Close()
                consLote = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            consLote = False
        End Try
        Return consLote
    End Function
    Private Function GeraXmlNota() As Boolean
        Dim Cert As X509Certificate
        Dim ClsAssinatura As New AssinarXml()
        Dim NF As New GerarXmlNota()

        Try

            If NomeCertificado = "" Then
                Cert = ClsAssinatura.BuscarNome("")
                NomeCertificado = Cert.Subject
            End If

            'Identificação da Nota

            NF.CNPJ = "11111111111111"
            NF.cUF = "21" 'caso Maranão usar 21
            NF.cNF = "11111111"
            NF.natOp = "Natureza da Operação so a descrição"
            NF.indPag = "0" '0=Pagamento à vista; 1=Pagamento a prazo; 2=Outros. 
            NF.smod = "55"
            NF.serie = "001"
            NF.nNF = "25"
            NF.dhEmi = "2020-03-01T12:00:00" 'formato ToString("yyyy-MM-ddThh:mm:ss")
            NF.dhSaiEnt = "2020-03-01T12:00:00"  'formato ToString("yyyy-MM-ddThh:mm:ss")
            NF.tpNF = "1"
            NF.idDest = "1" 'colocar 1 se for dentro do estado e 2 se for fora do estado
            NF.cMunFG = "2111300"
            NF.tpImp = "1"
            NF.tpRmis = "1"
            NF.tpAmb = TpAmb.ToString().Trim()
            NF.finNFe = "1"
            NF.finFinal = "1"
            NF.indPres = "0"
            NF.procEmi = "0"
            NF.verPros = "1.0.0.0"

            NF.addCabecalho()

            'identificação do emitente 

            NF.CNPJ = "11111111111112"
            NF.xNome = "Razão Social"
            NF.xFant = "nome de fantazia"
            NF.xLgr = "Logradouro"
            NF.nro = "Numero"
            NF.xBairro = "Bairro.Trim()"
            NF.cMun = "2111300" '"Codigo do Municipio conforme tabela do IBGE"
            NF.xMun = "SÃO LUIS"
            NF.UF = "MA"
            NF.CEP = "CEP"
            NF.cPais = "1058"
            NF.xPais = "BRASIL"
            NF.fone = "Telefone"
            NF.IE = "Inscrição estadual"
            NF.IM = "inscrição Municipal"
            NF.CNAE = "verificar com contador"
            NF.CRT = "3"

            NF.addEmitente()


            'Identificação do Destinatario

            NF.CPFDest = "CPF"
            NF.xNomeDest = "Nome"
            NF.xLgrDest = "Endereco"
            NF.nroDest = "numero"
            NF.xBairroDest = "Bairro"
            NF.cMunDest = "Codigo do Municipio tabela ibge"
            NF.xMunDest = "Cidade.Trim()"
            NF.UFDest = "Uf"
            NF.CEPDest = "Cep"
            NF.cPaisDest = "1058"
            NF.xPaisDest = "BRASIL"
            NF.foneDest = "telefone"
            NF.indIEdest = "9"

            NF.addDestinatario()

            'item da Nota

            'Produtos ou serviços
            NF.nItem = "1" 'numero do Item encremetar se tiver mais de um
            NF.cProd = "Codigo do Produto"
            NF.xProd = "Descrição do Produto"
            NF.NCM = "00"
            NF.CFOP = "5933"
            NF.uCom = "UM"
            NF.qCom = "1"
            NF.vUnCom = "10,00"
            NF.vProd = "10,00"
            'NF.cEANTrib = "XXX"
            NF.uTrib = "UM"
            NF.qTrib = "1"
            NF.vUnTrib = "10,00"
            NF.indTot = "1"

            NF.addItemNota()
            'impostos 
            NF.vUnTrib = "0,00"
            NF.vICMS = "0,00"
            NF.addTotalNota()

            NF.FinalizaXml(NomeCertificado)

        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function



    Private Sub Retornalote()
        Dim sql As String
        Dim Dr As SqlDataReader

        Numerolote = 0

        sql = "select max(numlote) + 1 as total from Lote "
        'passa a conexao para o datareader

        If Dr.Read() Then
            'Numerolote =  Convert.ToInt32(Dr["total"])
        End If
        Dr.Close()


    End Sub
    Private Function GravarLoteNota(iNota As Integer, Chave As String) As Boolean
        Dim sql As String = ""
        Dim Dr As SqlDataReader
        Try
            Call Retornalote()

            sql = "insert into Lote(NumLote) values ( " & Numerolote & " ) "
            'passa a conexao para o datareader
            Dr.Close()

            sql = ""

            sql = "insert LoteNota(fkNumLote,fkNumNota,ChaveNFe) values (" & Numerolote & ", " & iNota & ",'" & Chave & "')"
            'passa a conexao para o datareader
            Dr.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        End Try

        Return True

    End Function
    Private Sub bynEnviar_Click(sender As Object, e As EventArgs) Handles bynEnviar.Click

        'Nota recebe o numero da nota, '
        Dim bSicrona As Boolean

        If CheckBox1.Checked Then
            bSicrona = True
        Else
            bSicrona = False

        End If

        If GeraXmlNota() Then
            If GravarLoteNota(Nota, sChaveNota) Then
                If GerarLote(Numerolote, NomeCertificado, bSicrona) Then
                    If bSicrona = False Then
                        If Consultalote(Numerolote, NomeCertificado, Nota) Then
                            MessageBox.Show("Nota Enviada com sucesso!!!")
                        End If
                    Else
                        MessageBox.Show("Nota Enviada com sucesso!!!")
                    End If
                End If

            End If
        End If

    End Sub

    Private Sub FrmExemplo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim MERDA As String = "TESTE&"

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim NomeArqRet As String = "C:\Teste\Ret_AutorizaLote_314202112.xml"

        Dim sDados As String = ""
        Dim sr As StreamReader


        Dim doc2 As New XmlDocument

        sr = File.OpenText(NomeArqRet)
        sDados = sr.ReadToEnd()
        sr.Close()

        doc2.PreserveWhitespace = False

        doc2.LoadXml(sDados)

        Dim No, sbno, xy, z As XmlNode

        For Each No In doc2.ChildNodes
            For Each sbno In No.ChildNodes
                If sbno.Name = "cStat" Then
                    cStat = sbno.InnerText
                ElseIf sbno.Name = "xMotivo" Then
                    xMotivo = sbno.InnerText
                ElseIf sbno.Name = "dhRecbto" Then
                    dhRecibo = Convert.ToDateTime(sbno.InnerText).ToString("yyyy-MM-dd")
                ElseIf (cStat = "104") And Not IsDBNull(xMotivo) And Not IsDBNull(dhRecibo) Then
                    For Each xy In sbno.ChildNodes
                        For Each z In xy.ChildNodes
                            If z.Name = "verAplic" Then
                                VerAplic = z.InnerText
                            ElseIf z.Name = "chNFe" Then
                                chNfe = z.InnerText
                            ElseIf z.Name = "digVal" Then
                                digVal = z.InnerText
                            ElseIf z.Name = "cStat" Then
                                cStatNfe = z.InnerText
                            ElseIf z.Name = "nProt" Then
                                Protocolo = z.InnerText
                            ElseIf z.Name = "xMotivo" Then
                                xMotivoNfe = z.InnerText
                            End If
                        Next
                    Next
                End If
            Next
        Next


    End Sub
End Class

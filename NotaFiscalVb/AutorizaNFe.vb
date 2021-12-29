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


Public Class AutorizaNFe

    Private _cStat As String
    Private _xMotivo As String
    Private _dhRecibo As String
    Private _nRecibo As String
    Private _Protocolo As String

    Private _chNfe As String
    Private _digVal As String
    Private _cStatNfe As String
    Private _xMotivoNfe As String
    Private _VerAplic As String

    Private XmlNota As New AssinarXml()
    Public Property digVal As String
        Get
            Return _digVal
        End Get
        Set(value As String)
            _digVal = value
        End Set
    End Property

    Public Property VerAplic As String
        Get
            Return _VerAplic
        End Get
        Set(value As String)
            _VerAplic = value
        End Set
    End Property
    Public Property chNfe As String
        Get
            Return _chNfe
        End Get
        Set(value As String)
            _chNfe = value
        End Set
    End Property

    Public Property cStat As String
        Get
            Return _cStat
        End Get
        Set(value As String)
            _cStat = value
        End Set
    End Property
    Public Property cStatNfe As String
        Get
            Return _cStatNfe
        End Get
        Set(value As String)
            _cStatNfe = value
        End Set
    End Property

    Public Property xMotivo As String
        Get
            Return _xMotivo
        End Get
        Set(value As String)
            _xMotivo = value
        End Set
    End Property

    Public Property xMotivoNfe As String
        Get
            Return _xMotivoNfe
        End Get
        Set(value As String)
            _xMotivoNfe = value
        End Set
    End Property

    Public Property dhRecibo As String
        Get
            Return _dhRecibo
        End Get
        Set(value As String)
            _dhRecibo = value
        End Set
    End Property

    Public Property nRecibo As String
        Get
            Return _nRecibo
        End Get
        Set(value As String)
            _nRecibo = value
        End Set
    End Property

    Public Property Protocolo As String
        Get
            Return _Protocolo
        End Get
        Set(value As String)
            _Protocolo = value
        End Set
    End Property
    Public Sub New()

    End Sub

    Public Sub New(ByVal cStat As String, ByVal xMotivo As String, ByVal dhRecibo As String, ByVal nRecibo As String, ByVal Protocolo As String)

        Me._cStat = cStat
        Me._xMotivo = xMotivo
        Me._dhRecibo = dhRecibo
        Me._nRecibo = nRecibo
        Me._Protocolo = Protocolo

    End Sub

    Public Function AutorizandoNota(ByVal Caminho As String, ByVal NumLote As String, ByVal NomeCert As String, ByVal Tipomsg As Boolean
                                    ) As Boolean
        Dim doc As New XmlDocument()
        Dim bConsulta As Boolean
        Dim Resposta As XmlNode
        Dim NomeArq As String

        'Tipomsg se verdadeira colocar indSinc = 1 sicrona
        ''Tipomsg se false colocar indSinc = 0 assicrona

        Try
            bConsulta = True
            NomeArq = sGerados & "LoteNFeV4_" & NumLote.ToString() & DateTime.Now.ToString("yyyyMM") & ".xml"
            Dim EscrevCons As New XmlTextWriter(NomeArq, Encoding.GetEncoding("UTF-8"))

            EscrevCons.WriteStartDocument()
            EscrevCons.Formatting = Formatting.Indented
            EscrevCons.WriteStartElement("enviNFe")
            EscrevCons.WriteAttributeString("xmlns", "http://www.portalfiscal.inf.br/nfe")
            EscrevCons.WriteAttributeString("versao", VersaoNota)
            EscrevCons.WriteElementString("idLote", Microsoft.VisualBasic.Strings.Trim(NumLote))

            If Tipomsg = True Then
                EscrevCons.WriteElementString("indSinc", Microsoft.VisualBasic.Strings.Trim("1"))
            Else
                EscrevCons.WriteElementString("indSinc", Microsoft.VisualBasic.Strings.Trim("0"))
            End If

            EscrevCons.WriteEndElement() 'fim nó Consulta
            EscrevCons.WriteEndDocument()
            EscrevCons.Close()

            Dim sr2 As StreamReader
            Dim sw2 As StreamWriter

            Dim _stringXml As String

            Dim DocLote As New XmlDocument()
            Dim DocNota As New XmlDocument()
            Dim elementsByTagName As XmlNodeList
            Dim NodeNfe As XmlNodeList

            sr2 = File.OpenText(NomeArq)
            _stringXml = sr2.ReadToEnd()
            sr2.Close()


            DocLote.LoadXml(_stringXml)

            Dim DirDiretorio As DirectoryInfo = New DirectoryInfo(Caminho)
            Dim y() As FileInfo
            Dim x As FileInfo
            Dim sRpsNotas As String


            y = DirDiretorio.GetFiles("*.xml", SearchOption.AllDirectories)

            For Each x In y

                sRpsNotas = x.FullName
                DocNota.Load(sRpsNotas)
                DocNota.PreserveWhitespace = True
                DocNota.PreserveWhitespace = True
                elementsByTagName = DocLote.GetElementsByTagName("enviNFe")
                NodeNfe = DocNota.GetElementsByTagName("NFe")
                elementsByTagName.Item(0).AppendChild(DocLote.ImportNode(NodeNfe.Item(0), True))

            Next

            Dim sNomeArq As String = sGerados & +"LoteNFeV4_" & NumLote.ToString() & DateTime.Now.ToString("yyyyMM") + ".xml"


            'salva o arquivo
            sw2 = File.CreateText(sNomeArq)
            sw2.Write(DocLote.OuterXml)
            sw2.Close()

            sr2 = File.OpenText(NomeArq)
            _stringXml = sr2.ReadToEnd()
            sr2.Close()

            DocLote.LoadXml(_stringXml)

            Dim _X509Cert As New X509Certificate()

            _X509Cert = XmlNota.BuscarNome(NomeCert)

            Dim NewResposta As String = ""


            If TipoEnvio = 55 Then
                Dim Autoriza As AutorizacaoLote.NFeAutorizacao4
                Autoriza = New AutorizacaoLote.NFeAutorizacao4()
                Autoriza.Url = "https://nfe.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                Autoriza.SoapVersion = System.Web.Services.Protocols.SoapProtocolVersion.Soap12
                Autoriza.ClientCertificates.Add(_X509Cert)
                Resposta = Autoriza.nfeAutorizacaoLote(DocLote)

                NewResposta = Convert.ToString(Resposta.OuterXml)

                Autoriza.Dispose()
            Else
                Dim AutorizaNFCe As AutorizacaoLote.NFeAutorizacao4
                AutorizaNFCe = New AutorizacaoLote.NFeAutorizacao4()
                AutorizaNFCe.Url = "https://nfce.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                AutorizaNFCe.SoapVersion = System.Web.Services.Protocols.SoapProtocolVersion.Soap12
                AutorizaNFCe.ClientCertificates.Add(_X509Cert)
                Resposta = AutorizaNFCe.nfeAutorizacaoLote(DocLote)

                NewResposta = Convert.ToString(Resposta.OuterXml)

                AutorizaNFCe.Dispose()
            End If

            Dim NomeArqRet As String = sRetonos & "Ret_AutorizaLote_" + NumLote + DateTime.Now.ToString("yyyyMM") + ".xml"

            Dim sDados As String = ""
            Dim sr As StreamReader

            'salva o arquivo na pasta de retorno e transforma em Xml
            sw2 = File.CreateText(NomeArqRet)
            sw2.Write(NewResposta)
            sw2.Close()

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
                    ElseIf (cStat = "104" Or cStat = "103") And Not IsDBNull(xMotivo) And Not IsDBNull(dhRecibo) Then
                        For Each xy In sbno.ChildNodes
                            If cStat = "103" Then
                                If xy.Name = "nRec" Then nRecibo = xy.InnerText
                            Else
                                For Each z In xy.ChildNodes
                                    If xy.Name = "verAplic" Then
                                        VerAplic = z.InnerText
                                    ElseIf xy.Name = "chNFe" Then
                                        chNfe = z.InnerText
                                    ElseIf xy.Name = "digVal" Then
                                        digVal = z.InnerText
                                    ElseIf xy.Name = "cStat" Then
                                        cStatNfe = z.InnerText
                                    ElseIf xy.Name = "xMotivo" Then
                                        xMotivoNfe = z.InnerText
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
            Next

            If Tipomsg Then
                If cStat = "104" And cStatNfe = "100" Then

                    AtualizaLote(cStat, xMotivo, dhRecibo, nRecibo, NumLote)

                    If chNfe <> "" And dhRecibo <> "" And digVal <> "" And cStatNfe <> "" And xMotivoNfe <> "" And VerAplic <> "" Then
                        AtualizaLoteNota(chNfe, dhRecibo, "0", digVal, cStatNfe, xMotivoNfe, VerAplic, NumLote, True)
                        XmlDistribuicao(chNfe, sNota & chNfe & "NFe.xml", sEnviados & chNfe & "NFe.xml")
                    End If
                Else
                    bConsulta = False
                    AtualizaLote(cStat, xMotivo, dhRecibo, nRecibo, NumLote)
                End If
            Else
                If cStat = "103" Then
                    AtualizaLote(cStat, xMotivo, dhRecibo, nRecibo, NumLote)
                Else
                    bConsulta = False
                    AtualizaLote(cStat, xMotivo, dhRecibo, nRecibo, NumLote)
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            bConsulta = False
        End Try

        Return bConsulta

    End Function





End Class

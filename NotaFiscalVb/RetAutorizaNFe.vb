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

Public Class RetAutorizaNFe

    Private _sRec As String
    Private _cStat As String
    Private _xMotivo As String
    Private _dhRecbto As String
    Private _cMsg As String
    Private _xMsg As String
    Private _protNfe As String
    Private _VerAplic As String

    Private _chNFe As String
    Private _dhRecbtoNf As String
    Private _nProt As String
    Private _digVal As String
    Private _cStatNf As String
    Private _xMot As String
    Private _Signature As String

    Private XmlNota As New AssinarXml()


    Public Property sRec() As String
        Get
            Return _sRec
        End Get
        Set(value As String)
            _sRec = value
        End Set
    End Property
    Public Property cStat() As String
        Get
            Return _cStat
        End Get
        Set(value As String)
            _cStat = value
        End Set
    End Property
    Public Property xMotivo() As String
        Get
            Return _xMotivo
        End Get
        Set(value As String)
            _xMotivo = value
        End Set
    End Property
    Public Property dhRecbto() As String
        Get
            Return _dhRecbto
        End Get
        Set(value As String)
            _dhRecbto = value
        End Set
    End Property
    Public Property cMsg() As String
        Get
            Return _cMsg
        End Get
        Set(value As String)
            _cMsg = value
        End Set
    End Property
    Public Property xMsg() As String
        Get
            Return _xMsg
        End Get
        Set(value As String)
            _xMsg = value
        End Set
    End Property
    Public Property protNfe() As String
        Get
            Return _protNfe
        End Get
        Set(value As String)
            _protNfe = value
        End Set
    End Property
    Public Property VerAplic() As String
        Get
            Return _VerAplic
        End Get
        Set(value As String)
            _VerAplic = value
        End Set
    End Property


    Public Property chNFe() As String
        Get
            Return _chNFe
        End Get
        Set(value As String)
            _chNFe = value
        End Set
    End Property
    Public Property dhRecbtoNf() As String
        Get
            Return _dhRecbtoNf
        End Get
        Set(value As String)
            _dhRecbtoNf = value
        End Set
    End Property
    Public Property nProt() As String
        Get
            Return _nProt
        End Get
        Set(value As String)
            _nProt = value
        End Set
    End Property
    Public Property digVal() As String
        Get
            Return _digVal
        End Get
        Set(value As String)
            _digVal = value
        End Set
    End Property
    Public Property cStatNf() As String
        Get
            Return _cStatNf
        End Get
        Set(value As String)
            _cStatNf = value
        End Set
    End Property
    Public Property xMot() As String
        Get
            Return _xMot
        End Get
        Set(value As String)
            _xMot = value
        End Set
    End Property
    Public Property Signature() As String
        Get
            Return _Signature
        End Get
        Set(value As String)
            _Signature = value
        End Set
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal sRec As String, ByVal cStat As String, ByVal xMotivo As String, ByVal dhRecbto As String, ByVal cMsg As String, _
                  ByVal protNfe As String, ByVal VerAplic As String, ByVal chNFe As String, ByVal dhRecbtoNf As String, ByVal nProt As String, _
                  ByVal digVal As String, ByVal cStatNf As String, ByVal xMot As String, ByVal Signature As String)



        Me._sRec = sRec
        Me._cStat = cStat
        Me._xMotivo = xMotivo
        Me._dhRecbto = dhRecbto
        Me._cMsg = cMsg
        Me._xMsg = xMsg
        Me._protNfe = protNfe
        Me._VerAplic = VerAplic
        Me._chNFe = chNFe
        Me._dhRecbtoNf = dhRecbtoNf
        Me._nProt = nProt
        Me._digVal = digVal
        Me._cStatNf = cStatNf
        Me._xMot = xMot
        Me._Signature = Signature


    End Sub

    Public Function RetornandoNFe(ByVal NumRecibo As String, ByVal NomeCert As String, ByVal NumLote As Integer, ByRef cStatus As String) As Boolean
        Dim Retorno As Boolean
        Dim doc As New XmlDocument
        Dim Resposta As XmlNode
        Dim NomeArq As String

        Try
            Retorno = True

            NomeArq = sRequisicoes + "Ret_Notas_Lote_" + NumLote + "_" + DateTime.Now.ToString("yyyyMM") + ".xml"
            Dim EscLote As New XmlTextWriter(NomeArq, Encoding.GetEncoding("UTF-8"))

            EscLote.WriteStartDocument()
            EscLote.Formatting = Formatting.Indented
            EscLote.WriteStartElement("consReciNFe")
            EscLote.WriteAttributeString("xmlns", "http://www.portalfiscal.inf.br/nfe")
            EscLote.WriteAttributeString("versao", VersaoNota.Trim())
            EscLote.WriteElementString("tpAmb", TpAmb.ToString().Trim())
            EscLote.WriteElementString("nRec", NumRecibo.Trim())
            EscLote.WriteEndElement() '//fim nó Consulta
            EscLote.WriteEndDocument() '
            EscLote.Close()


            Dim sr As StreamReader
            Dim sDados As String

            sr = File.OpenText(NomeArq)
            sDados = sr.ReadToEnd()
            sr.Close()

            doc.PreserveWhitespace = False

            doc.LoadXml(sDados)

            Dim _X509Cert As New X509Certificate()
            _X509Cert = XmlNota.BuscarNome(NomeCert)


            Dim res As String = ""

            If TipoEnvio = 55 Then


                Dim Retorna As New RetAutorizacaoLote.NFeRetAutorizacao4()
                Retorna.Url = "https://nfe.svrs.rs.gov.br/ws/NfeRetAutorizacao/NFeRetAutorizacao4.asmx"
                Retorna.SoapVersion = System.Web.Services.Protocols.SoapProtocolVersion.Soap12
                Retorna.ClientCertificates.Add(_X509Cert)
                Resposta = Retorna.nfeRetAutorizacaoLote(doc)
                res = Convert.ToString(Resposta.OuterXml)

                Retorna.Dispose()
            Else
                Dim RetornaNFCe As New RetAutorizacaoLoteNFCe.NFeRetAutorizacao4()
                RetornaNFCe.Url = "https://nfce.svrs.rs.gov.br/ws/NfeRetAutorizacao/NFeRetAutorizacao4.asmx"
                RetornaNFCe.SoapVersion = System.Web.Services.Protocols.SoapProtocolVersion.Soap12
                RetornaNFCe.ClientCertificates.Add(_X509Cert)
                Resposta = RetornaNFCe.nfeRetAutorizacaoLote(doc)
                res = Convert.ToString(Resposta.OuterXml)

                RetornaNFCe.Dispose()

            End If

            Dim sw2 As StreamWriter

            Dim NomeArqRet As String = sRetonos & "Ret_RetAutoriza_" & NumLote & "_" + DateTime.Now.ToString("yyyyMM") + ".xml"


            'salva o arquivo na pasta de retorno e transforma em Xml
            sw2 = File.CreateText(NomeArqRet)
            sw2.Write(res)
            sw2.Close() '

            Dim doc2 As New XmlDocument()

            sr = File.OpenText(NomeArqRet)
            sDados = sr.ReadToEnd()
            sr.Close()

            doc2.PreserveWhitespace = False

            doc2.LoadXml(sDados)

            Dim No, nNo, g, N As XmlNode


            For Each No In doc2.ChildNodes
                For Each nNo In No.ChildNodes
                    If nNo.Name = "cStat" Then
                        cStat = nNo.InnerText
                        cStatus = nNo.InnerText
                    ElseIf nNo.Name = "xMotivo" Then
                        xMotivo = nNo.InnerText
                    ElseIf nNo.Name = "dhRecbto" Then
                        dhRecbto = Convert.ToDateTime(nNo.InnerText).ToString("yyyy-MM-dd hh:mm:ss")
                    ElseIf nNo.Name = "verAplic" Then
                        VerAplic = nNo.InnerText
                    End If

                    If cStat = "104" And xMotivo <> vbNull And dhRecbto <> vbNull Then
                        AtualizaLote(cStat, xMotivo, dhRecbto, NumRecibo, NumLote)
                        For Each g In nNo.ChildNodes
                            For Each N In g.ChildNodes
                                If N.Name = "chNFe" Then
                                    chNFe = N.InnerText
                                ElseIf N.Name = "dhRecbto" Then
                                    dhRecbtoNf = Convert.ToDateTime(N.InnerText).ToString("yyyy-MM-dd hh:mm:ss")
                                ElseIf N.Name = "nProt" Then
                                    nProt = N.InnerText
                                ElseIf N.Name = "digVal" Then
                                    digVal = N.InnerText
                                ElseIf N.Name = "cStat" Then
                                    cStatNf = N.InnerText
                                ElseIf N.Name = "xMotivo" Then
                                    xMot = N.InnerText
                                End If
                                If chNFe <> "" And dhRecbtoNf <> "" And nProt <> "" And digVal <> "" And cStatNf <> "" And xMot <> "" And VerAplic <> "" Then
                                    AtualizaLoteNota(chNFe, dhRecbtoNf, nProt, digVal, cStatNf, xMot, VerAplic, NumLote, True)
                                    XmlDistribuicao(chNFe, sNota & chNFe & "NFe.xml", sEnviados & chNFe & "NFe.xml")
                                End If
                            Next
                        Next
                    End If
                Next
            Next
        Catch ex As Exception
            Retorno = False
        End Try
        Return Retorno

    End Function

    






End Class

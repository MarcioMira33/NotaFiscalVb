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
Public Class ClsCancelamentoNFe
    'Private _id As String
    'Private _cOrgao As String

    'Private _CpfCnpj As String
    'Private _chNFe As String
    'Private _dhEvento As String
    'Private _tpEvento As String
    'Private _nSeqEvento As String
    'Private _Versao As String
    'Private _descEnvento As String
    'Private _nProt As String
    'Private _xJust As String
    'Private _detEvento As String

    ''Dados do Lote'
    'Private _idLote As String
    'Private _VerAplic As String
    'Private _cStat As String
    'Private _xMotivo As String

    ''Dados da Nota
    'Private _cStatNota As String
    'Private _xMotivoNota As String
    'Private _xEvento As String
    'Private _dhRegEvento As String
    'Private _nProtocolo As String

    'Public Property cStatNota As String
    '    Get
    '        Return _cStatNota
    '    End Get
    '    Set(value As String)
    '        _cStatNota = value
    '    End Set
    'End Property
    'Public Property xMotivoNota As String
    '    Get
    '        Return _xMotivoNota
    '    End Get
    '    Set(value As String)
    '        _xMotivoNota = value
    '    End Set
    'End Property
    'Public Property xEvento As String
    '    Get
    '        Return _xEvento
    '    End Get
    '    Set(value As String)
    '        _xEvento = value
    '    End Set
    'End Property
    'Public Property dhRegEvento As String
    '    Get
    '        Return _dhRegEvento
    '    End Get
    '    Set(value As String)
    '        _dhRegEvento = value
    '    End Set
    'End Property
    'Public Property nProtocolo As String
    '    Get
    '        Return _nProtocolo
    '    End Get
    '    Set(value As String)
    '        _nProtocolo = value
    '    End Set
    'End Property

    'Public Property idLote As String
    '    Get
    '        Return _idLote
    '    End Get
    '    Set(value As String)
    '        _idLote = value
    '    End Set
    'End Property
    'Public Property VerAplic As String
    '    Get
    '        Return _VerAplic
    '    End Get
    '    Set(value As String)
    '        _VerAplic = value
    '    End Set
    'End Property
    'Public Property cStat As String
    '    Get
    '        Return _cStat
    '    End Get
    '    Set(value As String)
    '        _cStat = value
    '    End Set
    'End Property
    'Public Property xMotivo As String
    '    Get
    '        Return _xMotivo
    '    End Get
    '    Set(value As String)
    '        _xMotivo = value
    '    End Set
    'End Property


    'Public Property detEvento As String
    '    Get
    '        Return _detEvento
    '    End Get
    '    Set(value As String)
    '        _detEvento = value
    '    End Set
    'End Property

    'Public Property id As String
    '    Get
    '        Return _id
    '    End Get
    '    Set(value As String)
    '        _id = value
    '    End Set
    'End Property
    'Public Property cOrgao As String
    '    Get
    '        Return _cOrgao
    '    End Get
    '    Set(value As String)
    '        _cOrgao = value
    '    End Set
    'End Property

    'Public Property CpfCnpj As String
    '    Get
    '        Return _CpfCnpj
    '    End Get
    '    Set(value As String)
    '        _CpfCnpj = value
    '    End Set
    'End Property
    'Public Property chNFe As String
    '    Get
    '        Return _chNFe
    '    End Get
    '    Set(value As String)
    '        _chNFe = value
    '    End Set
    'End Property
    'Public Property dhEvento As String
    '    Get
    '        Return _dhEvento
    '    End Get
    '    Set(value As String)
    '        _dhEvento = value
    '    End Set
    'End Property
    'Public Property tpEvento As String
    '    Get
    '        Return _tpEvento
    '    End Get
    '    Set(value As String)
    '        _tpEvento = value
    '    End Set
    'End Property
    'Public Property nSeqEvento As String
    '    Get
    '        Return _nSeqEvento
    '    End Get
    '    Set(value As String)
    '        _nSeqEvento = value
    '    End Set
    'End Property
    'Public Property Versao As String
    '    Get
    '        Return _Versao
    '    End Get
    '    Set(value As String)
    '        _Versao = value
    '    End Set
    'End Property
    'Public Property descEnvento As String
    '    Get
    '        Return _descEnvento
    '    End Get
    '    Set(value As String)
    '        _descEnvento = value
    '    End Set
    'End Property
    'Public Property nProt As String
    '    Get
    '        Return _nProt
    '    End Get
    '    Set(value As String)
    '        _nProt = value
    '    End Set
    'End Property
    'Public Property xJust As String
    '    Get
    '        Return _xJust
    '    End Get
    '    Set(value As String)
    '        _xJust = value
    '    End Set
    'End Property

    'Public Sub New(ByVal id As String, ByVal cOrgao As String, ByVal CpfCnpj As String, ByVal chNFe As String, ByVal dhEvento As String, ByVal tpEvento As String, _
    '               ByVal nSeqEvento As String, ByVal detEvento As String, ByVal Versao As String, ByVal descEnvento As String, ByVal nProt As String, ByVal xJust As String _
    '               , ByVal cStatNota As String, ByVal xMotivoNota As String, ByVal xEvento As String, ByVal dhRegEvento As String, ByVal nProtocolo As String, _
    '               ByVal idLote As String, ByVal VerAplic As String, ByVal cStat As String, ByVal xMotivo As String)

    '    Me._id = id
    '    Me._cOrgao = cOrgao
    '    Me._CpfCnpj = CpfCnpj
    '    Me._chNFe = chNFe
    '    Me._dhEvento = dhEvento
    '    Me._tpEvento = tpEvento
    '    Me._nSeqEvento = nSeqEvento
    '    Me._Versao = Versao
    '    Me._descEnvento = descEnvento
    '    Me._nProt = nProt
    '    Me._xJust = xJust
    '    Me._detEvento = detEvento

    '    Me._idLote = idLote
    '    Me._VerAplic = VerAplic
    '    Me._cStat = cStat
    '    Me._xMotivo = xMotivo

    '    Me._cStatNota = cStatNota
    '    Me._xMotivoNota = xMotivoNota
    '    Me._xEvento = xEvento
    '    Me._dhRegEvento = dhRegEvento
    '    Me._nProtocolo = nProtocolo


    'End Sub

    'Public Sub New()

    'End Sub

    'Public Function CancelarNota(NumLote As String, TipoPessoa As String) As Boolean
    '    Dim NomeArq As String
    '    Dim _X509Cert As New X509Certificate
    '    Try
    '        NomeArq = GCancelamento & chNFe & "-ped-can.xml"
    '        Dim wrt As New XmlTextWriter(NomeArq, System.Text.Encoding.GetEncoding("UTF-8"))

    '        With wrt
    '            .WriteStartDocument()
    '            .WriteStartElement("envEvento")
    '            .WriteAttributeString("versao", "1.00")
    '            .WriteAttributeString("xmlns", "http://www.portalfiscal.inf.br/nfe")

    '            .WriteElementString("idLote", Microsoft.VisualBasic.Strings.Trim(NumLote))
    '            .WriteStartElement("evento")
    '            .WriteAttributeString("versao", "1.00")
    '            .WriteAttributeString("xmlns", "http://www.portalfiscal.inf.br/nfe")

    '            .WriteStartElement("infEvento")
    '            .WriteAttributeString("Id", "ID110111" & chNFe.Trim & "01")
    '            .WriteElementString("cOrgao", "21")
    '            .WriteElementString("tpAmb", GTpAmb)

    '            ' 'TIPO DE PESSO
    '            'If TipoPessoa = "F" Then
    '            '.WriteElementString("CPF", CpfCnpj)
    '            'Else
    '            .WriteElementString("CNPJ", GCnpjEmp)
    '            ' End If

    '            .WriteElementString("chNFe", chNFe)
    '            .WriteElementString("dhEvento", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") & "-03:00")
    '            .WriteElementString("tpEvento", "110111")
    '            .WriteElementString("nSeqEvento", "1")
    '            .WriteElementString("verEvento", "1.00")
    '            .WriteStartElement("detEvento")
    '            .WriteAttributeString("versao", "1.00")
    '            .WriteElementString("descEvento", "Cancelamento")
    '            .WriteElementString("nProt", nProt)
    '            .WriteElementString("xJust", "Valor indevido do produto")
    '            .WriteEndElement()
    '            .WriteEndElement()
    '            .WriteEndElement()
    '            .Close()
    '        End With

    '        Dim clsAssin As New ClsAssinarXml()

    '        If GNomeCertificado = "" Then
    '            _X509Cert = clsAssin.BuscarNome("")
    '            GNomeCertificado = _X509Cert.Subject
    '        Else
    '            _X509Cert = clsAssin.BuscarNome(GNomeCertificado)
    '        End If

    '        Dim sr As StreamReader
    '        Dim sw As StreamWriter
    '        Dim nfeDadosMsg As String
    '        'Dim XMLStringAssinado As String

    '        sr = File.OpenText(NomeArq)
    '        nfeDadosMsg = sr.ReadToEnd()
    '        sr.Close()

    '        XMLStringAssinado = ""

    '        clsAssin.Assinar(nfeDadosMsg, "infEvento", _X509Cert)

    '        sw = File.CreateText(NomeArq)
    '        sw.Write(XMLStringAssinado)
    '        sw.Close()

    '        sr = File.OpenText(NomeArq)
    '        nfeDadosMsg = sr.ReadToEnd()
    '        sr.Close()

    '        Dim doc As New XmlDocument()

    '        doc.LoadXml(nfeDadosMsg)

    '        Dim nRespostas As String = ""
    '        Dim Resposta As XmlNode

    '        'Alterado em 02/09/2020**************************************************************************************
    '        '
    '        'Dim Canc As New RecepcaoEvento.NFeRecepcaoEvento4()
    '        'Canc.Url = "https://www.sefazvirtual.fazenda.gov.br/NFeRecepcaoEvento4/NFeRecepcaoEvento4.asmx"
    '        'Canc.SoapVersion = System.Web.Services.Protocols.SoapProtocolVersion.Soap12
    '        'Canc.ClientCertificates.Add(_X509Cert)
    '        'Resposta = Canc.nfeRecepcaoEvento(doc)
    '        'nRespostas = Convert.ToString(Resposta.OuterXml)
    '        'Canc.Dispose()


    '        If GTipoEnvio = 55 Then 'NFe
    '            Dim Canc As New RecepcaoEvento.NFeRecepcaoEvento4()
    '            Canc.Url = "https://www.sefazvirtual.fazenda.gov.br/NFeRecepcaoEvento4/NFeRecepcaoEvento4.asmx"
    '            Canc.SoapVersion = System.Web.Services.Protocols.SoapProtocolVersion.Soap12
    '            Canc.ClientCertificates.Add(_X509Cert)
    '            Resposta = Canc.nfeRecepcaoEvento(doc)
    '            nRespostas = Convert.ToString(Resposta.OuterXml)
    '            Canc.Dispose()

    '        Else 'NFCe

    '            Dim CancNfce As New RecepcaoEventoNFCe.NFeRecepcaoEvento4()
    '            CancNfce.Url = "https://nfce.svrs.rs.gov.br/ws/recepcaoevento/recepcaoevento4.asmx"
    '            CancNfce.SoapVersion = System.Web.Services.Protocols.SoapProtocolVersion.Soap12
    '            CancNfce.ClientCertificates.Add(_X509Cert)
    '            Resposta = CancNfce.nfeRecepcaoEvento(doc)
    '            nRespostas = Convert.ToString(Resposta.OuterXml)
    '            CancNfce.Dispose()
    '        End If
    '        'Fim*********************************************************************************************


    '        Dim NomeArqRet As String = GCancelamento & "Ret_Cance" & "Id" & chNFe + ".xml"

    '        Dim sDados As String
    '        'Dim sr As StreamReader

    '        sw = File.CreateText(NomeArqRet)
    '        sw.Write(nRespostas)
    '        sw.Close()


    '        Dim doc2 As New XmlDocument()

    '        sr = File.OpenText(NomeArqRet)
    '        sDados = sr.ReadToEnd()
    '        sr.Close()
    '        doc2.PreserveWhitespace = True
    '        doc2.LoadXml(sDados)

    '        Dim No, No1, No2 As XmlNode

    '        For Each No In doc2.ChildNodes
    '            If No.Name = "idLote" Then
    '                idLote = No.InnerText
    '            ElseIf No.Name = "verAplic" Then
    '                VerAplic = No.InnerText
    '            ElseIf No.Name = "cStat" Then
    '                cStat = No.InnerText
    '            ElseIf No.Name = "xMotivo" Then
    '                xMotivo = No.InnerText
    '            End If
    '            If cStat = "135" Then
    '                For Each No1 In No.ChildNodes
    '                    For Each No2 In No1.ChildNodes
    '                        If No2.Name = "cStat" Then
    '                            cStatNota = No2.InnerText
    '                        ElseIf No2.Name = "xMotivo" Then
    '                            xMotivoNota = No2.Name
    '                        ElseIf No2.Name = "xEvento" Then
    '                            xEvento = No2.InnerText
    '                        ElseIf No2.Name = "dhRegEvento" Then
    '                            dhRegEvento = No2.InnerText
    '                        ElseIf No2.Name = "nProt" Then
    '                            nProtocolo = No2.InnerText
    '                        End If
    '                    Next
    '                Next
    '            End If
    '        Next


    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '        Return False
    '    End Try

    '    Return True
    'End Function




End Class

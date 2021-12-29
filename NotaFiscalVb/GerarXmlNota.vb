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
Public Class GerarXmlNota

    Private ChaveNota As String
    Private NomeArq As String

    'Identificação da Nota fiscal eletrônica
    Private _cUF As String
    Private _cNF As String
    Private _natOp As String
    Private _indPag As String
    Private _smod As String
    Private _serie As String
    Private _nNF As String
    Private _dhEmi As String
    Private _dhSaiEnt As String
    Private _tpNF As String
    Private _idDest As String
    Private _cMunFG As String
    Private _tpImp As String
    Private _tpRmis As String
    Private _cDV As String
    Private _tpAmb As String
    Private _finNFe As String
    Private _finFinal As String
    Private _indPres As String
    Private _procEmi As String
    Private _verPros As String
    Private _dhCont As String
    Private _xJust As String

    'Identicação do emitente da nota
    Private _CNPJ As String
    Private _CPF As String
    Private _xNome As String
    Private _xFant As String
    Private _xLgr As String
    Private _nro As String
    Private _xCpl As String
    Private _xBairro As String
    Private _cMun As String
    Private _xMun As String
    Private _UF As String
    Private _CEP As String
    Private _cPais As String
    Private _xPais As String
    Private _fone As String
    Private _IE As String
    Private _IEST As String
    Private _IM As String
    Private _CNAE As String
    Private _CRT As String

    Private _CNPJDest As String
    Private _CPFDest As String
    Private _idEstrangeiro As String
    Private _xNomeDest As String
    Private _xLgrDest As String
    Private _nroDest As String
    Private _xCplDest As String
    Private _xBairroDest As String
    Private _cMunDest As String
    Private _xMunDest As String
    Private _UFDest As String
    Private _CEPDest As String
    Private _cPaisDest As String
    Private _xPaisDest As String
    Private _foneDest As String
    Private _indIEdest As String
    Private _IEDest As String
    Private _ISUFDest As String
    Private _IMDest As String
    Private _email As String



    'Produts e Serviços
    Private _nItem As String
    Private _cProd As String
    Private _cEAN As String
    Private _xProd As String
    Private _NCM As String
    Private _CFOP As String
    Private _uCom As String
    Private _qCom As String
    Private _vUnCom As String
    Private _vProd As String
    Private _cEANTrib As String
    Private _uTrib As String
    Private _qTrib As String
    Private _vUnTrib As String
    Private _vFrete As String
    'Private _vSeg As String
    'Private _vDesc As String
    'Private _vOutro As String
    Private _indTot As String

    'Impostos
    'Private _vBC As String
    'Private _vAliq As String
    'Private _vISSQN As String
    'Private _cMunFGIss As String
    'Private _cListServ As String
    'Private _vDeducao As String
    'Private _vOutroIss As String
    'Private _vDescIncond As String
    'Private _vDescCond As String
    'Private _vISSRet As String
    'Private _indISS As String
    'Private _cServico As String
    'Private _cMunIss As String
    'Private _cPaisIss As String
    'Private _nProcesso As String
    'Private _indIncentivo As String

    Private _vICMS As String

    Private wr As XmlTextWriter

    Public Property vICMS() As String
        Get
            Return _vICMS
        End Get
        Set(value As String)
            _vICMS = value
        End Set
    End Property

#Region "Identificação da Nota"



    Public Property cUF() As String
        Get
            Return _cUF
        End Get
        Set(value As String)
            _cUF = value
        End Set
    End Property
    Public Property cNF() As String
        Get
            Return _cNF
        End Get
        Set(value As String)
            _cNF = value
        End Set
    End Property
    Public Property natOp() As String
        Get
            Return _natOp
        End Get
        Set(value As String)
            _natOp = value
        End Set
    End Property
    Public Property indPag() As String
        Get
            Return _indPag
        End Get
        Set(value As String)
            _indPag = value
        End Set
    End Property
    Public Property smod() As String
        Get
            Return _smod
        End Get
        Set(value As String)
            _smod = value
        End Set
    End Property
    Public Property serie() As String
        Get
            Return _serie
        End Get
        Set(value As String)
            _serie = value
        End Set
    End Property
    Public Property nNF() As String
        Get
            Return _nNF
        End Get
        Set(value As String)
            _nNF = value
        End Set
    End Property
    Public Property dhEmi() As String
        Get
            Return _dhEmi
        End Get
        Set(value As String)
            _dhEmi = value
        End Set
    End Property
    Public Property dhSaiEnt() As String
        Get
            Return _dhSaiEnt
        End Get
        Set(value As String)
            _dhSaiEnt = value
        End Set
    End Property
    Public Property tpNF() As String
        Get
            Return _tpNF
        End Get
        Set(value As String)
            _tpNF = value
        End Set
    End Property
    Public Property idDest() As String
        Get
            Return _idDest
        End Get
        Set(value As String)
            _idDest = value
        End Set
    End Property
    Public Property cMunFG() As String
        Get
            Return _cMunFG
        End Get
        Set(value As String)
            _cMunFG = value
        End Set
    End Property
    Public Property tpImp() As String
        Get
            Return _tpImp
        End Get
        Set(value As String)
            _tpImp = value
        End Set
    End Property
    Public Property tpRmis() As String
        Get
            Return _tpRmis
        End Get
        Set(value As String)
            _tpRmis = value
        End Set
    End Property
    Public Property cDV() As String
        Get
            Return _cDV
        End Get
        Set(value As String)
            _cDV = value
        End Set
    End Property
    Public Property tpAmb() As String
        Get
            Return _tpAmb
        End Get
        Set(value As String)
            _tpAmb = value
        End Set
    End Property
    Public Property finNFe() As String
        Get
            Return _finNFe
        End Get
        Set(value As String)
            _finNFe = value
        End Set
    End Property
    Public Property finFinal() As String
        Get
            Return _finFinal
        End Get
        Set(value As String)
            _finFinal = value
        End Set
    End Property
    Public Property indPres() As String
        Get
            Return _indPres
        End Get
        Set(value As String)
            _indPres = value
        End Set
    End Property
    Public Property procEmi() As String
        Get
            Return _procEmi
        End Get
        Set(value As String)
            _procEmi = value
        End Set
    End Property
    Public Property verPros() As String
        Get
            Return _verPros
        End Get
        Set(value As String)
            _verPros = value
        End Set
    End Property
    Public Property dhCont() As String
        Get
            Return _dhCont
        End Get
        Set(value As String)
            _dhCont = value
        End Set
    End Property
    Public Property xJust() As String
        Get
            Return _xJust
        End Get
        Set(value As String)
            _xJust = value
        End Set
    End Property
#End Region

#Region "Identificação do Emitente"
    Public Property CNPJ() As String
        Get
            Return _CNPJ
        End Get
        Set(value As String)
            _CNPJ = value
        End Set
    End Property
    Public Property CPF() As String
        Get
            Return _CPF
        End Get
        Set(value As String)
            _CPF = value
        End Set
    End Property
    Public Property xNome() As String
        Get
            Return _xNome
        End Get
        Set(value As String)
            _xNome = value
        End Set
    End Property
    Public Property xFant() As String
        Get
            Return _xFant
        End Get
        Set(value As String)
            _xFant = value
        End Set
    End Property
    Public Property xLgr() As String
        Get
            Return _xLgr
        End Get
        Set(value As String)
            _xLgr = value
        End Set
    End Property
    Public Property nro() As String
        Get
            Return _nro
        End Get
        Set(value As String)
            _nro = value
        End Set
    End Property
    Public Property xCpl() As String
        Get
            Return _xCpl
        End Get
        Set(value As String)
            _xCpl = value
        End Set
    End Property
    Public Property xBairro() As String
        Get
            Return _xBairro
        End Get
        Set(value As String)
            _xBairro = value
        End Set
    End Property
    Public Property cMun() As String
        Get
            Return _cMun
        End Get
        Set(value As String)
            _cMun = value
        End Set
    End Property
    Public Property xMun() As String
        Get
            Return _xMun
        End Get
        Set(value As String)
            _xMun = value
        End Set
    End Property
    Public Property UF() As String
        Get
            Return _UF
        End Get
        Set(value As String)
            _UF = value
        End Set
    End Property
    Public Property CEP() As String
        Get
            Return _CEP
        End Get
        Set(value As String)
            _CEP = value
        End Set
    End Property
    Public Property cPais() As String
        Get
            Return _cPais
        End Get
        Set(value As String)
            _cPais = value
        End Set
    End Property
    Public Property xPais() As String
        Get
            Return _xPais
        End Get
        Set(value As String)
            _xPais = value
        End Set
    End Property
    Public Property fone() As String
        Get
            Return _fone
        End Get
        Set(value As String)
            _fone = value
        End Set
    End Property
    Public Property IE() As String
        Get
            Return _IE
        End Get
        Set(value As String)
            _IE = value
        End Set
    End Property
    Public Property IEST() As String
        Get
            Return _IEST
        End Get
        Set(value As String)
            _IEST = value
        End Set
    End Property
    Public Property IM() As String
        Get
            Return _IM
        End Get
        Set(value As String)
            _IM = value
        End Set
    End Property
    Public Property CNAE() As String
        Get
            Return _CNAE
        End Get
        Set(value As String)
            _CNAE = value
        End Set
    End Property
    Public Property CRT() As String
        Get
            Return _CRT
        End Get
        Set(value As String)
            _CRT = value
        End Set
    End Property
#End Region

#Region "Identificação do Destinatario"
    Public Property CNPJDest() As String
        Get
            Return _CNPJDest
        End Get
        Set(value As String)
            _CNPJDest = value
        End Set
    End Property
    Public Property CPFDest() As String
        Get
            Return _CPFDest
        End Get
        Set(value As String)
            _CPFDest = value
        End Set
    End Property
    Public Property idEstrangeiro() As String
        Get
            Return _idEstrangeiro
        End Get
        Set(value As String)
            _idEstrangeiro = value
        End Set
    End Property
    Public Property xNomeDest() As String
        Get
            Return _xNomeDest
        End Get
        Set(value As String)
            _xNomeDest = value
        End Set
    End Property
    Public Property xLgrDest() As String
        Get
            Return _xLgrDest
        End Get
        Set(value As String)
            _xLgrDest = value
        End Set
    End Property
    Public Property nroDest() As String
        Get
            Return _nroDest
        End Get
        Set(value As String)
            _nroDest = value
        End Set
    End Property
    Public Property xCplDest() As String
        Get
            Return _xCplDest
        End Get
        Set(value As String)
            _xCplDest = value
        End Set
    End Property
    Public Property xBairroDest() As String
        Get
            Return _xBairroDest
        End Get
        Set(value As String)
            _xBairroDest = value
        End Set
    End Property
    Public Property cMunDest() As String
        Get
            Return _cMunDest
        End Get
        Set(value As String)
            _cMunDest = value
        End Set
    End Property
    Public Property xMunDest() As String
        Get
            Return _xMunDest
        End Get
        Set(value As String)
            _xMunDest = value
        End Set
    End Property
    Public Property UFDest() As String
        Get
            Return _UFDest
        End Get
        Set(value As String)
            _UFDest = value
        End Set
    End Property
    Public Property CEPDest() As String
        Get
            Return _CEPDest
        End Get
        Set(value As String)
            _CEPDest = value
        End Set
    End Property
    Public Property cPaisDest() As String
        Get
            Return _cPaisDest
        End Get
        Set(value As String)
            _cPaisDest = value
        End Set
    End Property
    Public Property xPaisDest() As String
        Get
            Return _xPaisDest
        End Get
        Set(value As String)
            _xPaisDest = value
        End Set
    End Property
    Public Property foneDest() As String
        Get
            Return _foneDest
        End Get
        Set(value As String)
            _foneDest = value
        End Set
    End Property
    Public Property indIEdest() As String
        Get
            Return _indIEdest
        End Get
        Set(value As String)
            _indIEdest = value
        End Set
    End Property
    Public Property IEDest() As String
        Get
            Return _IEDest
        End Get
        Set(value As String)
            _IEDest = value
        End Set
    End Property
    Public Property ISUFDest() As String
        Get
            Return _ISUFDest
        End Get
        Set(value As String)
            _ISUFDest = value
        End Set
    End Property
    Public Property IMDest() As String
        Get
            Return _IMDest
        End Get
        Set(value As String)
            _IMDest = value
        End Set
    End Property
    Public Property email() As String
        Get
            Return _email
        End Get
        Set(value As String)
            _email = value
        End Set
    End Property
#End Region

#Region "Item da Nota"
    Public Property nItem() As String
        Get
            Return _nItem
        End Get
        Set(value As String)
            _nItem = value
        End Set
    End Property
    Public Property cProd() As String
        Get
            Return _cProd
        End Get
        Set(value As String)
            _cProd = value
        End Set
    End Property
    Public Property cEAN() As String
        Get
            Return _cEAN
        End Get
        Set(value As String)
            _cEAN = value
        End Set
    End Property
    Public Property xProd() As String
        Get
            Return _xProd
        End Get
        Set(value As String)
            _xProd = value
        End Set
    End Property
    Public Property NCM() As String
        Get
            Return _NCM
        End Get
        Set(value As String)
            _NCM = value
        End Set
    End Property
    Public Property CFOP() As String
        Get
            Return _CFOP
        End Get
        Set(value As String)
            _CFOP = value
        End Set
    End Property
    Public Property uCom() As String
        Get
            Return _uCom
        End Get
        Set(value As String)
            _uCom = value
        End Set
    End Property
    Public Property qCom() As String
        Get
            Return _qCom
        End Get
        Set(value As String)
            _qCom = value
        End Set
    End Property
    Public Property vUnCom() As String
        Get
            Return _vUnCom
        End Get
        Set(value As String)
            _vUnCom = value
        End Set
    End Property
    Public Property vProd() As String
        Get
            Return _vProd
        End Get
        Set(value As String)
            _vProd = value
        End Set
    End Property
    Public Property cEANTrib() As String
        Get
            Return _cEANTrib
        End Get
        Set(value As String)
            _cEANTrib = value
        End Set
    End Property
    Public Property uTrib() As String
        Get
            Return _uTrib
        End Get
        Set(value As String)
            _uTrib = value
        End Set
    End Property
    Public Property qTrib() As String
        Get
            Return _qTrib
        End Get
        Set(value As String)
            _qTrib = value
        End Set
    End Property
    Public Property vUnTrib() As String
        Get
            Return _vUnTrib
        End Get
        Set(value As String)
            _vUnTrib = value
        End Set
    End Property
    Public Property vFrete() As String
        Get
            Return _vFrete
        End Get
        Set(value As String)
            _vFrete = value
        End Set
    End Property
    Public Property indTot() As String
        Get
            Return _indTot
        End Get
        Set(value As String)
            _indTot = value
        End Set
    End Property
#End Region

    Public Sub New()

    End Sub

    Public Sub addCabecalho()
        Dim Chave As String
        Dim sDv As String
        Dim sChave, Id As String

        Chave = ChaveNFe(cUF, dhEmi, CNPJ, "55", Microsoft.VisualBasic.Strings.Format(Convert.ToInt16(serie), "000"), nNF, tpRmis, nNF)

        sChave = ""
        sChaveNota = ""

        sChave = Right(Chave, 44)

        ChaveNota = sChave
        sChaveNota = sChave

        Id = "" : sDv = "" : cDV = ""

        Id = sChave
        sDv = Microsoft.VisualBasic.Strings.Right(Chave, 1)
        cDV = sDv


        Dim newcNF As String = Microsoft.VisualBasic.Strings.Mid(sChave, 36, 8)
        NomeArq = sNota & sChave & "NFe.xml"

        wr = New XmlTextWriter(NomeArq, System.Text.Encoding.GetEncoding("UTF-8"))
        wr.Formatting = Formatting.Indented

        wr.WriteStartDocument()
        wr.WriteStartElement("NFe")
        wr.WriteAttributeString("xmlns", "http://www.portalfiscal.inf.br/nfe")

        'Escreve os sub-elementos da tag de informaçãoes 
        wr.WriteStartElement("infNFe")
        wr.WriteAttributeString("versao", VersaoNota.Trim())
        wr.WriteAttributeString("Id", "NFe" + sChave.Trim())


        'Escreve os sub-elementos da Identificação da Nota 
        wr.WriteStartElement("ide")
        wr.WriteElementString("cUF", cUF)
        wr.WriteElementString("cNF", newcNF.Trim()) 'codigo da chave de acesso
        wr.WriteElementString("natOp", natOp.Trim())
        wr.WriteElementString("indPag", indPag)
        wr.WriteElementString("mod", smod)
        wr.WriteElementString("serie", serie)
        wr.WriteElementString("nNF", nNF)
        wr.WriteElementString("dhEmi", dhEmi + "-03:00")
        wr.WriteElementString("dhSaiEnt", dhSaiEnt + "-03:00")
        wr.WriteElementString("tpNF", tpNF)
        wr.WriteElementString("idDest", idDest)
        wr.WriteElementString("cMunFG", cMunFG)
        wr.WriteElementString("tpImp", tpImp)
        wr.WriteElementString("tpEmis", tpRmis)
        wr.WriteElementString("cDV", cDV)
        wr.WriteElementString("tpAmb", tpAmb)
        wr.WriteElementString("finNFe", finNFe)
        wr.WriteElementString("indFinal", finFinal)
        wr.WriteElementString("indPres", indPres)
        wr.WriteElementString("procEmi", procEmi)
        wr.WriteElementString("verProc", verPros)
        wr.WriteEndElement()



    End Sub

    Public Sub addEmitente()
        'Dados do Emitente
        wr.WriteStartElement("emit")
        wr.WriteElementString("CNPJ", CNPJ.Trim())
        wr.WriteElementString("xNome", xNome.Trim())
        wr.WriteElementString("xFant", xFant.Trim())
        wr.WriteStartElement("enderEmit")
        wr.WriteElementString("xLgr", xLgr.Trim())
        wr.WriteElementString("nro", nro.Trim())
        'wr.WriteElementString("xCpl", xCpl.Trim())
        wr.WriteElementString("xBairro", xBairro.Trim())
        wr.WriteElementString("cMun", cMun.Trim())
        wr.WriteElementString("xMun", "BRASILIA")
        wr.WriteElementString("UF", UF.Trim())
        wr.WriteElementString("CEP", CEP.Trim())
        wr.WriteElementString("cPais", "1058")
        wr.WriteElementString("xPais", xPais.Trim())
        'wr.WriteElementString("fone", Util.RetornaSoNumero(fone.Trim()))
        wr.WriteEndElement()
        wr.WriteElementString("IE", IE)
        wr.WriteElementString("IM", IM)
        wr.WriteElementString("CNAE", CNAE)
        wr.WriteElementString("CRT", "3")

        wr.WriteEndElement()
    End Sub

    Public Sub addDestinatario()
        'Identificação do Destinatario
        wr.WriteStartElement("dest")
        wr.WriteElementString("CPF", CPFDest)
        wr.WriteElementString("xNome", xNomeDest.Trim())
        wr.WriteStartElement("enderDest")
        wr.WriteElementString("xLgr", xLgrDest.Trim())
        wr.WriteElementString("nro", nroDest.Trim())
        wr.WriteElementString("xBairro", xBairroDest.ToUpper())
        wr.WriteElementString("cMun", cMunDest.Trim())
        wr.WriteElementString("xMun", xMunDest.Trim())
        wr.WriteElementString("UF", UFDest.Trim())
        wr.WriteElementString("CEP", CEPDest.Trim())
        wr.WriteElementString("cPais", "1058")
        wr.WriteElementString("xPais", "BRASIL")
        wr.WriteElementString("fone", foneDest.Trim())
        wr.WriteEndElement()
        wr.WriteElementString("indIEDest", indIEdest)
        wr.WriteEndElement()
    End Sub

    Public Sub addItemNota()
        wr.WriteStartElement("det")
        wr.WriteAttributeString("nItem", Convert.ToString(nItem))
        wr.WriteStartElement("prod")
        wr.WriteElementString("cProd", cProd)
        wr.WriteElementString("cEAN", String.Empty)
        wr.WriteElementString("xProd", xProd.Trim())
        wr.WriteElementString("NCM", "00")
        wr.WriteElementString("CFOP", CFOP.Trim())
        wr.WriteElementString("uCom", uCom.Trim())
        wr.WriteElementString("qCom", qCom)
        wr.WriteElementString("vUnCom", FormataDecimal(vUnCom.Trim()))
        wr.WriteElementString("vProd", FormataDecimal(vProd.Trim()))
        wr.WriteElementString("cEANTrib", String.Empty)
        wr.WriteElementString("uTrib", uTrib)
        wr.WriteElementString("qTrib", qTrib)
        wr.WriteElementString("vUnTrib", FormataDecimal(vUnTrib.Trim()))
        wr.WriteElementString("indTot", "1")
        wr.WriteEndElement()
    End Sub

    Public Sub addImpostos()

        wr.WriteStartElement("imposto")
        wr.WriteStartElement("ICMS")
        wr.WriteStartElement("ICMS00")
        wr.WriteElementString("orig", "0")
        wr.WriteElementString("CST", "00")
        wr.WriteElementString("modBC", "3")
        If vICMS = "" Then

            wr.WriteElementString("pICMS", "0.00")
            wr.WriteElementString("vICMS", "0.00")

        Else

            wr.WriteElementString("pICMS", "4.0000")
            wr.WriteElementString("vICMS", FormataDecimal(vICMS.Trim()))
        End If
        wr.WriteEndElement()
        wr.WriteEndElement()

        wr.WriteStartElement("PIS")
        wr.WriteStartElement("PISAliq")
        wr.WriteElementString("CST", "01")
        wr.WriteElementString("vBC", FormataDecimal(vProd.Trim()))
        wr.WriteElementString("pPIS", "0.0000")
        wr.WriteElementString("vPIS", "0.00")
        wr.WriteEndElement()
        wr.WriteEndElement()


        wr.WriteStartElement("COFINS")
        wr.WriteStartElement("COFINSAliq")
        wr.WriteElementString("CST", "01")
        wr.WriteElementString("vBC", FormataDecimal(vProd.Trim()))
        wr.WriteElementString("pCOFINS", "0.0000")
        wr.WriteElementString("vCOFINS", "0.00")
        wr.WriteEndElement()
        wr.WriteEndElement()


        wr.WriteEndElement()
        wr.WriteEndElement()

    End Sub

    Public Sub addTotalNota()
        wr.WriteStartElement("total")
        wr.WriteStartElement("ICMSTot")
        wr.WriteElementString("vBC", FormataDecimal(vUnTrib.Trim()))
        wr.WriteElementString("vICMS", FormataDecimal(vICMS.Trim()))
        wr.WriteElementString("vICMSDeson", "0.00")
        wr.WriteElementString("vFCP", "0.00")
        wr.WriteElementString("vBCST", "0.00")
        wr.WriteElementString("vST", "0.00")
        wr.WriteElementString("vFCPST", "0.00")
        wr.WriteElementString("vFCPSTRet", "0.00")
        wr.WriteElementString("vProd", "0.00")
        wr.WriteElementString("vFrete", "0.00")
        wr.WriteElementString("vSeg", "0.00")
        wr.WriteElementString("vDesc", "0.00")
        wr.WriteElementString("vII", "0.00")
        wr.WriteElementString("vIPI", "0.00")
        wr.WriteElementString("vIPIDevol", "0.00")
        wr.WriteElementString("vPIS", "0.00")
        wr.WriteElementString("vCOFINS", "0.00")
        wr.WriteElementString("vOutro", "0.00")
        wr.WriteElementString("vNF", FormataDecimal(vUnTrib.Trim()))
        wr.WriteEndElement()


        wr.WriteStartElement("transp")
        wr.WriteElementString("modFrete", "0")
        wr.WriteStartElement("vol")
        wr.WriteElementString("qVol", "0")
        wr.WriteEndElement()
        wr.WriteEndElement()

        wr.WriteStartElement("cobr")
        wr.WriteStartElement("fat")
        wr.WriteElementString("nFat", "0")
        wr.WriteElementString("vOrig", FormataDecimal(vUnTrib.Trim()))
        wr.WriteElementString("vDesc", "0")
        wr.WriteElementString("vLiq", FormataDecimal(vUnTrib.Trim()))
        wr.WriteEndElement()

        wr.WriteStartElement("dup")
        wr.WriteElementString("nDup", "001")
        wr.WriteElementString("dVenc", dhEmi.Substring(0, 10))
        wr.WriteElementString("vDup", FormataDecimal(vUnTrib.Trim()))
        wr.WriteEndElement()
        wr.WriteEndElement()

        wr.WriteStartElement("pag")
        wr.WriteStartElement("detPag")
        wr.WriteElementString("indPag", "0")
        wr.WriteElementString("tPag", "01")
        wr.WriteElementString("vPag", FormataDecimal(vUnTrib.Trim()))
        wr.WriteEndElement()
        wr.WriteEndElement()




        wr.WriteStartElement("infAdic")
        wr.WriteElementString("infCpl", "Procon DF F-151 VEN 2000 BL B-60 S-240 SCS - Tel: 151; REFERENTE MENSALIDADES DE NOVEMBRO E DEZEMBRO DE 2019 - CPD 21861 - SERVICOS PRESTADOS NA DEPENDENCIA DA CONTRATADA (SEM RETENCAO DO INSS, CONFORME ART. 149, INCISO VI, I N RFB 971/2009). ISSQN - IMUNE, CONFORME ATO DECLARTORIO N 60/2012 - SEFAZ. NAO HA RETENCAO DE TRIBUTOS FEDERAIS, CONFORME DECLARACAO DE IMUNIDADE E ISENCAO POR OCASIAO DE ADESAO AO PROUNI (LEI N 11.096, DE 13 DE JANEIRO DE 2005).;Empresa Simples Nacional Nao gera credito de ICMS.")
        wr.WriteEndElement()



        wr.WriteEndDocument()

        wr.Close()


    End Sub

    Public Sub FinalizaXml(ByRef NomeCert As String)
        Dim Assinatura As New AssinarXml

        If Assinatura.AssinandoNota(ChaveNota, NomeArq, NomeCert) = True Then

        End If

    End Sub




End Class

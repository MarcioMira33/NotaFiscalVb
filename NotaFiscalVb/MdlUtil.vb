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
Imports System.Data
Imports System.Data.SqlClient

Module MdlUtil

    Public sGerados As String = "c:\NFe\Lotes\Gerados\"
    Public sEnviados As String = "c:\NFe\Lotes\Enviados\"
    Public sNota As String = "c:\NFe\Lotes\NFe\"
    Public sRequisicoes As String = "c:\NFe\Lotes\Requisicoes\"
    Public sRetonos As String = "c:\NFe\Lotes\Retornos\"
    Public sCancelamento As String = "c:\NFe\Lotes\Canceladas\"

    Public TipoEnvio As Integer


    Public sUrlQrCode As String = "http://www.sefaz.mt.gov.br/nfce/consultanfce"

    Public VersaoNota As String = "4.00"
    Public TpAmb As Integer = 1
    Public cUf As String = 53
    Public Serie As String = 1

    Public XMLStringAssinado As String = ""

    Public NomeCertificado As String
    Public sChaveNota As String
    Public Function GeraSha1(sTexto As String) As String

        Dim strToHash As String = sTexto
        strToHash = Command$()

        Dim sha1Obj As New System.Security.Cryptography.SHA1CryptoServiceProvider
        Dim bytesToHash() As Byte = System.Text.Encoding.ASCII.GetBytes(strToHash)

        bytesToHash = sha1Obj.ComputeHash(bytesToHash)

        Dim strResult As String = ""

        For Each b As Byte In bytesToHash
            strResult = strResult & b.ToString("x2")
        Next

        Return strResult
    End Function

   
    Public Function GreraQrCode(sChave As String, sVersaoQr As String, stpAmb As String, sCsc As String, sHash As String) As String
        Dim sResposta As String

        'so falta colocar o CSC que tem qie ver com o contador

        Dim sTexto As String = "chNFe=" & sChave & "&nVersao=" & sVersaoQr & "&tpAmb=" & stpAmb & "&CSC=" & sCsc & "&cHashQRCode=" & sHash
        sResposta = ""

        sHash = GeraSha1(sTexto)

        sResposta = "![CDATA[" & "http://" & sUrlQrCode & "/nfce/qrcode?p=" & "&chNFe=" & sChave & "&nVersao=" & sVersaoQr & "&tpAmb=" & stpAmb & "&CSC=" & sCsc & "&cHashQRCode=" & sHash & "]]"

        Return sResposta
    End Function
    Public Sub XmlDistribuicao(schave As String, sNomeArqAssi As String, sNewNome As String)
        Dim sql As String
        Dim Dr As SqlDataReader

        sql = ""

        sql = "select * from LoteNota where ChaveNFe = '" & schave.Trim() & "' "
        'colocar aqui o Dr

        Dim wr As New XmlTextWriter(sNewNome, Encoding.GetEncoding("UTF-8"))

        'If Dr.Read() Then
        '    With wr

        '        .WriteStartDocument()
        '        .WriteStartElement("nfeProc")
        '        .WriteAttributeString("versao", VersaoNota)
        '        .WriteAttributeString("xmlns", "http://www.portalfiscal.inf.br/nfe")
        '        .WriteStartElement("protNFe")
        '        .WriteAttributeString("versao", VersaoNota)
        '        .WriteStartElement("infProt")
        '        .WriteElementString("tpAmb", TpAmb.ToString())
        '        .WriteElementString("verAplic", Dr["VerAplic"].ToString().Trim())
        '        .WriteElementString("chNFe", Dr["ChaveNFe"].ToString().Trim())
        '        .WriteElementString("dhRecbto", Convert.ToDateTime(Dr["dhRecibo"]).ToString("yyyy-MM-ddThh:mm:ss") + "-03:00")
        '        .WriteElementString("nProt", Dr["NumProtocolo"].ToString().Trim())
        '        .WriteElementString("digVal", Dr["DigVal"].ToString().Trim())
        '        .WriteElementString("cStat", Dr["Status"].ToString().Trim())
        '        .WriteElementString("xMotivo", Dr["Motivo"].ToString().Trim())
        '        .WriteEndElement();
        '        .WriteEndElement();
        '        .WriteEndElement();
        '        .Close();

        '    End With
        'End If

        Dim DocLote As New XmlDocument()
        Dim DocNfe As New XmlDocument()

        Dim elementsByTagName As XmlNodeList
        Dim NodeNfe As XmlNodeList

        Dim _stringXml As String
        Dim sr2 As StreamReader
        Dim sr As StreamReader

        Dim sw2 As StreamWriter


        sr2 = File.OpenText(sNewNome)
        _stringXml = sr2.ReadToEnd()
        sr2.Close()

        sr = File.OpenText(sNewNome)
        _stringXml = sr.ReadToEnd()
        sr.Close()


        DocLote.LoadXml(_stringXml)
        DocNfe.Load(sNomeArqAssi)
        DocNfe.PreserveWhitespace = True
        DocLote.PreserveWhitespace = True
        elementsByTagName = DocLote.GetElementsByTagName("nfeProc")
        'adiciona o node NFe no envNFe
        NodeNfe = DocNfe.GetElementsByTagName("NFe")
        elementsByTagName.Item(0).PrependChild(DocLote.ImportNode(NodeNfe.Item(0), True))


        'salva o arquivo
        sw2 = File.CreateText(sNewNome)
        sw2.Write(DocLote.OuterXml)
        sw2.Close()


    End Sub
    Public Sub AtualizaLote(cStat As String, xMotivo As String, dhRecibo As String, nRecibo As String, NumLote As Integer)
        Dim sql As String
        Try
            sql = "update Lote set "
            sql = sql & " status =  '" & cStat & "',"
            sql = sql & " Motivo =  '" & xMotivo & "',"
            sql = sql & " dhRecibo =  '" & dhRecibo & "',"
            sql = sql & " Recibo  = '" & nRecibo & "'"
            sql = sql & " where"
            sql = sql & " Lote =  " & NumLote & ""
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Public Sub AtualizaLoteNota(chNFe As String, dhRecbtoNf As String, protNfe As String, digVal As String, cStatNf As String,
                                xMotivo As String, VerAplic As String, NumLote As Integer, TipoMsg As Boolean)

        Dim sql As String
        Try
            sql = "update LoteNota set "
            sql = sql & " ChaveNFe = '" & chNFe & "' ,"
            sql = sql & " dhRecibo  = '" & dhRecbtoNf & "' ,"
            If TipoMsg = False Then
                sql = sql & " NumProtocolo = '" & protNfe & "' ,"
            End If
            sql = sql & " DigVal = '" & digVal & "' ,"
            sql = sql & " Status = '" & cStatNf & "'  ,"
            sql = sql & " Motivo = '" & xMotivo & "' ,"
            sql = sql & " VerAplic = '" & VerAplic & "'"
            sql = sql & " where"
            sql = sql & " fkNumLote = " & NumLote & "  and"
            sql = sql & " ChaveNFe = '" & chNFe & "'"
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub

    Public Function FormataDecimal(sString As String) As String
        Dim newNumber As String = ""
        newNumber = Replace(Microsoft.VisualBasic.Strings.Format(Convert.ToDouble(sString), "##########0.00"), ",", ".")
        Return newNumber
    End Function

    Public Function ChaveNFe(ByVal cUF As String, ByVal cdata As String, ByVal CNPJ As String, ByVal Modelo As String, ByVal Serie As String, ByVal NumNota As String, ByVal tpEmi As String, ByVal NumDoc As String) As String
        Dim Dv As String
        Dim Resultado As String
        Dim data As String
        Dim rndNumero As New Random()
        Dim numero As String = ""
        numero = rndNumero.Next().ToString()

        data = Mid(cdata, 3, 2) + Microsoft.VisualBasic.Strings.Mid(cdata, 6, 2)

        Resultado = Format(Convert.ToInt16(cUF), "00")
        Resultado = Resultado & data
        Resultado = Resultado & CNPJ
        Resultado = Resultado & Modelo
        Resultado = Resultado & Serie
        Resultado = Resultado & Format(Convert.ToInt16(NumNota), "000000000")
        Resultado = Resultado & tpEmi
        If Len(numero) < 8 Then
            Resultado = Resultado & Format(Convert.ToDouble(numero), "00000000")
        Else
            Resultado = Resultado & Left(numero, 8)
        End If

        Dv = modulo11(Resultado)

        Return ("NFe" & Resultado & Dv).Trim()


    End Function

    Public Function modulo11(ByVal strnumero As String) As String

        Dim Base As Integer
        Dim Resto As Integer
        Dim Soma, Contador, Peso, Digito As Integer

        Base = 9
        Soma = 0
        Peso = 2

        For Contador = Len(strnumero) To 1 Step -1
            Soma = Soma + Val(Mid(strnumero, Contador, 1)) * Peso
            If Peso < Base Then
                Peso = Peso + 1
            Else
                Peso = 2
            End If
        Next

        Resto = (Soma Mod 11)

        If Resto = 0 Then
            Digito = 0
        ElseIf Resto = 1 Then
            Digito = 0
        Else
            Digito = 11 - Resto
            If Digito > 9 Then
                Digito = 0
            End If
        End If

        modulo11 = CStr(Digito).Trim


    End Function

    Public Function XmlDistribuicao(ByVal sNomeArqAssi As String, ByVal sNewNome As String,
                                    ByVal VerAplic As String, ByVal chNFe As String, ByVal dhRecbto As String, ByVal nProt As String, _
                                    ByVal digVal As String, cStat As String, ByVal xMotivo As String) As Boolean

        Dim bDistb As Boolean

        Try
            bDistb = True

            Dim wr As New XmlTextWriter(sNewNome, Encoding.GetEncoding("UTF-8"))


            wr.WriteStartDocument()
            wr.WriteStartElement("nfeProc")
            wr.WriteAttributeString("versao", VersaoNota)
            wr.WriteAttributeString("xmlns", "http://www.portalfiscal.inf.br/nfe")
            wr.WriteStartElement("protNFe")
            wr.WriteAttributeString("versao", VersaoNota.Trim())
            wr.WriteStartElement("infProt")
            wr.WriteElementString("tpAmb", TpAmb.ToString().Trim())
            wr.WriteElementString("verAplic", VerAplic.Trim())
            wr.WriteElementString("chNFe", chNFe.Trim())
            wr.WriteElementString("dhRecbto", dhRecbto + "-03:00")
            wr.WriteElementString("nProt", nProt.Trim())
            wr.WriteElementString("digVal", digVal.Trim())
            wr.WriteElementString("cStat", cStat.Trim())
            wr.WriteElementString("xMotivo", xMotivo.Trim())
            wr.WriteEndElement()
            wr.WriteEndElement()
            wr.WriteEndElement()
            wr.Close()


            Dim DocLote As New XmlDocument()
            Dim DocNfe As New XmlDocument()

            Dim elementsByTagName As XmlNodeList
            Dim NodeNfe As XmlNodeList

            Dim _stringXml As String
            Dim sr2 As StreamReader
            Dim sr As StreamReader

            Dim sw2 As StreamWriter


            sr2 = File.OpenText(sNewNome)
            _stringXml = sr2.ReadToEnd()
            sr2.Close()

            sr = File.OpenText(sNewNome)
            _stringXml = sr.ReadToEnd()
            sr.Close()


            DocLote.LoadXml(_stringXml)
            DocNfe.Load(sNomeArqAssi)
            DocNfe.PreserveWhitespace = True
            DocLote.PreserveWhitespace = True
            elementsByTagName = DocLote.GetElementsByTagName("nfeProc")
            'adiciona o node NFe no envNFe
            NodeNfe = DocNfe.GetElementsByTagName("NFe")
            elementsByTagName.Item(0).PrependChild(DocLote.ImportNode(NodeNfe.Item(0), True))


            'salva o arquivo
            sw2 = File.CreateText(sNewNome)
            sw2.Write(DocLote.OuterXml)
            sw2.Close()



        Catch ex As Exception
            bDistb = False
        End Try

        Return bDistb

    End Function


End Module

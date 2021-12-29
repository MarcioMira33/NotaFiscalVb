Imports System
Imports System.Security.Cryptography
Imports System.Security.Cryptography.Xml
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Xml

Imports System.Xml.Schema
Imports System.Net
Imports System.Data.SqlClient
Imports System.Threading.Tasks
Imports System.Data
Imports System.Runtime.InteropServices


Imports System.IO
Imports System.Collections.Generic
Imports System.Linq
Imports System.ServiceModel

Public Class AssinarXml

    Private XMLDoc As XmlDocument

    Public Function BuscarNome(ByVal Nome As String) As X509Certificate2

        Dim _X509Cert As New X509Certificate2()
        Dim collection, collection1, collection2 As X509Certificate2Collection
        Dim scollection As X509Certificate2Collection

        Try
            Dim Store As New X509Store("MY", StoreLocation.CurrentUser)

            Store.Open(OpenFlags.OpenExistingOnly)
            collection = Store.Certificates
            collection1 = collection.Find(X509FindType.FindByTimeValid, DateTime.Now, False)
            collection2 = collection.Find(X509FindType.FindByKeyUsage, X509KeyUsageFlags.DigitalSignature, False)

            If Nome = "" Then
                scollection = X509Certificate2UI.SelectFromCollection(collection2, "Certificado(s) Digital(is) disponível(is)", "Selecione o Certificado Digital para uso no aplicativo", X509SelectionFlag.SingleSelection)
                If scollection.Count = 0 Then
                    _X509Cert.Reset()
                Else
                    _X509Cert = scollection(0)
                End If
            Else
                scollection = collection2.Find(X509FindType.FindBySubjectDistinguishedName, Nome, False)
                If scollection.Count = 0 Then
                    _X509Cert.Reset()
                Else
                    _X509Cert = scollection(0)
                End If

            End If
            Store.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return _X509Cert

    End Function

    Public Function AssinandoNota(ByVal Nota As String, ByVal CamArq As String, ByRef Nomecert As String) As Boolean
        Dim Assinado As Boolean = True

        Try

            Dim rfUri, strXml As String
            Dim sr As StreamReader
            Dim sw As StreamWriter
            Dim cert As New X509Certificate2()

            If CamArq = "" Then
                MsgBox("Caminho do arquivo não informado!!")
                Assinado = False
            Else
                rfUri = "infNFe"
                'ler Arquivo Xml a ser Assinado  
                sr = File.OpenText(CamArq)
                strXml = sr.ReadToEnd()
                sr.Close()

                cert = BuscarNome(Nomecert)

                If Assinar(strXml, rfUri, cert) Then
                    sw = File.CreateText(CamArq)
                    sw.Write(XMLStringAssinado)
                    sw.Close()
                Else
                    MessageBox.Show("Erro ao assinar a nota")
                    Assinado = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Assinado = False
        End Try
        Return Assinado
    End Function

    Public Function Assinar(ByVal XMLString As String, ByVal RefUri As String, ByVal X509Cert As X509Certificate2) As Boolean

        Dim _xnome, x As String
        Dim store As New X509Store("MY", StoreLocation.CurrentUser)
        Dim collection, collection1 As X509Certificate2Collection
        Dim _X509Cert As New X509Certificate2()
        Dim qtdeRefUri As Integer

        Try
            _xnome = ""
            _xnome = X509Cert.Subject.ToString()
            store.Open(OpenFlags.OpenExistingOnly)
            collection = store.Certificates
            collection1 = collection.Find(X509FindType.FindBySubjectDistinguishedName, _xnome, False)


            _X509Cert = collection(0)
            x = _X509Cert.GetKeyAlgorithm().ToString()
            Dim XMLDoc As New XmlDocument()

            Dim doc As New XmlDocument()
            doc.PreserveWhitespace = False

            doc.LoadXml(XMLString)

            qtdeRefUri = doc.GetElementsByTagName(RefUri.Trim()).Count

            If qtdeRefUri = 0 Or qtdeRefUri > 1 Then
                MessageBox.Show("Erro quatidade de tag de referencia")
                Return False
            End If

            Dim _Uri As XmlAttributeCollection

            Dim reference As New Reference()
            Dim env As New XmlDsigEnvelopedSignatureTransform()
            Dim c14 As New XmlDsigC14NTransform()
            Dim keyInfo As New KeyInfo

            Dim xmlDigitalSignature As XmlElement

            Dim signedXml As New SignedXml(doc)

            signedXml.SigningKey = _X509Cert.PrivateKey
            _Uri = doc.GetElementsByTagName(RefUri.Trim()).Item(0).Attributes
            Dim _atributo As XmlAttribute
            For Each _atributo In _Uri
                If _atributo.Name = "Id" Then
                    reference.Uri = "#" + _atributo.InnerText
                End If
            Next
            'Add an enveloped transformation to the reference.
            reference.AddTransform(env)
            reference.AddTransform(c14)
            'Add the reference to the SignedXml object.
            signedXml.AddReference(reference)
            'Create a new KeyInfo object
            ' Load the certificate into a KeyInfoX509Data object
            'and add it to the KeyInfo object.
            KeyInfo.AddClause(New KeyInfoX509Data(_X509Cert))

            'Add the KeyInfo object to the SignedXml object.
            signedXml.KeyInfo = KeyInfo
            signedXml.ComputeSignature()



            ' Get the XML representation of the signature and save
            'it to an XmlElement object.
            xmlDigitalSignature = signedXml.GetXml()

            doc.DocumentElement.AppendChild(doc.ImportNode(xmlDigitalSignature, True))

            ''Append the element to the XML document.
            'If RefUri.Trim = "infNFe" Then
            '    doc.DocumentElement.LastChild.AppendChild(doc.ImportNode(xmlDigitalSignature, True))
            'Else
            '    doc.DocumentElement.AppendChild(doc.ImportNode(xmlDigitalSignature, True))
            'End If

            XMLDoc = New XmlDocument()
            XMLDoc.PreserveWhitespace = False
            XMLDoc = doc
            XMLStringAssinado = XMLDoc.OuterXml


        Catch ex As Exception
            MsgBox(ex.Message)
            Return False

        End Try

        Return True

    End Function





End Class

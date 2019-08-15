Imports System.Xml
Imports Microsoft.VisualBasic

Public Class brandInfo

    Private _Loaded As Boolean = False

    Private _BrandID As String = ""
    Private _Name As String = ""
    Private _colorR As String = ""
    Private _colorG As String = ""
    Private _colorB As String = ""
    Private _colorC As String = ""
    Private _colorM As String = ""
    Private _colorY As String = ""
    Private _colorK As String = ""
    Private _Logo As String = ""
    Private _PrivateLabel As Boolean = False
    Private _ShowSectionOnSheet As Boolean = False

    Public ReadOnly Property Loaded As Boolean
        Get
            Return _Loaded
        End Get
    End Property

    Public ReadOnly Property Name As String
        Get
            Return _Name
        End Get
    End Property

    Public ReadOnly Property colorR As String
        Get
            Return _colorR
        End Get
    End Property

    Public ReadOnly Property colorG As String
        Get
            Return _colorG
        End Get
    End Property

    Public ReadOnly Property colorB As String
        Get
            Return _colorB
        End Get
    End Property

    Public ReadOnly Property colorC As String
        Get
            Return _colorC
        End Get
    End Property

    Public ReadOnly Property colorM As String
        Get
            Return _colorM
        End Get
    End Property

    Public ReadOnly Property colorY As String
        Get
            Return _colorY
        End Get
    End Property

    Public ReadOnly Property colorK As String
        Get
            Return _colorK
        End Get
    End Property

    Public ReadOnly Property Logo As String
        Get
            Return _Logo
        End Get
    End Property

    Public ReadOnly Property PrivateLabel As Boolean
        Get
            Return _PrivateLabel
        End Get
    End Property

    Public ReadOnly Property ShowSectionOnSheet As Boolean
        Get
            Return _ShowSectionOnSheet
        End Get
    End Property



    Public Sub New(ByVal BrandID As String)

        _BrandID = BrandID

        Dim webFolder As String = System.Configuration.ConfigurationManager.AppSettings("WebPath")
        Dim LanguageInfo = XDocument.Load(webFolder & System.Configuration.ConfigurationManager.AppSettings("XML_BrandFile"))

        Dim docXML As XmlDocument = New XmlDocument()
        docXML.LoadXml(LanguageInfo.ToString)

        Dim itemRecord As XmlNodeList = docXML.SelectNodes("descendant::brand[id='" & BrandID & "']")
        If itemRecord.Count = 1 Then
            Dim brandInfo As XmlNode = itemRecord(0)

            _Name = (brandInfo.SelectSingleNode("name").InnerText)
            _colorR = (brandInfo.SelectSingleNode("colorR").InnerText)
            _colorG = (brandInfo.SelectSingleNode("colorG").InnerText)
            _colorB = (brandInfo.SelectSingleNode("colorB").InnerText)
            _colorC = (brandInfo.SelectSingleNode("colorC").InnerText)
            _colorM = (brandInfo.SelectSingleNode("colorM").InnerText)
            _colorY = (brandInfo.SelectSingleNode("colorY").InnerText)
            _colorK = (brandInfo.SelectSingleNode("colorK").InnerText)
            _Logo = (brandInfo.SelectSingleNode("logo").InnerText)
            _PrivateLabel = CBool((brandInfo.SelectSingleNode("privatelabel").InnerText))
            _ShowSectionOnSheet = CBool((brandInfo.SelectSingleNode("showsectioninfo").InnerText))

            _Loaded = True
        End If

    End Sub

End Class


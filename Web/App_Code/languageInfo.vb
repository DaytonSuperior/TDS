Imports System.Xml
Imports Microsoft.VisualBasic

Public Class languageInfo

    Private _Loaded As Boolean = False

    Private _BrandID As String = ""
    Private _Language As String = ""
    Private _footer1 As String = ""
    Private _footer2 As String = ""
    Private _footer3 As String = ""
    Private _footer4 As String = ""
    Private _chemfooterMoreInfo As String = ""
    Private _chemPageType As String = ""
    Private _chemManufacturerHeader As String = ""
    Private _chemManufacturer As String = ""
    Private _chemWarrantyHeader As String = ""
    Private _chemWarranty1 As String = ""
    Private _chemWarranty2 As String = ""
    Private _accPageType As String = ""
    Private _accManufacturerHeader As String = ""
    Private _accManufacturer As String = ""
    Private _cartDocHeader As String = ""
    Private _cartTOC As String = ""
    Private _cartTOCDocInfo As String = ""
    Private _cartTOCProducts As String = ""
    Private _cartHeader As String = ""
    Private _cartSubHeader As String = ""
    Private _cartPageType As String = ""
    Private _cartLastPage As String = ""
    Private _cartSafetyPage As String = ""
    Private _cartCopy As String = ""
    Private _cartManufacturerHeader As String = ""
    Private _cartManufacturer As String = ""
    Private _cartPreparedHeader As String = ""
    Private _cartWarrantyHeader As String = ""
    Private _cartOrderingInfo As String = ""
    Private _cartOrderProduct As String = ""
    Private _cartOrderDesc As String = ""
    Private _cartOrderWeight As String = ""
    Private _cartFrontCover As String = ""
    Private _cartBackCover As String = ""
    Private _accWarrantyHeader As String = ""
    Private _accWarranty1 As String = ""
    Private _formWarrantyHeader As String = ""
    Private _formWarranty1 As String = ""
    Private _formPageType As String = ""
    Private _formManufacturerHeader As String = ""
    Private _formManufacturer As String = ""
    Private _formRentalHeader As String = ""
    Private _formRental As String = ""

    Public ReadOnly Property Loaded As Boolean
        Get
            Return _Loaded
        End Get
    End Property

    Public ReadOnly Property footer1 As String
        Get
            Return _footer1
        End Get
    End Property

    Public ReadOnly Property footer2 As String
        Get
            Return _footer2
        End Get
    End Property

    Public ReadOnly Property footer3 As String
        Get
            Return _footer3
        End Get
    End Property

    Public ReadOnly Property footer4 As String
        Get
            Return _footer4
        End Get
    End Property

    Public ReadOnly Property chemfooterMoreInfo As String
        Get
            Return _chemfooterMoreInfo
        End Get
    End Property

    Public ReadOnly Property chemPageType As String
        Get
            Return _chemPageType
        End Get
    End Property

    Public ReadOnly Property chemManufacturerHeader As String
        Get
            Return _chemManufacturerHeader
        End Get
    End Property

    Public ReadOnly Property chemManufacturer As String
        Get
            Return _chemManufacturer
        End Get
    End Property

    Public ReadOnly Property chemWarrantyHeader As String
        Get
            Return _chemWarrantyHeader
        End Get
    End Property

    Public ReadOnly Property chemWarranty1 As String
        Get
            Return _chemWarranty1
        End Get
    End Property

    Public ReadOnly Property chemWarranty2 As String
        Get
            Return _chemWarranty2
        End Get
    End Property

    Public ReadOnly Property accPageType As String
        Get
            Return _accPageType
        End Get
    End Property

    Public ReadOnly Property accManufacturerHeader As String
        Get
            Return _accManufacturerHeader
        End Get
    End Property

    Public ReadOnly Property accManufacturer As String
        Get
            Return _accManufacturer
        End Get
    End Property

    Public ReadOnly Property cartDocHeader As String
        Get
            Return _cartDocHeader
        End Get
    End Property

    Public ReadOnly Property cartTOC As String
        Get
            Return _cartTOC
        End Get
    End Property

    Public ReadOnly Property cartTOCDocInfo As String
        Get
            Return _cartTOCDocInfo
        End Get
    End Property

    Public ReadOnly Property cartTOCProducts As String
        Get
            Return _cartTOCProducts
        End Get
    End Property

    Public ReadOnly Property cartHeader As String
        Get
            Return _cartHeader
        End Get
    End Property

    Public ReadOnly Property cartSubHeader As String
        Get
            Return _cartSubHeader
        End Get
    End Property

    Public ReadOnly Property cartPageType As String
        Get
            Return _cartPageType
        End Get
    End Property

    Public ReadOnly Property cartLastPage As String
        Get
            Return _cartLastPage
        End Get
    End Property

    Public ReadOnly Property cartSafetyPage As String
        Get
            Return _cartSafetyPage
        End Get
    End Property

    Public ReadOnly Property cartCopy As String
        Get
            Return _cartCopy
        End Get
    End Property

    Public ReadOnly Property cartManufacturerHeader As String
        Get
            Return _cartManufacturerHeader
        End Get
    End Property

    Public ReadOnly Property cartManufacturer As String
        Get
            Return _cartManufacturer
        End Get
    End Property

    Public ReadOnly Property cartPreparedHeader As String
        Get
            Return _cartPreparedHeader
        End Get
    End Property

    Public ReadOnly Property cartWarrantyHeader As String
        Get
            Return _cartWarrantyHeader
        End Get
    End Property

    Public ReadOnly Property cartOrderingInfo As String
        Get
            Return _cartOrderingInfo
        End Get
    End Property

    Public ReadOnly Property cartOrderProduct As String
        Get
            Return _cartOrderProduct
        End Get
    End Property

    Public ReadOnly Property cartOrderDesc As String
        Get
            Return _cartOrderDesc
        End Get
    End Property

    Public ReadOnly Property cartOrderWeight As String
        Get
            Return _cartOrderWeight
        End Get
    End Property

    Public ReadOnly Property cartFrontCover As String
        Get
            Return _cartFrontCover
        End Get
    End Property

    Public ReadOnly Property cartBackCover As String
        Get
            Return _cartBackCover
        End Get
    End Property

    Public ReadOnly Property accWarrantyHeader As String
        Get
            Return _accWarrantyHeader
        End Get
    End Property

    Public ReadOnly Property accWarranty1 As String
        Get
            Return _accWarranty1
        End Get
    End Property

    Public ReadOnly Property formWarrantyHeader As String
        Get
            Return _formWarrantyHeader
        End Get
    End Property

    Public ReadOnly Property formWarranty1 As String
        Get
            Return _formWarranty1
        End Get
    End Property

    Public ReadOnly Property formPageType As String
        Get
            Return _formPageType
        End Get
    End Property

    Public ReadOnly Property formManufacturerHeader As String
        Get
            Return _formManufacturerHeader
        End Get
    End Property

    Public ReadOnly Property formManufacturer As String
        Get
            Return _formManufacturer
        End Get
    End Property

    Public ReadOnly Property formRentalHeader As String
        Get
            Return _formRentalHeader
        End Get
    End Property

    Public ReadOnly Property formRental As String
        Get
            Return _formRental
        End Get
    End Property

    Public Sub New(ByVal BrandID As String, ByVal Language As String)

        _BrandID = BrandID
        _Language = Language

        Dim webFolder As String = System.Configuration.ConfigurationManager.AppSettings("WebPath")
        Dim LanguageInfo = XDocument.Load(webFolder & System.Configuration.ConfigurationManager.AppSettings("XML_LanguageFile"))

        Dim docXML As XmlDocument = New XmlDocument()
        docXML.LoadXml(LanguageInfo.ToString)

        Dim itemRecord As XmlNodeList = docXML.SelectNodes("descendant::language[BrandID='" & BrandID & "' and Language='" & Language & "']")
        If itemRecord.Count = 1 Then
            Dim langInfo As XmlNode = itemRecord(0)

            _footer1 = (langInfo.SelectSingleNode("footer1").InnerText)
            _footer2 = (langInfo.SelectSingleNode("footer2").InnerText)
            _footer3 = (langInfo.SelectSingleNode("footer3").InnerText)
            _footer4 = (langInfo.SelectSingleNode("footer4").InnerText)
            _chemfooterMoreInfo = Trim(langInfo.SelectSingleNode("chemfooterMoreInfo").InnerText)
            _chemPageType = Trim(langInfo.SelectSingleNode("chemPageType").InnerText)
            _chemManufacturerHeader = Trim(langInfo.SelectSingleNode("chemManufacturerHeader").InnerText)
            _chemManufacturer = Trim(langInfo.SelectSingleNode("chemManufacturer").InnerText)
            _chemWarrantyHeader = Trim(langInfo.SelectSingleNode("chemWarrantyHeader").InnerText)
            _chemWarranty1 = Trim(langInfo.SelectSingleNode("chemWarranty1").InnerText)
            _chemWarranty2 = Trim(langInfo.SelectSingleNode("chemWarranty2").InnerText)
            _accPageType = Trim(langInfo.SelectSingleNode("accPageType").InnerText)
            _accManufacturerHeader = Trim(langInfo.SelectSingleNode("accManufacturerHeader").InnerText)
            _accManufacturer = Trim(langInfo.SelectSingleNode("accManufacturer").InnerText)
            _cartDocHeader = Trim(langInfo.SelectSingleNode("cartDocHeader").InnerText)
            _cartTOC = Trim(langInfo.SelectSingleNode("cartTOC").InnerText)
            _cartTOCDocInfo = Trim(langInfo.SelectSingleNode("cartTOCDocInfo").InnerText)
            _cartTOCProducts = Trim(langInfo.SelectSingleNode("cartTOCProducts").InnerText)
            _cartHeader = Trim(langInfo.SelectSingleNode("cartHeader").InnerText)
            _cartSubHeader = Trim(langInfo.SelectSingleNode("cartSubHeader").InnerText)
            _cartPageType = Trim(langInfo.SelectSingleNode("cartPageType").InnerText)
            _cartLastPage = Trim(langInfo.SelectSingleNode("cartLastPage").InnerText)
            _cartSafetyPage = Trim(langInfo.SelectSingleNode("cartSafetyPage").InnerText)
            _cartCopy = Trim(langInfo.SelectSingleNode("cartCopy").InnerText)
            _cartManufacturerHeader = Trim(langInfo.SelectSingleNode("cartManufacturerHeader").InnerText)
            _cartManufacturer = Trim(langInfo.SelectSingleNode("cartManufacturer").InnerText)
            _cartPreparedHeader = Trim(langInfo.SelectSingleNode("cartPreparedHeader").InnerText)
            _cartWarrantyHeader = Trim(langInfo.SelectSingleNode("cartWarrantyHeader").InnerText)
            _cartOrderingInfo = Trim(langInfo.SelectSingleNode("cartOrderingInfo").InnerText)
            _cartOrderProduct = Trim(langInfo.SelectSingleNode("cartOrderProduct").InnerText)
            _cartOrderDesc = Trim(langInfo.SelectSingleNode("cartOrderDesc").InnerText)
            _cartOrderWeight = Trim(langInfo.SelectSingleNode("cartOrderWeight").InnerText)
            _cartFrontCover = Trim(langInfo.SelectSingleNode("cartFrontCover").InnerText)
            _cartBackCover = Trim(langInfo.SelectSingleNode("cartBackCover").InnerText)
            _accWarrantyHeader = Trim(langInfo.SelectSingleNode("accWarrantyHeader").InnerText)
            _accWarranty1 = Trim(langInfo.SelectSingleNode("accWarranty1").InnerText)
            _formWarrantyHeader = Trim(langInfo.SelectSingleNode("formWarrantyHeader").InnerText)
            _formWarranty1 = Trim(langInfo.SelectSingleNode("formWarranty1").InnerText)
            _formPageType = Trim(langInfo.SelectSingleNode("formPageType").InnerText)
            _formManufacturerHeader = Trim(langInfo.SelectSingleNode("formManufacturerHeader").InnerText)
            _formManufacturer = Trim(langInfo.SelectSingleNode("formManufacturer").InnerText)
            _formRentalHeader = Trim(langInfo.SelectSingleNode("formRentalHeader").InnerText)
            _formRental = Trim(langInfo.SelectSingleNode("formRental").InnerText)

            _Loaded = True
        End If

    End Sub

End Class


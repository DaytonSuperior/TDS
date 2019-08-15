Imports Microsoft.VisualBasic
Imports System.Xml
Imports System.Collections.Generic
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO
Imports System.Linq

Public Class pdfHelper

#Region "Privates"
    Private _pdfContentByte As PdfContentByte
    Private _ct As ColumnText
#End Region


    ''' <summary>
    ''' Get/Set the content byte variable
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property pdfContentByte() As PdfContentByte
        Get
            Return _pdfContentByte
        End Get
        Set(ByVal value As PdfContentByte)
            _pdfContentByte = value
        End Set
    End Property

    ''' <summary>
    ''' Init a new instance
    ''' </summary>
    ''' <param name="pdfContentByte"></param>
    ''' <remarks></remarks>
    Public Sub New(pdfContentByte As PdfContentByte)
        _pdfContentByte = pdfContentByte
    End Sub

    ''' <summary>
    ''' add text to the pdfcontent
    ''' </summary>
    ''' <param name="text">The text to add</param>
    ''' <param name="font">The font (type and size)</param>
    ''' <param name="lowerLeftx">starting left x</param>
    ''' <param name="lowerLefty">starting left y</param>
    ''' <param name="upperRightx">ending right x</param>
    ''' <param name="upperRighty">ending right y</param>
    ''' <param name="leading">character leading</param>
    ''' <param name="alignment">text alignment</param>
    ''' <remarks></remarks>
    Public Sub PlaceText(text As String, font As iTextSharp.text.Font, _
                         lowerLeftx As Single, lowerLefty As Single, upperRightx As Single, upperRighty As Single, leading As Single, alignment As Integer)
        If _ct Is Nothing Then
            _ct = New ColumnText(_pdfContentByte)
        End If

        _ct.SetSimpleColumn(New Phrase(text, font), lowerLeftx, lowerLefty, upperRightx, upperRighty, leading, alignment)
        _ct.Go()
    End Sub

    Public Sub AddContent(ct As ColumnText, ele As iTextSharp.text.Element)
        ct.AddElement(ele)
    End Sub


End Class

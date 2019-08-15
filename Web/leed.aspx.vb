Imports System.Xml
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO

Partial Class leed
    Inherits System.Web.UI.Page

#Region "Privates"
    Dim Columns = {{48, 36, 564, 700}}

    Private checkHeight As Single = 15.0F
    Private checkWidth As Single = 15.0F

    Private leadingHdrF As Single = 0.0F
    Private leadingHdrM As Single = 2.0F
    Private spacingHdrAfter As Single = 4.0F
    Private leadingParaF As Single = 0.0F
    Private leadingParaM As Single = 1.1F
    Private leadingParaMHTML As Single = 1.1F
    Private spacingParaAfter As Single = 3.0F
    Private charspacingPara As Single = 0.2F
    Private tableHugeFont As Font = New Font(Font.FontFamily.HELVETICA, 17, Font.BOLD, BaseColor.BLACK)
    Private tableLargeFont As Font = New Font(Font.FontFamily.HELVETICA, 14, Font.BOLD, BaseColor.BLACK)
    Private tableHdrFont As Font = New Font(Font.FontFamily.HELVETICA, 11, Font.BOLD, BaseColor.BLACK)
    Private tableHdrFontSm As Font = New Font(Font.FontFamily.HELVETICA, 9, Font.BOLD, BaseColor.BLACK)
    Private tableFontSm As Font = New Font(Font.FontFamily.HELVETICA, 9, Font.NORMAL, BaseColor.BLACK)
    Private tableAttrHdrFont As Font = New Font(Font.FontFamily.HELVETICA, 10, Font.BOLD, BaseColor.WHITE)
    Private tableAttrFooterFont As Font = New Font(Font.FontFamily.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)
    Private tableLargeAttrFont As Font = New Font(Font.FontFamily.HELVETICA, 14, Font.NORMAL, BaseColor.BLACK)
    Private tableAttrFont As Font = New Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)
    Private tableLblFont As Font = New Font(Font.FontFamily.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)
    Private zapFont As Font = New Font(Font.FontFamily.ZAPFDINGBATS, 8.0F, Font.NORMAL)

    Private FontPath As String = Server.MapPath("resources")
    Private HumBold As BaseFont = BaseFont.CreateFont(FontPath + "\arialbd.ttf", BaseFont.CP1252, BaseFont.EMBEDDED)
    Private HumBoldExtra As BaseFont = BaseFont.CreateFont(FontPath + "\ariblk.ttf", BaseFont.CP1252, BaseFont.EMBEDDED)
    Private HumBoldItalic As BaseFont = BaseFont.CreateFont(FontPath + "\arialbi.ttf", BaseFont.CP1252, BaseFont.EMBEDDED)
    Private HumBoldLite As BaseFont = BaseFont.CreateFont(FontPath + "\ARIALNB.ttf", BaseFont.CP1252, BaseFont.EMBEDDED)
    Private HumReg As BaseFont = BaseFont.CreateFont(FontPath + "\arial.ttf", BaseFont.CP1252, BaseFont.EMBEDDED)


    Dim productFont As Font = FontFactory.GetFont("HumReg", 12, Font.BOLD, BaseColor.BLACK)
    Dim sectionFont As Font = FontFactory.GetFont("HumBold", 16, Font.BOLD, BaseColor.WHITE)
    Dim productTypeFont As Font = FontFactory.GetFont("HumReg", 11, Font.BOLDITALIC, BaseColor.BLACK)
    Dim sectionTypeFont As Font = FontFactory.GetFont("HumReg", 8, Font.BOLDITALIC, BaseColor.GRAY)
    Dim sheetTypeFont As Font = FontFactory.GetFont("HumReg", 12, Font.BOLD, BaseColor.GRAY)
    Dim AttrHdrFont As Font = FontFactory.GetFont("HumBold", 11, Font.BOLD, BaseColor.BLACK)
    Dim AttrAnswerFont As Font = FontFactory.GetFont("HumReg", 11, Font.NORMAL, BaseColor.BLACK)
    Dim AttrFont As Font = FontFactory.GetFont("HumReg", 9, Font.ITALIC, BaseColor.BLACK)
    Dim smAttrHdrFont As Font = FontFactory.GetFont("HumBold", 9, Font.BOLD, BaseColor.BLACK)
    Dim smAttrFont As Font = FontFactory.GetFont("HumReg", 8, Font.NORMAL, BaseColor.BLACK)
    Dim smAttrNameFont As Font = FontFactory.GetFont("HumReg", 8, Font.BOLD, BaseColor.BLACK)

    Private noColor As BaseColor = Nothing
    Private bColor As BaseColor = iTextSharp.text.BaseColor.WHITE
    Private altColor As BaseColor = iTextSharp.text.BaseColor.LIGHT_GRAY
    Private hdrColor As BaseColor = New BaseColor(0, 148, 208)
    Private footerColor As BaseColor = New BaseColor(149, 224, 255)
    Private ColorDifficult As BaseColor = iTextSharp.text.BaseColor.RED
    Private ColorEasy As BaseColor = iTextSharp.text.BaseColor.GREEN
    Private ColorHard As BaseColor = iTextSharp.text.BaseColor.YELLOW
    Private cellb As Single = 47.0F
    Private cellm As Single = 30.0F
    Private cells As Single = 20.0F
    Private cellt As Single = 5.0F

    Private _ProductID As Long = 0
    Private _ProductName As String = ""
    Private _PMName As String = ""
    Private _PMSignature As String = ""
    Private _LetterDate As String = ""
    Private _SiteCount As Integer = 0
    Private _CreditCount As Integer = 0
#End Region

    Private Sub LeedProductPDF_Init(sender As Object, e As EventArgs) Handles Me.Init

    End Sub

    Private Sub LeedProductPDF_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Request("pid") <> "" AndAlso IsNumeric(Request("pid")) Then
            GenPDF(Request("pid"))
        Else
            GenPDF(1)
        End If
    End Sub

    Private Sub GenPDF(ByVal ProductID As Long)

        Dim webFolder As String = System.Configuration.ConfigurationManager.AppSettings("WebPath")
        Dim books = XDocument.Load(webFolder & System.Configuration.ConfigurationManager.AppSettings("XML_LeedFile"))

        Dim letterXML As XmlDocument = New XmlDocument()
        letterXML.LoadXml(books.ToString)

        'Dim LetterRecord As XmlNodeList = letterXML.SelectNodes("/Letters/Letter[@id='" & ProductID.ToString & "']")
        Dim LetterRecord As XmlNodeList = letterXML.SelectNodes("descendant::Letter[id='" & ProductID & "']")
        If LetterRecord.Count = 1 Then
            Dim Letter As XmlNode = LetterRecord(0)

            _ProductName = Trim(Letter.SelectSingleNode("ProductName").InnerText)
            _LetterDate = Trim(Letter.SelectSingleNode("LetterDate").InnerText)
            _PMName = Trim(Letter.SelectSingleNode("Signing/Manager").InnerText)
            _PMSignature = Trim(Letter.SelectSingleNode("Signing/ImageFile").InnerText)
            _SiteCount = Letter.SelectNodes("Manufacturing/Location").Count
            _CreditCount = Letter.SelectNodes("Credits/Credit").Count

            _ProductID = ProductID
            Dim success As Boolean = False

            'Response.ContentType = "application/vnd.adobe.xfdf"
            Dim ws As New XmlWriterSettings
            ws.CheckCharacters = True
            ws.CloseOutput = True
            ws.Indent = True

            Dim g As Guid
            g = Guid.NewGuid()

            Dim aFileName As String = "LeedLetter_" & g.ToString & ".pdf"
            Dim DestPDF = Server.MapPath("/pdfs/" & aFileName)

            'Delete the dest PDF if it exists
            If System.IO.File.Exists(DestPDF) = True Then
                Try
                    System.IO.File.Delete(DestPDF)
                Catch
                End Try
            End If

            Dim imagepath As String = Server.MapPath("/images/brand1.png")
            Dim logox As Single = 24
            Dim logoy As Single = 735
            Dim logoh As Single = 44
            Dim logow As Single = 158

            Dim txtnote As String = ""

            Dim notex1 As Single = 48
            Dim notex2 As Single = 588
            Dim notey1 As Single = 20
            Dim notey2 As Single = 40

            Dim tablew As Single = 540.0F
            Dim cellIndent1 As Single = 15.0F
            Dim cellIndent2 As Single = 10.0F
            Dim cellw As Single = 50.0F
            Dim cellFinal As Single = 75.0F

            Dim doc As New Document

            Dim txtheader As String = "LEED Information: " & _ProductName

            Dim headerx1 As Single = 48
            Dim headerx2 As Single = 588
            Dim headery1 As Single = 730
            Dim headery2 As Single = 770

            'Dim graph1x As Single = 75
            'Dim graph1y As Single = -615
            'Dim graph1h As Single = 925
            'Dim graph1w As Single = 1300
            Dim graph1x As Single = 50
            Dim graph1y As Single = 325
            Dim graph1h As Single = 350
            Dim graph1w As Single = 600

            Dim bgx As Single = 48
            Dim bgy As Single = 48
            Dim bgh As Single = 287
            Dim bgw As Single = 500


            Dim writer As PdfWriter

            Try
                writer = PdfWriter.GetInstance(doc, New FileStream(DestPDF, FileMode.Create))

                Dim rect As New Rectangle(0, 0, 612, 792)
                doc.SetPageSize(rect)
                doc.SetMargins(48, -92, 65, 24)
                doc.Open()

                doc.NewPage()
                Dim ct As New ColumnText(writer.DirectContent)
                Dim column As Integer = 0
                ct.SetSimpleColumn(Columns(column, 0), Columns(column, 1), Columns(column, 2), Columns(column, 3))
                Dim status As Integer = 0
                Dim YLine As Double = ct.YLine

                Dim aTable As PdfPTable
                Dim hlCell As PdfPCell
                Dim widths As Single() = New Single() {cellIndent1, cellIndent2, cellw, cellw, cellw, cellw, 60.0F, 60.0F, cellw, cellFinal}

                ' Adding the empty page for the graph
                doc.NewPage()
                aTable = New PdfPTable(10)
                aTable.HorizontalAlignment = 0
                aTable.TotalWidth = 520.0F
                aTable.LockedWidth = True
                aTable.SetWidths(widths)

                ' SPACE DOWN FROM THE TOP - This assumes a total 14 items possible between the manufacturing locations and LEED options
                ' adjust this number as necessary
                Dim ItemMax As Integer = 14
                Dim SpacingCount As Integer = 0
                Dim LocationTotal As Integer = 0
                Dim CreditTotal As Integer = 0
                SpacingCount = (_SiteCount + _CreditCount)

                If SpacingCount < 11 Then
                    SpacingCount = 11
                End If

                While SpacingCount < ItemMax
                    hlCell = New PdfPCell(New Phrase("", tableAttrFont))
                    hlCell.Colspan = 10
                    hlCell.Border = 0
                    AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_MIDDLE)
                    SpacingCount += 1
                End While

                ' Add date
                Try
                    hlCell = New PdfPCell(New Phrase("Date: " & CDate(_LetterDate).ToShortDateString, tableAttrFont))
                Catch
                    hlCell = New PdfPCell(New Phrase("Date: " & Now().ToShortDateString, tableAttrFont))
                End Try
                hlCell.Colspan = 10
                hlCell.Border = 0
                AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_MIDDLE)

                hlCell = New PdfPCell(New Phrase("", tableAttrFont))
                hlCell.Colspan = 10
                hlCell.Border = 0
                AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_MIDDLE)

                hlCell = New PdfPCell(New Phrase("To whom it may concern:", tableAttrFont))
                hlCell.Colspan = 10
                hlCell.Border = 0
                AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_MIDDLE)

                hlCell = New PdfPCell(New Phrase("The following product may qualify for LEED credits:", tableAttrFont))
                hlCell.Colspan = 10
                hlCell.Border = 0
                AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_MIDDLE)

                hlCell = New PdfPCell(New Phrase("Product Name: " & _ProductName, tableAttrFont))
                hlCell.Colspan = 10
                hlCell.Border = 0
                AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_MIDDLE)

                hlCell = New PdfPCell(New Phrase("Manufacturing Locations:", tableLblFont))
                hlCell.Colspan = 10
                hlCell.Border = 0
                AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_MIDDLE)

                Dim manList As XmlNodeList = Letter.SelectNodes("Manufacturing/Location")
                For Each xItem As XmlNode In manList

                    hlCell = New PdfPCell(New Phrase("", tableAttrFont))
                    hlCell.Colspan = 1
                    hlCell.Border = 0
                    AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_TOP)

                    hlCell = New PdfPCell(New Phrase("n", zapFont))
                    hlCell.Colspan = 1
                    hlCell.Border = 0
                    AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_TOP)

                    hlCell = New PdfPCell(New Phrase(Trim(xItem.InnerText).Replace(vbCrLf, ""), tableAttrFont))
                    hlCell.Colspan = 8
                    hlCell.Border = 0
                    AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_TOP)
                Next

                hlCell = New PdfPCell(New Phrase("LEED Credits:", tableLblFont))
                hlCell.Colspan = 10
                hlCell.Border = 0
                AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_MIDDLE)

                Dim creditList As XmlNodeList = Letter.SelectNodes("Credits/Credit")
                For Each xItem As XmlNode In creditList
                    hlCell = New PdfPCell(New Phrase("", tableAttrFont))
                    hlCell.Colspan = 1
                    hlCell.Border = 0
                    AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_TOP)

                    hlCell = New PdfPCell(New Phrase("n", zapFont))
                    hlCell.Colspan = 1
                    hlCell.Border = 0
                    AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_TOP)

                    hlCell = New PdfPCell(New Phrase(Trim(xItem.InnerText).Replace(vbCrLf, ""), tableAttrFont))
                    hlCell.Colspan = 8
                    hlCell.Border = 0
                    AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_TOP)
                Next

                hlCell = New PdfPCell(New Phrase("", tableAttrFont))
                hlCell.Colspan = 1
                hlCell.Border = 0
                AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_TOP)

                hlCell = New PdfPCell(New Phrase("n", zapFont))
                hlCell.Colspan = 1
                hlCell.Border = 0
                AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_TOP)

                hlCell = New PdfPCell(New Phrase("MR Credit 5.1 & MR Credit 5.2 are based on the distance between the Manufacturing facility and the job site. These credits are by default a possibility for every product, but the manufacturing location(s) will help in determining if the credit applies on a job-by-job basis.", tableAttrFont))
                hlCell.Colspan = 8
                hlCell.Border = 0
                AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_TOP)

                hlCell = New PdfPCell(New Phrase("", tableAttrFont))
                hlCell.Colspan = 10
                hlCell.Border = 0
                AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_TOP)

                hlCell = New PdfPCell(New Phrase("For additional information or specific questions (including manufacturing locations) on LEED credits for our products please contact Technical Services Toll Free at: 1-800-555-1212", tableAttrFont))
                hlCell.Colspan = 10
                hlCell.Border = 0
                AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_MIDDLE)

                hlCell = New PdfPCell(New Phrase("", tableAttrFont))
                hlCell.Colspan = 10
                hlCell.Border = 0
                AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_TOP)

                hlCell = New PdfPCell(New Phrase("Best Regards,", tableAttrFont))
                hlCell.Colspan = 10
                hlCell.Border = 0
                AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_MIDDLE)

                Try
                    Dim SigFile As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(webFolder & "signatures\" & _PMSignature)
                    Dim picw As Integer = 150
                    Dim pich As Integer = 150
                    Dim whRatio As Double
                    If SigFile.Width > 150 Then
                        picw = 150
                    End If

                    whRatio = 150 / SigFile.Width
                    pich = SigFile.Height * whRatio

                    SigFile.ScaleAbsoluteWidth(picw)
                    SigFile.ScaleAbsoluteHeight(pich)

                    Dim imageCell As PdfPCell = New PdfPCell(SigFile)
                    imageCell.Colspan = 4
                    imageCell.Border = 0
                    '                imageCell.setHorizontalAlignment(Element.ALIGN_CENTER)
                    AddCell(aTable, imageCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_MIDDLE)

                    hlCell = New PdfPCell(New Phrase("", tableAttrFont))
                    hlCell.Colspan = 6
                    hlCell.Border = 0
                    AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_MIDDLE)
                Catch
                End Try

                hlCell = New PdfPCell(New Phrase(_PMName, tableAttrFont))
                hlCell.Colspan = 10
                hlCell.Border = 0
                AddCell(aTable, hlCell, bColor, cells, Element.ALIGN_LEFT, Element.ALIGN_MIDDLE)

                doc.Add(aTable)

                success = True
            Catch ex As Exception
                'MsgBox(ex.Message)
                success = False
            Finally
                doc.Close()
                writer.Close()
            End Try

            ' Stamp logos!
            ' keep track of the current page
            Dim CurrentPageCtr As Integer = 1

            ' create a reader for a certain document
            Dim readerQQ As PdfReader = New PdfReader(DestPDF)
            Dim n As Integer = readerQQ.NumberOfPages

            Dim document As Document = New Document(readerQQ.GetPageSizeWithRotation(1))

            ' custom file name for testing
            g = Guid.NewGuid()
            aFileName = "LEED_Letter_" & DatePart(DateInterval.WeekOfYear, Now) & "_" & Now.Year & "_" & g.ToString & ".pdf"

            'Delete the dest PDF if it exists
            If System.IO.File.Exists(webFolder & "pdfs\" & Trim(aFileName)) = True Then
                Try
                    System.IO.File.Delete(webFolder & "pdfs\" & Trim(aFileName))
                Catch
                End Try
            End If

            writer = PdfWriter.GetInstance(document, New FileStream(webFolder & "pdfs\" & Trim(aFileName), FileMode.Create))
            document.Open()

            Dim cb As PdfContentByte = writer.DirectContent
            Dim PdfHelper As New pdfHelper(cb)

            Dim page As PdfImportedPage
            Dim rotation As Integer

            ' add content
            Dim i As Integer = 0

            Do While (i < n)
                i += 1

                document.SetPageSize(readerQQ.GetPageSizeWithRotation(i))
                document.NewPage()
                document.NewPage()
                page = writer.GetImportedPage(readerQQ, i)
                rotation = readerQQ.GetPageRotation(i)
                If (rotation = 90 Or rotation = 270) Then
                    cb.AddTemplate(page, 0, -1.0F, 1.0F, 0, 0, readerQQ.GetPageSizeWithRotation(i).Height)
                Else
                    cb.AddTemplate(page, 1.0F, 0, 0, 1.0F, 0, 0)
                End If

                ' add the logo
                Dim logoimg As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(imagepath)
                logoimg.SetAbsolutePosition(logox, logoy)
                logoimg.ScaleAbsoluteWidth(logow)
                logoimg.ScaleAbsoluteHeight(logoh)
                cb.AddImage(logoimg)


                ' add the black line
                cb.SetLineWidth(0.75F)
                cb.SetGrayStroke(0.0F)

                cb.MoveTo(logoimg.ScaledWidth + 20.0F, logoy + 17.0F)
                cb.LineTo(logoimg.ScaledWidth + 20.0F + (headerx2 - (logoimg.ScaledWidth + 20.0F)), logoy + 17.0F)

                cb.Stroke()

                ' reset color
                cb.SetColorFill(New CMYKColor(0.0F, 0.0F, 0.0F, 1.0F))

                ' add note
                PdfHelper.PlaceText(txtnote, smAttrFont, notex1, notey1, notex2, notey2, 14, Element.ALIGN_RIGHT)

                ' add header and type
                PdfHelper.PlaceText(txtheader, productFont, headerx1, headery1, headerx2, headery2, 14, Element.ALIGN_RIGHT)
                'PdfHelper.PlaceText(txtProductType, productTypeFont, headerx1, headery1, headerx2, headery2 - 15, 14, Element.ALIGN_RIGHT)
                'PdfHelper.PlaceText(txtPageType, sheetTypeFont, 48.0F, logoy - logoimg.ScaledHeight - 25.0F, headerx2, logoy - 10.0F, 14, Element.ALIGN_LEFT)

                CurrentPageCtr += 1
            Loop

            Try
                document.Close()
            Catch ex As Exception
                Response.Write(ex.Message)
                Response.End()
                success = False
            End Try

            Try
                readerQQ.Close()
            Catch
            End Try

            'If success Then
            ' at this point generate the PDF if we haven't detected a change and a valid user
            'Response.Clear()
            'Response.Redirect("./pdfs/" & aFileName)

            'DestPDF = HttpContext.Current.Server.MapPath("/pdfs/" & aFileName)

            'Response.Clear()
            'Response.Redirect("./pdfs/" & aFileName)
        End If

    End Sub

    Protected Sub WriteEmptyCells(ByRef aTable As PdfPTable, ByVal num As Integer, Optional ByVal BorderTop As Boolean = False, Optional ByVal BorderBottom As Boolean = False, Optional ByVal BorderLeft As Boolean = False, Optional ByVal BorderRight As Boolean = False)
        For counter As Integer = 1 To num Step 1
            Dim hlCell As PdfPCell = New PdfPCell(New Phrase("", AttrFont))
            hlCell.BorderWidth = 0.25

            If Not BorderBottom Then
                hlCell.BorderWidthBottom = 0
            End If
            If Not BorderTop Then
                hlCell.BorderWidthTop = 0
            End If
            If Not BorderLeft Then
                hlCell.BorderWidthLeft = 0
            End If
            If Not BorderRight Then
                hlCell.BorderWidthRight = 0
            End If
            AddCell(aTable, hlCell, noColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE, BorderTop, BorderBottom, BorderLeft, BorderRight)
        Next
    End Sub

    Private Sub AddCell(ByRef aTable As PdfPTable, ByRef hlCell As PdfPCell, bColor As BaseColor, cellHeight As Single, cellHAlign As Integer, cellVAlign As Integer, Optional ByVal BorderTop As Boolean = True, Optional ByVal BorderBottom As Boolean = True, Optional ByVal BorderLeft As Boolean = True, Optional ByVal BorderRight As Boolean = True)
        hlCell.BorderWidth = 0.25

        If Not BorderBottom Then
            hlCell.BorderWidthBottom = 0
        End If
        If Not BorderTop Then
            hlCell.BorderWidthTop = 0
        End If
        If Not BorderLeft Then
            hlCell.BorderWidthLeft = 0
        End If
        If Not BorderRight Then
            hlCell.BorderWidthRight = 0
        End If

        hlCell.MinimumHeight = cellHeight
        hlCell.BackgroundColor = bColor
        hlCell.HorizontalAlignment = cellHAlign
        hlCell.VerticalAlignment = cellVAlign
        aTable.AddCell(hlCell)
    End Sub
End Class


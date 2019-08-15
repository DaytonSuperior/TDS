Imports System.Xml
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO

Partial Class tds_forming
    Inherits System.Web.UI.Page

#Region "Privates"
    Private _ProductID As Long

    Dim Columns = {{48, 36, 288, 700}, {312, 36, 564, 700}}

    Private leadingHdrF As Single = 0.0F
    Private leadingHdrM As Single = 2.0F
    Private spacingHdrAfter As Single = 4.0F
    Private leadingParaF As Single = 0.0F
    Private leadingParaM As Single = 1.1F
    Private spacingParaAfter As Single = 3.0F
    Private charspacingPara As Single = 0.2F

    Private cells As Single = 16.0F
    Private bColor As BaseColor = iTextSharp.text.BaseColor.WHITE
    Private fColor As BaseColor = iTextSharp.text.BaseColor.BLACK
    Private gColor As BaseColor = iTextSharp.text.BaseColor.GRAY
    Private hdrbColor As BaseColor = New iTextSharp.text.BaseColor(System.Drawing.Color.FromArgb(0, 148, 208))
    Private hdrfColor As BaseColor = iTextSharp.text.BaseColor.WHITE

    Private FontPath As String = Server.MapPath("/Resources")
    Private HumBold As BaseFont = BaseFont.CreateFont(FontPath + "\arialbd.ttf", BaseFont.CP1252, BaseFont.EMBEDDED)
    Private HumBoldExtra As BaseFont = BaseFont.CreateFont(FontPath + "\ariblk.ttf", BaseFont.CP1252, BaseFont.EMBEDDED)
    Private HumBoldItalic As BaseFont = BaseFont.CreateFont(FontPath + "\arialbi.ttf", BaseFont.CP1252, BaseFont.EMBEDDED)
    Private HumBoldLite As BaseFont = BaseFont.CreateFont(FontPath + "\ARIALNB.ttf", BaseFont.CP1252, BaseFont.EMBEDDED)
    Private HumReg As BaseFont = BaseFont.CreateFont(FontPath + "\arial.ttf", BaseFont.CP1252, BaseFont.EMBEDDED)

    Private tableAttrFont10 As Font = New Font(HumReg, 10, Font.NORMAL, BaseColor.BLACK)
    Private tableAttrFont9 As Font = New Font(HumReg, 9, Font.NORMAL, BaseColor.BLACK)
    Private tableAttrFont8 As Font = New Font(HumReg, 8, Font.NORMAL, BaseColor.BLACK)
    Private tableAttrFont7 As Font = New Font(HumReg, 7, Font.NORMAL, BaseColor.BLACK)
    Private tableAttrFont6 As Font = New Font(HumReg, 6, Font.NORMAL, BaseColor.BLACK)
    Private tableAttrFont10white As Font = New Font(HumReg, 10, Font.NORMAL, BaseColor.WHITE)
    Private tableAttrFont9white As Font = New Font(HumReg, 9, Font.NORMAL, BaseColor.WHITE)
    Private tableAttrFont8white As Font = New Font(HumReg, 8, Font.NORMAL, BaseColor.WHITE)
    Private tableAttrFont7white As Font = New Font(HumReg, 7, Font.NORMAL, BaseColor.WHITE)
    Private tableAttrFont6white As Font = New Font(HumReg, 6, Font.NORMAL, BaseColor.WHITE)

    Private _ProductName As String = ""

    Private _Language As String = ""
    Private _BrandID As Integer
    Private _Brand As String = ""
    Private _SharepointID As String = ""
    Private _SitefinityID As String = ""
    Private _ArtifactName As String = ""
    Private _PricebookCode As String = ""
    Private _ManagedBy As String = ""
    Private _Updated As String = ""

#End Region

    Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init
        If Request("pid") <> "" AndAlso IsNumeric(Request("pid")) Then
            _ProductID = Request("pid")
        Else
            _ProductID = 1
        End If

        GenPDF(_ProductID)
    End Sub

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load

    End Sub

    Protected Sub GenPDF(ByVal ProductID As Long)

        ' Top header - above the line, name of product
        Dim productFont As Font = New Font(HumReg, 18, Font.BOLD, BaseColor.BLACK) ' was HumBold and FONT.BOLDITALIC
        ' Product Type - just below the top line, under product name
        Dim productTypeFont As Font = New Font(HumReg, 11, Font.BOLD, BaseColor.GRAY) ' was FONT.BOLDITALIC
        ' Type of document - just below logo i.e. "TECHNICAL DATA SHEET"
        Dim sheetTypeFont As Font = New Font(HumReg, 12, Font.BOLD, BaseColor.GRAY)
        ' Headers in the document (usually upper case)
        Dim AttrHdrFont As Font = New Font(HumBold, 11, Font.NORMAL, BaseColor.BLACK) ' was FONT.BOLD
        ' Document Text (under a header)
        Dim AttrFont As Font = New Font(HumReg, 10, Font.NORMAL, BaseColor.BLACK)
        ' Sub section headers (used in chemicals only right now, usually lower case)
        Dim smAttrHdrFont As Font = New Font(HumBold, 11, Font.NORMAL, BaseColor.BLACK) ' was FONT.BOLD
        ' Sub section text (under a sub header)
        Dim smAttrFont As Font = New Font(HumReg, 10, Font.NORMAL, BaseColor.BLACK)
        ' Next 2 are for the price book lookoup tables
        Dim tableAttrHdrFont As Font = New Font(HumReg, 8, Font.BOLD, BaseColor.BLACK) ' was HumBold
        Dim tableAttrFont As Font = New Font(HumReg, 7, Font.NORMAL, BaseColor.BLACK)
        ' Next 2 -> Sub section: Notes (used in chemicals only right now, slightly smaller than other headers/subheaders)
        Dim NoteHdrFont As Font = New Font(HumReg, 10, Font.BOLD, BaseColor.BLACK) ' was HumBold
        Dim NoteFont As Font = New Font(HumReg, 9, Font.NORMAL, BaseColor.BLACK)
        Dim zapFont As Font = New Font(Font.FontFamily.ZAPFDINGBATS, 10.0F, Font.NORMAL)

        Dim headerx1 As Single = 48
        Dim headerx2 As Single = 588
        Dim headery1 As Single = 730
        Dim headery2 As Single = 770

        Dim footerx1 As Single = 24
        Dim footerx2 As Single = 588
        Dim footery1 As Single = 6
        Dim footery2 As Single = 26

        Dim webFolder As String = System.Configuration.ConfigurationManager.AppSettings("WebPath")
        Dim books = XDocument.Load(webFolder & System.Configuration.ConfigurationManager.AppSettings("XML_FormingFile"))

        Dim tdsXML As XmlDocument = New XmlDocument()
        tdsXML.LoadXml(books.ToString)

        'Dim LetterRecord As XmlNodeList = tdsXML.SelectNodes("/sheets/tds[@id='" & ProductID & "']")
        Dim LetterRecord As XmlNodeList = tdsXML.SelectNodes("descendant::tds[id='" & ProductID & "']")
        If LetterRecord.Count = 1 Then
            Dim Letter As XmlNode = LetterRecord(0)

            _ProductID = ProductID
            _ProductName = Trim(Letter.SelectSingleNode("productname").InnerText)
            _Language = Trim(Letter.SelectSingleNode("language").InnerText)
            _BrandID = Trim(Letter.SelectSingleNode("brandid").InnerText)
            _Brand = Trim(Letter.SelectSingleNode("productbrand").InnerText)
            _SharepointID = Trim(Letter.SelectSingleNode("sharepointid").InnerText)
            _SitefinityID = Trim(Letter.SelectSingleNode("sitefinityid").InnerText)
            _ArtifactName = Trim(Letter.SelectSingleNode("artifactname").InnerText)
            _PricebookCode = Trim(Letter.SelectSingleNode("pricebookcode").InnerText)
            _ManagedBy = Trim(Letter.SelectSingleNode("managedby").InnerText)
            _Updated = Trim(Letter.SelectSingleNode("tsudpated").InnerText)

            Response.ContentType = "application/vnd.adobe.xfdf"
            Dim ws As New XmlWriterSettings()
            ws.CheckCharacters = True
            ws.CloseOutput = True
            ws.Indent = True

            Dim Lang As String = _Language
            Dim BrandID As Integer = _BrandID

            Dim tdsLangInfo As New languageInfo(BrandID, Lang)
            Dim tdsBrandInfo As New brandInfo(BrandID)

            If tdsLangInfo IsNot Nothing And tdsBrandInfo IsNot Nothing Then
            Else
                Response.Write("Could not find language information")
                Response.End()
            End If

            Dim txtHeader As String = _ProductName
            Dim txtPageType As String = tdsLangInfo.formPageType
            Dim txtProductType As String = ""
            Dim txtSection As String = ""
            Dim txtCategory As String = ""
            Dim txtFooter1 As String = tdsLangInfo.footer1
            Dim txtFooter2 As String = tdsLangInfo.footer2
            Dim txtFooter3 As String = tdsLangInfo.footer3
            Dim txtFooter4 As String = tdsLangInfo.footer4

            Dim bGuid As Guid
            bGuid = Guid.NewGuid()

            Dim DestPDF = Server.MapPath("/pdfs/tds_forming_" & bGuid.ToString & ".pdf")

            'Delete the dest PDF if it exists
            If System.IO.File.Exists(DestPDF) = True Then
                Try
                    System.IO.File.Delete(DestPDF)
                Catch
                End Try
            End If

            ' step 4 - open 3 PDFs (filePath (Lease page 1), BOM page(s), TaCPath (T&C pages)) + merge
            ' will increment this as we need to for more bom pages
            Dim BOMPage As Integer = 1

            ' default total
            Dim TotalPages As Integer = 0
            ' inc this for each extra BOM page

            ' Create BOM PDF Page(s)

            Dim imagepath As String = Server.MapPath("/images/" & tdsBrandInfo.Logo)
            hdrbColor = New iTextSharp.text.BaseColor(System.Drawing.Color.FromArgb(tdsBrandInfo.colorR, tdsBrandInfo.colorG, tdsBrandInfo.colorB))
            Dim logox As Single = 24
            Dim logoy As Single = 735
            Dim logoh As Single = 44
            Dim logow As Single = 158

            Dim tableh As Single = 53 '88 x 400
            Dim tablew As Single = 240

            Dim doc As New Document

            Dim writer As PdfWriter

            Try
                writer = PdfWriter.GetInstance(doc, New FileStream(DestPDF, FileMode.Create))

                Dim rect As New Rectangle(0, 0, 612, 792)
                doc.SetPageSize(rect)
                doc.SetMargins(48, 48, 90, 24)
                doc.Open()

                Dim ct As New ColumnText(writer.DirectContent)
                Dim column As Integer = 0
                ct.SetSimpleColumn(Columns(column, 0), Columns(column, 1), Columns(column, 2), Columns(column, 3))
                Dim status As Integer = 0
                Dim YLine As Double

                Dim LastHeader As String = ""
                Dim LastHeaderID As String = ""
                Dim LastHeaderWritten As Boolean = True

                Dim attrList As XmlNodeList = Letter.SelectNodes("attributes/attribute")
                For Each attr As XmlNode In attrList

                    Dim _attr_name As String = Trim(attr("name").InnerText)
                    Dim _attr_type As String = Trim(attr("type").InnerText)
                    Dim _attr_header As String = Trim(attr("header").InnerText)
                    Dim _attr_grid As String = Trim(attr("grid").InnerText)
                    Dim _attr_value As String = Trim(attr("value").InnerText)
                    Dim _attr_showonsheet As Boolean = CBool(Trim(attr("showonsheet").InnerText))

                    If ((_attr_value.ToString.Length > 1) Or IsNumeric(_attr_value)) Or (_attr_type = "12") Then
                        Select Case _attr_type
                            Case "0"
                            ' special case, product name, not used
                            Case "1"
                                ' product type
                                txtProductType = _attr_value
                            Case "2"
                                ' section
                                txtSection = _attr_value
                            Case "7"
                                ' category
                                txtCategory = _attr_value
                            Case "3"
                                ' header with text

                                If LastHeaderWritten = False AndAlso LastHeaderID = _attr_header Then
                                    ' need to write the header
                                    If _attr_showonsheet = True Then
                                        WriteHeadedParagraph(doc, ct, column, _attr_name, AttrHdrFont, _attr_value, AttrFont, False, LastHeader, AttrHdrFont)
                                        LastHeaderWritten = True
                                    End If
                                Else
                                    If _attr_showonsheet = True Then
                                        WriteHeadedParagraph(doc, ct, column, _attr_name, AttrHdrFont, _attr_value, AttrFont, False)
                                    End If
                                End If

                            Case "8"
                                ' note with text
                                If LastHeaderWritten = False AndAlso LastHeaderID = _attr_header Then
                                    ' need to write the header
                                    If _attr_showonsheet = True Then
                                        WriteHeadedParagraph(doc, ct, column, _attr_name, NoteHdrFont, _attr_value, NoteFont, True, LastHeader, AttrHdrFont)
                                        LastHeaderWritten = True
                                    End If
                                Else
                                    If _attr_showonsheet = True Then
                                        WriteHeadedParagraph(doc, ct, column, _attr_name, NoteHdrFont, _attr_value, NoteFont, True)
                                    End If
                                End If

                            Case "9"
                                ' subheader with text
                                If LastHeaderWritten = False AndAlso LastHeaderID = _attr_header Then
                                    ' need to write the header
                                    If _attr_showonsheet = True Then
                                        WriteHeadedParagraph(doc, ct, column, _attr_name, smAttrHdrFont, _attr_value, smAttrFont, True, LastHeader, AttrHdrFont)
                                        LastHeaderWritten = True
                                    End If
                                Else
                                    If _attr_showonsheet = True Then
                                        WriteHeadedParagraph(doc, ct, column, _attr_name, smAttrHdrFont, _attr_value, smAttrFont, True)
                                    End If
                                End If

                            Case "11"
                                'product Image
                                If _attr_showonsheet = True Then
                                    Try

                                        Dim img As System.Drawing.Image
                                        img = System.Drawing.Image.FromFile(webFolder & "images\tds\" & _attr_value)

                                        Dim bc As iTextSharp.text.BaseColor = New iTextSharp.text.BaseColor(0, 0, 0)
                                        Dim pic As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(img, bc)

                                        Dim picw As Integer = tablew
                                        Dim pich As Integer = tableh
                                        Dim whRatio As Double
                                        If pic.Width > tablew Then
                                            picw = tablew
                                        End If

                                        whRatio = tablew / pic.Width
                                        pich = pic.Height * whRatio

                                        pic.ScaleAbsoluteWidth(picw)
                                        pic.ScaleAbsoluteHeight(pich)

                                        YLine = ct.YLine
                                        AddImageNoHeader(ct, pic)
                                        CheckImageColumns(status, ct, column, doc, YLine, pich)
                                        AddImageNoHeader(ct, pic)
                                        status = ct.Go()

                                    Catch ex As Exception
                                        'Response.Write(ex.Message)
                                    End Try
                                End If
                            Case "12"
                                ' Technical data

                                Dim GridID As Integer = 0
                                Dim GridXMLName As String = ""
                                Select Case _attr_grid
                                    Case "Grid_TechData"
                                        GridID = 1
                                        GridXMLName = "gridTechData"
                                    Case "Grid_Install"
                                        GridID = 2
                                        GridXMLName = "gridInstall"
                                    Case Else
                                        GridID = 0
                                End Select

                                If GridID > 0 Then

                                    Dim gridInfo As XmlNodeList = Letter.SelectNodes(GridXMLName)
                                    For Each grid As XmlNode In gridInfo

                                        Dim gridRows As XmlNodeList = Letter.SelectNodes(GridXMLName & "/rows/row")
                                        If gridRows.Count > 0 Then

                                            Dim _grid_header As String = Trim(grid("header").InnerText)
                                            Dim _grid_desc As String = Trim(grid("description").InnerText)
                                            Dim _grid_cols As Integer = CInt(Trim(grid("colCount").InnerText))


                                            Dim tStatus As Integer
                                            CheckTableColumns(tStatus, ct, column, doc, YLine, gridRows)

                                            If _grid_header <> "" Then
                                                Dim attrHdrParaT As Paragraph = New Paragraph
                                                attrHdrParaT.SetLeading(leadingHdrF, leadingHdrM)
                                                attrHdrParaT.SpacingAfter = spacingHdrAfter
                                                Dim AttrHdrChunkT As Chunk = New Chunk(_grid_header, AttrHdrFont)
                                                Dim AttrHdrPhraseT As Phrase = New Phrase(AttrHdrChunkT)
                                                attrHdrParaT.Add(AttrHdrPhraseT)

                                                If _grid_desc <> "" Then
                                                    Dim attrParaT As Paragraph = New Paragraph
                                                    attrParaT.SetLeading(leadingParaF, leadingParaM)
                                                    attrParaT.SpacingAfter = spacingParaAfter
                                                    Dim AttrChunkT As Chunk = New Chunk(_grid_desc, AttrFont)
                                                    AttrChunkT.SetCharacterSpacing(charspacingPara)
                                                    Dim AttrPhraseT As Phrase = New Phrase(AttrChunkT)
                                                    attrParaT.Add(AttrPhraseT)
                                                    AddParagraph(ct, attrHdrParaT, attrParaT)
                                                    status = ct.Go()
                                                    YLine = ct.YLine
                                                Else
                                                    AddSingleParagraph(ct, attrHdrParaT)
                                                    status = ct.Go()
                                                    YLine = ct.YLine
                                                End If

                                            End If

                                            WriteTable(doc, ct, column, YLine, gridRows, _grid_cols)
                                        End If
                                    Next

                                Else
                                    If _attr_showonsheet = True Then
                                        Try
                                            Dim img As System.Drawing.Image
                                            img = System.Drawing.Image.FromFile(webFolder & "images\tds\" & _attr_value)

                                            Dim bc As iTextSharp.text.BaseColor = New iTextSharp.text.BaseColor(0, 0, 0)
                                            Dim pic As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(img, bc)

                                            Dim picw As Integer = tablew
                                            Dim pich As Integer = tableh
                                            Dim whRatio As Double
                                            If pic.Width > tablew Then
                                                picw = tablew
                                            End If

                                            whRatio = tablew / pic.Width
                                            pich = pic.Height * whRatio

                                            pic.ScaleAbsoluteWidth(picw)
                                            pic.ScaleAbsoluteHeight(pich)

                                            YLine = ct.YLine
                                            AddImageNoHeader(ct, pic)
                                            CheckImageColumns(status, ct, column, doc, YLine, pich)
                                            AddImageNoHeader(ct, pic)
                                            status = ct.Go()

                                        Catch
                                        End Try
                                    End If
                                End If

                            Case "5"
                            ' special case, will position later
                            ' will redefine the logo!
                            ' imagepath = UPDATED PATH
                            Case "6"

                                If _attr_showonsheet = True Then
                                    Dim ValueStrings As String = _attr_value.Replace("<ul>", "")
                                    ValueStrings = ValueStrings.Replace("</li>", "")
                                    ValueStrings = ValueStrings.Replace("</ul>", "")

                                    ' list
                                    Dim attrHdrPara As Paragraph = New Paragraph
                                    attrHdrPara.SetLeading(leadingHdrF, leadingHdrM)
                                    attrHdrPara.SpacingAfter = spacingHdrAfter
                                    Dim AttrHdrChunk As Chunk = New Chunk(_attr_name, AttrHdrFont)
                                    Dim AttrHdrPhrase As Phrase = New Phrase(AttrHdrChunk)
                                    attrHdrPara.Add(AttrHdrPhrase)

                                    ' this check adds 50 pixels to the yline - this takes into consideration the fact that an item will be below the header
                                    YLine = ct.YLine
                                    AddHeading(ct, attrHdrPara)
                                    CheckTextColumnsWithBuffer(status, ct, column, doc, YLine, 50)
                                    AddHeading(ct, attrHdrPara)
                                    status = ct.Go()
                                    YLine = ct.YLine

                                    ValueStrings = ValueStrings.Replace("[li]", "|")
                                    Dim ItemCtr As Integer = 1
                                    For Each ValueString As String In ValueStrings.Split("|")

                                        '                                    ValueString = ValueString.Replace("li]", "")
                                        Dim attrPara As Paragraph = New Paragraph
                                        attrPara.SetLeading(leadingParaF, leadingParaM)
                                        attrPara.SpacingAfter = spacingParaAfter * 2

                                        If ValueString.Length > 5 Then
                                            Dim FeatureChunk As Chunk = New Chunk("n", zapFont)
                                            If Not (ValueString.StartsWith(" ")) Then
                                                ValueString = " " & ValueString
                                            End If
                                            Dim AttrChunk As Chunk = New Chunk(ValueString, AttrFont)
                                            AttrChunk.SetCharacterSpacing(0.1F)
                                            Dim AttrPhrase As Phrase = New Phrase(FeatureChunk)
                                            AttrPhrase.Add(AttrChunk)
                                            attrPara.Add(AttrPhrase)
                                            attrPara.IndentationLeft = 20.0F
                                            attrPara.FirstLineIndent = -10.0F
                                            attrPara.SpacingAfter = 3.5F
                                            attrPara.Leading = 9.0F

                                            ItemCtr += 1
                                        End If

                                        AddSingleParagraph(ct, attrPara)
                                        CheckTextColumns(status, ct, column, doc, YLine)
                                        AddSingleParagraph(ct, attrPara)
                                        status = ct.Go()

                                        YLine = ct.YLine
                                    Next

                                End If

                            Case "13"

                                If _attr_showonsheet = True Then
                                    Dim ValueStrings As String = _attr_value.Replace("<ul>", "")
                                    ValueStrings = ValueStrings.Replace("</li>", "")
                                    ValueStrings = ValueStrings.Replace("</ul>", "")

                                    ' list
                                    Dim attrHdrPara As Paragraph = New Paragraph
                                    attrHdrPara.SetLeading(leadingHdrF, leadingHdrM)
                                    attrHdrPara.SpacingAfter = spacingHdrAfter
                                    Dim AttrHdrChunk As Chunk = New Chunk(_attr_name, AttrHdrFont)
                                    Dim AttrHdrPhrase As Phrase = New Phrase(AttrHdrChunk)
                                    attrHdrPara.Add(AttrHdrPhrase)

                                    ' this check adds 50 pixels to the yline - this takes into consideration the fact that an item will be below the header
                                    YLine = ct.YLine
                                    AddHeading(ct, attrHdrPara)
                                    CheckTextColumnsWithBuffer(status, ct, column, doc, YLine, 100)
                                    AddHeading(ct, attrHdrPara)
                                    status = ct.Go()
                                    YLine = ct.YLine

                                    ValueStrings = ValueStrings.Replace("[li]", "|")
                                    Dim ItemCtr As Integer = 1
                                    For Each ValueString As String In ValueStrings.Split("|")

                                        Dim attrPara As Paragraph = New Paragraph
                                        attrPara.SetLeading(leadingParaF, leadingParaM)
                                        attrPara.SpacingAfter = spacingParaAfter '* 2

                                        If ValueString.Length > 5 Then
                                            Dim FeatureChunk As Chunk = New Chunk(ItemCtr.ToString & ". ", AttrFont)
                                            If Not (ValueString.StartsWith(" ")) Then
                                                ValueString = " " & ValueString
                                            End If
                                            Dim AttrChunk As Chunk = New Chunk(ValueString.Trim, AttrFont)
                                            AttrChunk.SetCharacterSpacing(0.1F)
                                            Dim AttrPhrase As Phrase = New Phrase(FeatureChunk)
                                            AttrPhrase.Add(AttrChunk)
                                            attrPara.Add(AttrPhrase)
                                            attrPara.IndentationLeft = 10.0F
                                            attrPara.FirstLineIndent = -10.0F
                                            attrPara.SpacingAfter = 5.0F
                                            attrPara.Leading = 10.0F

                                            ItemCtr += 1
                                        End If

                                        AddSingleParagraph(ct, attrPara)
                                        CheckTextColumns(status, ct, column, doc, YLine)
                                        AddSingleParagraph(ct, attrPara)
                                        status = ct.Go()

                                        YLine = ct.YLine
                                    Next

                                End If
                            Case Else

                        End Select
                        '                    Select Case attr.AttributeType.ToString
                    ElseIf _attr_type = 3 Then
                        LastHeader = _attr_name
                        LastHeaderID = _attr_header
                        LastHeaderWritten = False
                    ElseIf Not (IsDBNull(_attr_value)) AndAlso _attr_value <> "" AndAlso _attr_value.ToString.ToUpper = "X" Then

                        ' in case we ned to write out the header, we'll save it here
                        LastHeader = _attr_name
                        LastHeaderID = _attr_header
                        LastHeaderWritten = False

                    End If

                Next

                Try

                    Dim OHeader As String = ""
                    Dim rowCtr As Integer = 0
                    Dim bColor As BaseColor = iTextSharp.text.BaseColor.WHITE

                    Dim aTable As PdfPTable

                    Dim orderList As XmlNodeList = Letter.SelectNodes("related/product")
                    For Each orderInfo As XmlNode In orderList

                        Dim _item_mobiledesc As String = Trim(orderInfo("mobiledesc").InnerText)
                        Dim _item_winfo As String = Trim(orderInfo("winfo").InnerText)
                        Dim _item_oheader As String = Trim(orderInfo("oheader").InnerText)
                        Dim _item_productcode As String = Trim(orderInfo("productcode").InnerText)

                        If _item_mobiledesc <> "" Then
                            If OHeader <> _item_oheader Then

                                Try
                                    ' may need to move this around a bit!
                                    Dim oinfoHdrPara As Paragraph = New Paragraph
                                    oinfoHdrPara.SpacingAfter = 1.0F
                                    Dim oinfoHdrChunk As Chunk = New Chunk(tdsLangInfo.cartOrderingInfo, AttrHdrFont)
                                    Dim oinfoHdrPhrase As Phrase = New Phrase(oinfoHdrChunk)
                                    oinfoHdrPara.Add(oinfoHdrPhrase)

                                    Dim hlCell As PdfPCell
                                    hlCell = New PdfPCell(New Phrase(tdsLangInfo.cartOrderProduct, tableAttrHdrFont))
                                    hlCell.Border = 0
                                    hlCell.HorizontalAlignment = Element.ALIGN_CENTER

                                    Dim hmCell As PdfPCell
                                    hmCell = New PdfPCell(New Phrase(tdsLangInfo.cartOrderDesc, tableAttrHdrFont))
                                    hmCell.Border = 0

                                    Dim hrCell As PdfPCell
                                    hrCell = New PdfPCell(New Phrase(tdsLangInfo.cartOrderWeight, tableAttrHdrFont))
                                    hrCell.Border = 0
                                    hrCell.HorizontalAlignment = Element.ALIGN_RIGHT

                                    If OHeader <> "" Then
                                        WriteAttribute(doc, ct, column, aTable, tableAttrFont)
                                        aTable = New PdfPTable(3)
                                    Else
                                        aTable = New PdfPTable(3)
                                        Dim tabhCell As New PdfPHeaderCell()
                                        tabhCell.Colspan = 3
                                        tabhCell.AddElement(oinfoHdrPara)
                                        tabhCell.Border = 0
                                        aTable.AddCell(tabhCell)

                                        hlCell.AddHeader(tabhCell)
                                    End If

                                    aTable.HorizontalAlignment = Element.ALIGN_CENTER
                                    aTable.TotalWidth = 240.0F
                                    aTable.LockedWidth = True
                                    aTable.WidthPercentage = 100.0F
                                    aTable.SetWidths({60, 130, 50})

                                    If _item_oheader <> "" AndAlso _item_oheader.ToUpper() <> "N/A" Then
                                        Dim oinfoSecPara As Paragraph = New Paragraph
                                        Dim oinfoSecChunk As Chunk = New Chunk(_item_oheader, tableAttrHdrFont)
                                        Dim oinfoSecPhrase As Phrase = New Phrase(oinfoSecChunk)
                                        oinfoSecPara.Add(oinfoSecPhrase)

                                        Dim hCell As New PdfPHeaderCell
                                        hCell.Colspan = 3
                                        hCell.AddElement(oinfoSecPara)
                                        hCell.Border = 0
                                        aTable.AddCell(hCell)

                                        hlCell.AddHeader(hCell)
                                    End If

                                    aTable.AddCell(hlCell)
                                    aTable.AddCell(hmCell)
                                    aTable.AddCell(hrCell)

                                    OHeader = _item_oheader
                                    rowCtr = 0
                                Catch
                                End Try
                            End If

                            If rowCtr Mod 2 = 0 Then
                                bColor = iTextSharp.text.BaseColor.LIGHT_GRAY
                            Else
                                bColor = iTextSharp.text.BaseColor.WHITE
                            End If

                            Try
                                Dim celldata As String = "N/A"

                                If _item_productcode <> "" Then
                                    celldata = _item_productcode
                                End If
                                Dim lCell As PdfPCell
                                lCell = New PdfPCell(New Phrase(CStr(celldata), tableAttrFont))
                                lCell.BorderWidth = 0.25
                                lCell.BorderWidthRight = 0.25
                                lCell.BorderWidthLeft = 0.25
                                lCell.BackgroundColor = bColor
                                lCell.HorizontalAlignment = Element.ALIGN_CENTER
                                aTable.AddCell(lCell)

                                Dim mCell As PdfPCell
                                celldata = ""
                                If _item_mobiledesc <> "" Then
                                    celldata = _item_mobiledesc
                                End If
                                mCell = New PdfPCell(New Phrase(CStr(celldata), tableAttrFont))
                                mCell.BorderWidth = 0.25
                                mCell.BorderWidthRight = 0.25
                                mCell.BackgroundColor = bColor
                                aTable.AddCell(mCell)

                                Dim rCell As PdfPCell
                                rCell = New PdfPCell(New Phrase(CStr(_item_winfo), tableAttrFont))
                                rCell.BorderWidth = 0.25
                                rCell.BorderWidthRight = 0.25
                                rCell.BackgroundColor = bColor
                                rCell.HorizontalAlignment = Element.ALIGN_RIGHT
                                aTable.AddCell(rCell)
                            Catch
                            End Try

                            rowCtr += 1
                        End If
                    Next

                    WriteAttribute(doc, ct, column, aTable, tableAttrFont)
                    aTable = Nothing
                Catch ex1 As Exception
                End Try

                ct.YLine -= 15

                ' Manufacturer
                WriteHeadedParagraph(doc, ct, column, tdsLangInfo.formManufacturerHeader, AttrHdrFont, tdsLangInfo.formManufacturer, AttrFont, True)

                ' warranty area
                Dim warrHdrPara As Paragraph = New Paragraph
                warrHdrPara.SetLeading(leadingHdrF, leadingHdrM)
                Dim warrHdrChunk As Chunk = New Chunk(tdsLangInfo.formWarrantyHeader, AttrHdrFont)
                Dim warrHdrPhrase As Phrase = New Phrase(warrHdrChunk)
                warrHdrPara.Add(warrHdrPhrase)

                Dim warrPara As Paragraph = New Paragraph
                warrPara.SetLeading(leadingParaF, leadingParaM)
                warrPara.SpacingAfter = 10.0F

                Dim warrChunk As Chunk = New Chunk(tdsLangInfo.formWarranty1, NoteFont)
                'warrChunk.SetCharacterSpacing(charspacingPara)
                Dim warrPhrase As Phrase = New Phrase(warrChunk)
                warrPara.Add(warrPhrase)

                YLine = ct.YLine
                AddParagraph(ct, warrHdrPara, warrPara)
                CheckTextColumns(status, ct, column, doc, YLine)
                AddParagraph(ct, warrHdrPara, warrPara)
                status = ct.Go()

            Catch ex As Exception
                'Response.Clear()
                ' Response.End()
            Finally
                doc.Close()
                writer.Close()
            End Try

            ' keep track of the current page
            Dim CurrentPageCtr As Integer = 1

            ' create a reader for a certain document
            Dim readerQQ As PdfReader = New PdfReader(DestPDF)
            Dim n As Integer = readerQQ.NumberOfPages
            TotalPages = n

            Dim document As Document = New Document(readerQQ.GetPageSizeWithRotation(1))

            Dim aFileName As String = _ProductID & ".pdf"
            If _ArtifactName <> "" Then
                aFileName = _ArtifactName
            End If

            Dim aGuid As Guid
            aGuid = Guid.NewGuid()
            aFileName = aGuid.ToString & ".pdf"

            Try
                writer = PdfWriter.GetInstance(document, New FileStream(webFolder & "pdfs\" & Trim(aFileName), FileMode.Create))
            Catch
                aFileName = _ProductID & ".pdf"
                writer = PdfWriter.GetInstance(document, New FileStream(webFolder & "pdfs\" & Trim(aFileName), FileMode.Create))
            End Try
            ' open the document
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

                ' add the black line - will need to remove if not dayton logo!
                cb.SetLineWidth(0.75F)
                cb.SetGrayStroke(0.0F)
                cb.MoveTo(logoimg.ScaledWidth + 20.0F, logoy + 17.0F)
                cb.LineTo(logoimg.ScaledWidth + 20.0F + (headerx2 - (logoimg.ScaledWidth + 20.0F)), logoy + 17.0F)

                cb.Stroke()

                ' reset color
                cb.SetColorFill(New CMYKColor(0.0F, 0.0F, 0.0F, 1.0F))

                ' add header and type
                PdfHelper.PlaceText(txtHeader, productFont, headerx1, headery1, headerx2, headery2, 14, Element.ALIGN_RIGHT)
                PdfHelper.PlaceText(txtProductType, productTypeFont, headerx1, headery1, headerx2, headery2 - 15, 14, Element.ALIGN_RIGHT)
                PdfHelper.PlaceText(txtPageType, sheetTypeFont, 48.0F, logoy - logoimg.ScaledHeight - 25.0F, headerx2, logoy - 10.0F, 14, Element.ALIGN_LEFT)

                PdfHelper.PlaceText(txtFooter1 & CurrentPageCtr.ToString & txtFooter2 & TotalPages.ToString & txtFooter3, New Font(Font.FontFamily.HELVETICA, 8, Font.NORMAL), footerx1, footery1, footerx2, footery2, 14, Element.ALIGN_CENTER)
                If _Updated <> "" Then
                    PdfHelper.PlaceText(txtFooter4 & CDate(_Updated).ToShortDateString, New Font(Font.FontFamily.HELVETICA, 8, Font.NORMAL), footerx1, footery1, footerx2, footery2, 14, Element.ALIGN_RIGHT)
                End If


                CurrentPageCtr += 1

            Loop

            ' close the document
            Try
                document.Close()
            Catch ex As Exception
            End Try

            Try
                readerQQ.Close()
            Catch
            End Try

        End If

    End Sub


    Protected Sub WriteHeadedParagraph(ByRef doc As Document, ct As ColumnText, ByRef column As Integer, Attribute As String, AttrHdrFont As Font, ValueText As String, AttrFont As Font, SkipHeaderSpacing As Boolean, Optional ByRef SectionHdr As String = "", Optional SectionHdrFont As Font = Nothing)

        Dim Status As Integer
        Dim yLine As Double


        Dim sectionHdrPara As Paragraph = New Paragraph
        sectionHdrPara.SetLeading(leadingHdrF, leadingHdrM)
        sectionHdrPara.SpacingAfter = 1.0F
        If SectionHdr <> "" Then
            Dim SectionHdrChunk As Chunk = New Chunk(SectionHdr, SectionHdrFont)
            Dim SectionHdrPhrase As Phrase = New Phrase(SectionHdrChunk)
            sectionHdrPara.Add(SectionHdrPhrase)

        End If

        Dim attrHdrPara As Paragraph = New Paragraph
        attrHdrPara.SetLeading(leadingHdrF, leadingHdrM)
        If (Not (SkipHeaderSpacing)) Then
            attrHdrPara.SpacingAfter = spacingHdrAfter
        End If
        Dim AttrHdrChunk As Chunk = New Chunk(Attribute, AttrHdrFont)
        Dim AttrHdrPhrase As Phrase = New Phrase(AttrHdrChunk)
        attrHdrPara.Add(AttrHdrPhrase)

        Dim attrPara As Paragraph = New Paragraph
        attrPara.SetLeading(leadingParaF, leadingParaM)
        attrPara.SpacingAfter = spacingParaAfter
        Dim AttrChunk As Chunk = New Chunk(ValueText, AttrFont)
        AttrChunk.SetCharacterSpacing(charspacingPara)
        Dim AttrPhrase As Phrase = New Phrase(AttrChunk)
        attrPara.Add(AttrPhrase)

        yLine = ct.YLine
        If SectionHdr <> "" Then
            AddHeadedParagraph(ct, sectionHdrPara, attrHdrPara, attrPara)
        Else
            AddParagraph(ct, attrHdrPara, attrPara)
        End If
        CheckTextColumns(Status, ct, column, doc, yLine)
        If SectionHdr <> "" Then
            AddHeadedParagraph(ct, sectionHdrPara, attrHdrPara, attrPara)
        Else
            AddParagraph(ct, attrHdrPara, attrPara)
        End If
        Status = ct.Go()
    End Sub

    Protected Sub CheckTextColumns(ByRef Status As Integer, ByRef ct As ColumnText, ByRef column As Integer, ByRef doc As Document, ByRef yLine As Double)
        Status = ct.Go(True)

        If (ColumnText.HasMoreText(Status)) Then
            column = (column + 1) Mod 2
            If (column = 0) Then
                doc.NewPage()
            End If
            ct.SetSimpleColumn(Columns(column, 0), Columns(column, 1), Columns(column, 2), Columns(column, 3))
            yLine = Columns(column, 3)
        End If

        ct.YLine = yLine
        ct.SetText(Nothing)
    End Sub

    Protected Sub CheckTextColumnsWithBuffer(ByRef Status As Integer, ByRef ct As ColumnText, ByRef column As Integer, ByRef doc As Document, ByRef yLine As Double, ByRef yLineBuffer As Double)

        Dim origCtY As Double = ct.YLine
        Dim origY As Double = yLine

        ct.YLine -= yLineBuffer
        Status = ct.Go(True)

        If (ColumnText.HasMoreText(Status)) Then
            column = (column + 1) Mod 2
            If (column = 0) Then
                doc.NewPage()
            End If
            ct.SetSimpleColumn(Columns(column, 0), Columns(column, 1), Columns(column, 2), Columns(column, 3))
            yLine = Columns(column, 3)
            ct.YLine = yLine
        Else
            ct.YLine = origCtY
        End If

        '        ct.YLine = yLine
        ct.SetText(Nothing)
    End Sub

    Protected Sub CheckImageColumns(ByRef Status As Integer, ByRef ct As ColumnText, ByRef column As Integer, ByRef doc As Document, ByRef yLine As Double, ByRef Pich As Decimal)
        Status = ct.Go(True)

        If CDbl(Pich) < 155 Then : Pich = 155 : End If
        If (yLine < CDbl(Pich)) Then
            column = (column + 1) Mod 2
            If (column = 0) Then
                doc.NewPage()
            End If
            ct.SetSimpleColumn(Columns(column, 0), Columns(column, 1), Columns(column, 2), Columns(column, 3))
            yLine = Columns(column, 3)
        End If

        ct.YLine = yLine
        ct.SetText(Nothing)

    End Sub

    Protected Sub AddHeading(ct As ColumnText, Hdr As Paragraph)
        ct.AddElement(Hdr)
    End Sub

    Protected Sub AddSingleParagraph(ct As ColumnText, Text As Paragraph)
        ct.AddElement(Text)
    End Sub

    Protected Sub AddParagraph(ct As ColumnText, Hdr As Paragraph, Text As Paragraph)
        ct.AddElement(Hdr)
        ct.AddElement(Text)
    End Sub

    Protected Sub AddHeadedParagraph(ct As ColumnText, Hdr As Paragraph, SubHdr As Paragraph, Text As Paragraph)
        ct.AddElement(Hdr)
        ct.AddElement(SubHdr)
        ct.AddElement(Text)
    End Sub

    Protected Sub AddImage(ct As ColumnText, Hdr As Paragraph, Img As iTextSharp.text.Image)
        ct.AddElement(Hdr)
        ct.AddElement(Img)
    End Sub

    Protected Sub AddImageNoHeader(ct As ColumnText, Img As iTextSharp.text.Image)
        ct.AddElement(Img)
    End Sub


    Protected Sub WriteAttribute(ByRef doc As Document, ct As ColumnText, ByRef column As Integer, ValueText As PdfPTable, AttrFont As Font)

        Dim Status As Integer
        Dim yLine As Double

        yLine = ct.YLine

        ct.AddElement(ValueText)
        CheckTextColumns(Status, ct, column, doc, yLine)
        ct.AddElement(ValueText)
        Status = ct.Go()
    End Sub

    Private Sub AddCell(ByRef aTable As PdfPTable, ByRef hlCell As PdfPCell, bColor As BaseColor, fColor As BaseColor, cellHeight As Single, cellHAlign As Integer, cellVAlign As Integer, Optional ByVal BorderTop As Boolean = True, Optional ByVal BorderBottom As Boolean = True, Optional ByVal BorderLeft As Boolean = True, Optional ByVal BorderRight As Boolean = True)
        hlCell.BorderWidth = 1.0

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

    Protected Sub CheckTableColumns(ByRef Status As Integer, ByRef ct As ColumnText, ByRef column As Integer, ByRef doc As Document, ByRef yLine As Double, ByRef GridRows As XmlNodeList)
        '        Status = ct.Go(True)

        Dim rowcount As Integer = GridRows.Count

        ' we may need to add additional variables to this method to make it text the image for more resizing, etc, for now this seems to work.
        If (ct.YLine < (rowcount * (cells * 1.2)) + 80) Then
            column = (column + 1) Mod 2
            If (column = 0) Then
                doc.NewPage()
            End If
            ct.SetSimpleColumn(Columns(column, 0), Columns(column, 1), Columns(column, 2), Columns(column, 3))
            yLine = Columns(column, 3)
            ct.YLine = yLine
        End If

        ct.SetText(Nothing)
    End Sub

    Protected Sub WriteTable(ByRef doc As Document, ct As ColumnText, ByRef column As Integer, ByRef YLine As Double, ByRef GridRows As XmlNodeList, ByVal maxCols As Integer)


        Dim tablepropw As Single = 240.0F
        Dim cellw As Single = 24.0F

        Dim maxRows As Integer = GridRows.Count
        Dim curCols As Integer = 0
        Dim AttrFont As Font
        Dim AttrFontWhite As Font

        If maxCols > 0 Then
            Dim Tabletempl As PdfPTable
            Tabletempl = New PdfPTable(maxCols)
            Tabletempl.HorizontalAlignment = 0
            Tabletempl.TotalWidth = tablepropw
            Tabletempl.LockedWidth = True
            Dim widths As Single()

            Dim cellwadj As Single = tablepropw / maxCols
            ' if you change the font here, change it later when the table is written!
            Select Case maxCols
                Case 1
                    widths = New Single() {cellwadj}
                    AttrFont = tableAttrFont9
                    AttrFontWhite = tableAttrFont9white
                Case 2
                    widths = New Single() {cellwadj, cellwadj}
                    AttrFont = tableAttrFont9
                    AttrFontWhite = tableAttrFont9white
                Case 3
                    widths = New Single() {cellwadj, cellwadj, cellwadj}
                    AttrFont = tableAttrFont9
                    AttrFontWhite = tableAttrFont9white
                Case 4
                    widths = New Single() {cellwadj, cellwadj, cellwadj, cellwadj}
                    AttrFont = tableAttrFont9
                    AttrFontWhite = tableAttrFont9white
                Case 5
                    widths = New Single() {cellwadj, cellwadj, cellwadj, cellwadj, cellwadj}
                    AttrFont = tableAttrFont9
                    AttrFontWhite = tableAttrFont9white
                Case 6
                    widths = New Single() {cellwadj, cellwadj, cellwadj, cellwadj, cellwadj, cellwadj}
                    AttrFont = tableAttrFont8
                    AttrFontWhite = tableAttrFont8white
                Case 7
                    widths = New Single() {cellwadj, cellwadj, cellwadj, cellwadj, cellwadj, cellwadj, cellwadj}
                    AttrFont = tableAttrFont8
                    AttrFontWhite = tableAttrFont8white
                Case 8
                    widths = New Single() {cellwadj, cellwadj, cellwadj, cellwadj, cellwadj, cellwadj, cellwadj, cellwadj}
                    AttrFont = tableAttrFont8
                    AttrFontWhite = tableAttrFont8white
                Case 9
                    widths = New Single() {cellwadj, cellwadj, cellwadj, cellwadj, cellwadj, cellwadj, cellwadj, cellwadj, cellwadj}
                    AttrFont = tableAttrFont7
                    AttrFontWhite = tableAttrFont7white
                Case 10
                    widths = New Single() {cellwadj, cellwadj, cellwadj, cellwadj, cellwadj, cellwadj, cellwadj, cellwadj, cellwadj, cellwadj}
                    AttrFont = tableAttrFont7
                    AttrFontWhite = tableAttrFont7white
                Case Else
                    widths = New Single() {cellwadj, cellwadj, cellwadj, cellwadj, cellwadj, cellwadj, cellwadj, cellwadj, cellwadj, cellwadj}
                    AttrFont = tableAttrFont9
                    AttrFontWhite = tableAttrFont9white
            End Select
            Tabletempl.SetWidths(widths)

            Dim aTable As PdfPTable = New PdfPTable(Tabletempl)
            Dim hlCell As PdfPCell

            If GridRows IsNot Nothing AndAlso GridRows.Count > 0 Then
                Dim curRow As Integer = 0
                Dim Col1RowSpan As Boolean = False
                Dim Col2RowSpan As Boolean = False
                Dim Col3RowSpan As Boolean = False
                Dim Col4RowSpan As Boolean = False
                Dim Col5RowSpan As Boolean = False
                Dim Col6RowSpan As Boolean = False
                Dim Col7RowSpan As Boolean = False
                Dim Col8RowSpan As Boolean = False
                Dim Col9RowSpan As Boolean = False
                Dim Col10RowSpan As Boolean = False

                Dim extraHeader As Boolean = False

                For Each gridRow As XmlNode In GridRows
                    Dim _col1val As String = Trim(gridRow("col1").InnerText)
                    Dim _col2val As String = Trim(gridRow("col2").InnerText)
                    Dim _col3val As String = Trim(gridRow("col3").InnerText)
                    Dim _col4val As String = Trim(gridRow("col4").InnerText)
                    Dim _col5val As String = Trim(gridRow("col5").InnerText)
                    Dim _col6val As String = Trim(gridRow("col6").InnerText)
                    Dim _col7val As String = Trim(gridRow("col7").InnerText)
                    Dim _col8val As String = Trim(gridRow("col8").InnerText)
                    Dim _col9val As String = Trim(gridRow("col9").InnerText)
                    Dim _col10val As String = Trim(gridRow("col10").InnerText)

                    Dim colsRendered As Integer = 0
                    Dim curCell As Integer = 1
                    Dim emptyrow As Boolean = False

                    If curRow < maxRows Then

                        '#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#
                        ' Note: Column 1 should not be spanned with any other column unless the header is spanned also!
                        '#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#
                        '# Row 1
                        '#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#
                        curCell = 0
                        Dim checkCell As Integer = curCell

                        If _col1val <> "" AndAlso _col1val = "EMPTYROW" Then
                            ' put a break in the table (no border)
                            hlCell = New PdfPCell(New Phrase("", AttrFont))
                            hlCell.Colspan = maxCols
                            AddCell(aTable, hlCell, bColor, fColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE, True, True, False, False)
                            emptyrow = True
                            extraHeader = True
                            colsRendered = maxCols
                        ElseIf _col1val <> "" Then
                            If curRow > 0 And Not (extraHeader) Then
                                hlCell = New PdfPCell(New Phrase(_col1val.ToString.Replace("|", vbCrLf), AttrFont))
                            Else
                                hlCell = New PdfPCell(New Phrase(_col1val.ToString.Replace("|", vbCrLf), AttrFontWhite))
                            End If

                            Dim spanCells As Integer = 1
                            Dim checkDone As Boolean = False
                            checkCell = curCell + 1

                            ' don't look ahead if we are the last row!
                            Col1RowSpan = False
                            If curRow < maxRows - 1 Then
                                Dim nextRow As String = GridRows(curRow + 1).ChildNodes(curCell).InnerText
                                If nextRow = "" Then
                                    hlCell.Rowspan = 2
                                    Col1RowSpan = True
                                End If
                            End If

                            'If curRow < 3 Then
                            If Not (checkDone) AndAlso Not (Col2RowSpan) AndAlso checkCell <= maxCols AndAlso _col2val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col3val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col4val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col5val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col6val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col7val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col8val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col9val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col10val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            'End If

                            ' default to the one below us
                            Dim nextRowCol As Integer = curCell
                            Dim SpannedCol As Boolean = False
                            If Col1RowSpan And _col2val = "" Then
                                ' check in the next row to see what the 1st column is that has a value
                                For i = curCell To maxCols
                                    Dim nextColCheck As String = GridRows(curRow - 1).ChildNodes(i).InnerText
                                    If nextColCheck <> "" Then
                                        nextRowCol = i

                                        If spanCells >= nextRowCol Then
                                            spanCells = nextRowCol - curCell
                                            SpannedCol = True
                                        End If

                                        Exit For
                                    End If
                                Next
                            End If

                            If (extraHeader) Then
                                spanCells = maxCols
                                SpannedCol = False
                            End If
                            hlCell.Colspan = spanCells

                            If curRow > 0 And Not (extraHeader) Then
                                AddCell(aTable, hlCell, bColor, fColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                            Else
                                AddCell(aTable, hlCell, hdrbColor, hdrfColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                                extraHeader = False
                            End If
                            colsRendered += spanCells
                            If SpannedCol Then : colsRendered += 1 : End If
                            checkCell = curCell + spanCells - 1

                        Else
                            If Col1RowSpan Then
                                colsRendered += 1
                                Col1RowSpan = False
                            Else

                                ' create a new cell width for all cells other than 1
                                Dim cellw1wide As Single = 40.0F
                                Dim cellwnew As Single = (tablepropw - cellw1wide) / maxCols
                                Select Case maxCols
                                    Case 1
                                        widths = New Single() {tablepropw}
                                    Case 2
                                        widths = New Single() {cellwadj, cellwadj}
                                    Case 3
                                        widths = New Single() {cellwadj, cellwadj, cellwadj}
                                    Case 4
                                        widths = New Single() {cellwadj, cellwadj, cellwadj, cellwadj}
                                    Case 5
                                        widths = New Single() {cellw1wide, cellwnew, cellwnew, cellwnew, cellwnew}
                                    Case 6
                                        widths = New Single() {cellw1wide, cellwnew, cellwnew, cellwnew, cellwnew, cellwnew}
                                    Case 7
                                        widths = New Single() {cellw1wide, cellwnew, cellwnew, cellwnew, cellwnew, cellwnew, cellwnew}
                                    Case 8
                                        widths = New Single() {cellw1wide, cellwnew, cellwnew, cellwnew, cellwnew, cellwnew, cellwnew, cellwnew}
                                    Case 9
                                        widths = New Single() {cellw1wide, cellwnew, cellwnew, cellwnew, cellwnew, cellwnew, cellwnew, cellwnew, cellwnew}
                                    Case 10
                                        widths = New Single() {cellw1wide, cellwnew, cellwnew, cellwnew, cellwnew, cellwnew, cellwnew, cellwnew, cellwnew, cellwnew}
                                    Case Else
                                        widths = New Single() {cellw1wide, cellwnew, cellwnew, cellwnew, cellwnew, cellwnew, cellwnew, cellwnew, cellwnew, cellwnew}
                                End Select
                                Tabletempl.SetWidths(widths)
                                aTable = New PdfPTable(Tabletempl)

                                colsRendered += 1
                                hlCell = New PdfPCell(New Phrase("", AttrFont))
                                If curRow > 0 Then
                                    AddCell(aTable, hlCell, bColor, fColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                                Else
                                    AddCell(aTable, hlCell, hdrbColor, hdrfColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                                End If
                            End If
                            'Else
                            '    If maxCols >= curCell Then
                            '        hlCell = New PdfPCell(New Phrase(_col1val.ToString, AttrFont))
                            '        If maxCols >= curCell + 1 Then
                            '            If _col2val Is Nothing Then
                            '                hlCell.Colspan = 2
                            '                colsRendered += 1
                            '            End If
                            '        End If

                            '        AddCell(aTable, hlCell, bColor, fColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                            '        colsRendered += 1
                            '    End If
                        End If

                        '#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#
                        '# Row 2
                        '#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#
                        curCell = 1
                        If emptyrow Then
                            ' skip
                        ElseIf colsRendered < maxCols AndAlso checkCell < curCell AndAlso _col2val <> "" Then
                            If curRow > 0 Then
                                hlCell = New PdfPCell(New Phrase(_col2val.ToString.Replace("|", vbCrLf), AttrFont))
                            Else
                                hlCell = New PdfPCell(New Phrase(_col2val.ToString.Replace("|", vbCrLf), AttrFontWhite))
                            End If
                            ' don't look ahead if we are the last row!

                            Dim spanCells As Integer = 1
                            Dim checkDone As Boolean = False
                            checkCell = curCell + 1

                            Col2RowSpan = False
                            If curRow < maxRows - 1 Then
                                Dim nextRow As String = GridRows(curRow + 1).ChildNodes(curCell).InnerText
                                If nextRow = "" Then
                                    hlCell.Rowspan = 2
                                    Col2RowSpan = True
                                End If
                            End If

                            'If curRow < 3 Then
                            If Not (checkDone) AndAlso Not (Col3RowSpan) AndAlso checkCell <= maxCols AndAlso _col3val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col4val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col5val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col6val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col7val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col8val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col9val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col10val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            'End If

                            ' default to the one below us
                            Dim nextRowCol As Integer = curCell
                            Dim SpannedCol As Boolean = False
                            If Col2RowSpan And _col3val = "" Then
                                ' check in the next row to see what the 1st column is that has a value
                                For i = curCell To maxCols
                                    Dim nextColCheck As String = GridRows(curRow - 1).ChildNodes(i).InnerText
                                    If nextColCheck <> "" Then
                                        nextRowCol = i

                                        If spanCells >= nextRowCol Then
                                            spanCells = nextRowCol - curCell
                                            SpannedCol = True
                                        End If
                                        Exit For
                                    End If
                                Next
                            End If

                            hlCell.Colspan = spanCells

                            If curRow > 0 Then
                                AddCell(aTable, hlCell, bColor, fColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                            Else
                                AddCell(aTable, hlCell, hdrbColor, hdrfColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                            End If
                            colsRendered += spanCells
                            If SpannedCol Then : colsRendered += 1 : End If
                            checkCell = curCell + spanCells - 1
                        Else
                            If Col2RowSpan Then
                                colsRendered += 1
                                Col2RowSpan = False
                            End If
                        End If

                        '#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#
                        '# Row 3
                        '#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#
                        curCell = 2
                        If emptyrow Then
                            ' skip
                        ElseIf colsRendered < maxCols AndAlso checkCell < curCell AndAlso _col3val <> "" Then
                            If curRow > 0 Then
                                hlCell = New PdfPCell(New Phrase(_col3val.ToString.Replace("|", vbCrLf), AttrFont))
                            Else
                                hlCell = New PdfPCell(New Phrase(_col3val.ToString.Replace("|", vbCrLf), AttrFontWhite))
                            End If
                            ' don't look ahead if we are the last row!

                            Dim spanCells As Integer = 1
                            Dim checkDone As Boolean = False
                            checkCell = curCell + 1

                            Col3RowSpan = False
                            If curRow < maxRows - 1 Then
                                Dim nextRow As String = GridRows(curRow + 1).ChildNodes(curCell).InnerText
                                If nextRow = "" Then
                                    hlCell.Rowspan = 2
                                    Col3RowSpan = True
                                End If
                            End If

                            'If curRow < 3 Then
                            If Not (checkDone) AndAlso Not (Col4RowSpan) AndAlso checkCell <= maxCols AndAlso _col4val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col5val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col6val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col7val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col8val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col9val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col10val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            'End If

                            ' default to the one below us
                            Dim nextRowCol As Integer = curCell
                            Dim SpannedCol As Boolean = False
                            If Col3RowSpan Then
                                ' check in the next row to see what the 1st column is that has a value
                                For i = curCell To maxCols
                                    Dim nextColCheck As String = GridRows(curRow - 1).ChildNodes(i).InnerText
                                    If nextColCheck <> "" Then
                                        nextRowCol = i

                                        If spanCells >= nextRowCol Then
                                            spanCells = nextRowCol - curCell
                                            SpannedCol = True
                                        End If
                                        Exit For
                                    End If
                                Next
                            End If

                            hlCell.Colspan = spanCells

                            If curRow > 0 Then
                                AddCell(aTable, hlCell, bColor, fColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                            Else
                                AddCell(aTable, hlCell, hdrbColor, hdrfColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                            End If
                            colsRendered += spanCells
                            If SpannedCol Then : colsRendered += 1 : End If
                            checkCell = curCell + spanCells - 1
                        Else
                            If Col3RowSpan Then
                                colsRendered += 1
                                Col3RowSpan = False
                            Else
                                'colsRendered += 1
                                'FillRowEmpty(aTable, colsRendered, maxCols, bColor, cells)
                            End If
                        End If

                        '#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#
                        '# Row 4
                        '#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#
                        curCell = 3
                        If emptyrow Then
                            ' skip
                        ElseIf colsRendered < maxCols AndAlso checkCell < curCell AndAlso _col4val <> "" Then
                            If curRow > 0 Then
                                hlCell = New PdfPCell(New Phrase(_col4val.ToString.Replace("|", vbCrLf), AttrFont))
                            Else
                                hlCell = New PdfPCell(New Phrase(_col4val.ToString.Replace("|", vbCrLf), AttrFontWhite))
                            End If
                            ' don't look ahead if we are the last row!

                            Dim spanCells As Integer = 1
                            Dim checkDone As Boolean = False
                            checkCell = curCell + 1

                            Col4RowSpan = False
                            If curRow < maxRows - 1 Then
                                Dim nextRow As String = GridRows(curRow + 1).ChildNodes(curCell - 1).InnerText
                                If nextRow = "" Then
                                    hlCell.Rowspan = 2
                                    Col4RowSpan = True
                                End If
                            End If

                            'If curRow < 3 Then
                            If Not (checkDone) AndAlso Not (Col5RowSpan) AndAlso checkCell <= maxCols AndAlso _col5val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col6val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col7val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col8val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col9val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col10val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            'End If

                            ' default to the one below us
                            Dim nextRowCol As Integer = curCell
                            Dim SpannedCol As Boolean = False
                            If Col4RowSpan Then
                                ' check in the next row to see what the 1st column is that has a value
                                For i = curCell To maxCols
                                    Dim nextColCheck As String = GridRows(curRow - 1).ChildNodes(i).InnerText
                                    If nextColCheck <> "" Then
                                        nextRowCol = i

                                        If spanCells >= nextRowCol Then
                                            spanCells = nextRowCol - curCell
                                            SpannedCol = True
                                        End If
                                        Exit For
                                    End If
                                Next
                            End If

                            hlCell.Colspan = spanCells

                            If curRow > 0 Then
                                AddCell(aTable, hlCell, bColor, fColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                            Else
                                AddCell(aTable, hlCell, hdrbColor, hdrfColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                            End If
                            colsRendered += spanCells
                            If SpannedCol Then : colsRendered += 1 : End If
                            checkCell = curCell + spanCells - 1
                        Else
                            If Col4RowSpan Then
                                colsRendered += 1
                                Col4RowSpan = False
                            Else
                                'colsRendered += 1
                            End If
                        End If

                        '#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#
                        '# Row 5
                        '#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#
                        curCell = 4
                        If emptyrow Then
                            ' skip
                        ElseIf colsRendered < maxCols AndAlso checkCell < curCell AndAlso _col5val <> "" Then
                            If curRow > 0 Then
                                hlCell = New PdfPCell(New Phrase(_col5val.ToString.Replace("|", vbCrLf), AttrFont))
                            Else
                                hlCell = New PdfPCell(New Phrase(_col5val.ToString.Replace("|", vbCrLf), AttrFontWhite))
                            End If
                            ' don't look ahead if we are the last row!

                            Dim spanCells As Integer = 1
                            Dim checkDone As Boolean = False
                            checkCell = curCell + 1

                            Col5RowSpan = False
                            If curRow < maxRows - 1 Then
                                Dim nextRow As String = GridRows(curRow + 1).ChildNodes(curCell).InnerText
                                If nextRow = "" Then
                                    hlCell.Rowspan = 2
                                    Col5RowSpan = True
                                End If
                            End If

                            'If curRow < 3 Then
                            If Not (checkDone) AndAlso Not (Col6RowSpan) AndAlso checkCell <= maxCols AndAlso _col6val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col7val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col8val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col9val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col10val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            'End If

                            ' default to the one below us
                            Dim nextRowCol As Integer = curCell
                            Dim SpannedCol As Boolean = False
                            If Col5RowSpan Then
                                ' check in the next row to see what the 1st column is that has a value
                                For i = curCell To maxCols
                                    Dim nextColCheck As String = GridRows(curRow - 1).ChildNodes(i).InnerText
                                    If nextColCheck <> "" Then
                                        nextRowCol = i

                                        If spanCells >= nextRowCol Then
                                            spanCells = nextRowCol - curCell
                                            SpannedCol = True
                                        End If
                                        Exit For
                                    End If
                                Next
                            End If

                            hlCell.Colspan = spanCells

                            If curRow > 0 Then
                                AddCell(aTable, hlCell, bColor, fColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                            Else
                                AddCell(aTable, hlCell, hdrbColor, hdrfColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                            End If
                            colsRendered += spanCells
                            If SpannedCol Then : colsRendered += 1 : End If
                            checkCell = curCell + spanCells - 1
                        Else
                            If Col5RowSpan Then
                                colsRendered += 1
                            End If
                        End If

                        '#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#
                        '# Row 6
                        '#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#
                        curCell = 5
                        If emptyrow Then
                            ' skip
                        ElseIf colsRendered < maxCols AndAlso checkCell < curCell AndAlso _col6val <> "" Then
                            If curRow > 0 Then
                                hlCell = New PdfPCell(New Phrase(_col6val.ToString.Replace("|", vbCrLf), AttrFont))
                            Else
                                hlCell = New PdfPCell(New Phrase(_col6val.ToString.Replace("|", vbCrLf), AttrFontWhite))
                            End If
                            ' don't look ahead if we are the last row!

                            Dim spanCells As Integer = 1
                            Dim checkDone As Boolean = False
                            checkCell = curCell + 1

                            Col6RowSpan = False
                            If curRow < maxRows - 1 Then
                                Dim nextRow As String = GridRows(curRow + 1).ChildNodes(curCell).InnerText
                                If nextRow = "" Then
                                    hlCell.Rowspan = 2
                                    Col6RowSpan = True
                                End If
                            End If

                            'If curRow < 3 Then
                            If Not (checkDone) AndAlso Not (Col7RowSpan) AndAlso checkCell <= maxCols AndAlso _col7val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col8val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col9val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col10val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            'End If

                            ' default to the one below us
                            Dim nextRowCol As Integer = curCell
                            Dim SpannedCol As Boolean = False
                            If Col6RowSpan Then
                                ' check in the next row to see what the 1st column is that has a value
                                For i = curCell To maxCols
                                    Dim nextColCheck As String = GridRows(curRow - 1).ChildNodes(i).InnerText
                                    If nextColCheck <> "" Then
                                        nextRowCol = i

                                        If spanCells >= nextRowCol Then
                                            spanCells = nextRowCol - curCell
                                            SpannedCol = True
                                        End If
                                        Exit For
                                    End If
                                Next
                            End If

                            hlCell.Colspan = spanCells

                            If curRow > 0 Then
                                AddCell(aTable, hlCell, bColor, fColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                            Else
                                AddCell(aTable, hlCell, hdrbColor, hdrfColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                            End If
                            colsRendered += spanCells
                            If SpannedCol Then : colsRendered += 1 : End If
                            checkCell = curCell + spanCells - 1
                        Else
                            If Col6RowSpan Then
                                colsRendered += 1
                            End If
                        End If

                        '#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#
                        '# Row 7
                        '#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#
                        curCell = 6
                        If emptyrow Then
                            ' skip
                        ElseIf colsRendered < maxCols AndAlso checkCell < curCell AndAlso _col7val <> "" Then
                            If curRow > 0 Then
                                hlCell = New PdfPCell(New Phrase(_col7val.ToString.Replace("|", vbCrLf), AttrFont))
                            Else
                                hlCell = New PdfPCell(New Phrase(_col7val.ToString.Replace("|", vbCrLf), AttrFontWhite))
                            End If
                            ' don't look ahead if we are the last row!

                            Dim spanCells As Integer = 1
                            Dim checkDone As Boolean = False
                            checkCell = curCell + 1

                            Col7RowSpan = False
                            If curRow < maxRows - 1 Then
                                Dim nextRow As String = GridRows(curRow + 1).ChildNodes(curCell).InnerText
                                If nextRow = "" Then
                                    hlCell.Rowspan = 2
                                    Col7RowSpan = True
                                End If
                            End If

                            'If curRow < 3 Then
                            If Not (checkDone) AndAlso Not (Col8RowSpan) AndAlso checkCell <= maxCols AndAlso _col8val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col9val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col10val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            'End If

                            ' default to the one below us
                            Dim nextRowCol As Integer = curCell
                            Dim SpannedCol As Boolean = False
                            If Col7RowSpan Then
                                ' check in the next row to see what the 1st column is that has a value
                                For i = curCell To maxCols
                                    Dim nextColCheck As String = GridRows(curRow - 1).ChildNodes(i).InnerText
                                    If nextColCheck <> "" Then
                                        nextRowCol = i

                                        If spanCells >= nextRowCol Then
                                            spanCells = nextRowCol - curCell
                                            SpannedCol = True
                                        End If
                                        Exit For
                                    End If
                                Next
                            End If

                            hlCell.Colspan = spanCells

                            If curRow > 0 Then
                                AddCell(aTable, hlCell, bColor, fColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                            Else
                                AddCell(aTable, hlCell, hdrbColor, hdrfColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                            End If
                            colsRendered += spanCells
                            If SpannedCol Then : colsRendered += 1 : End If
                            checkCell = curCell + spanCells - 1
                        Else
                            If Col7RowSpan Then
                                colsRendered += 1
                            End If
                        End If

                        '#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#
                        '# Row 8
                        '#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#
                        curCell = 7
                        If emptyrow Then
                            ' skip
                        ElseIf colsRendered < maxCols AndAlso checkCell < curCell AndAlso _col8val <> "" Then
                            If curRow > 0 Then
                                hlCell = New PdfPCell(New Phrase(_col8val.ToString.Replace("|", vbCrLf), AttrFont))
                            Else
                                hlCell = New PdfPCell(New Phrase(_col8val.ToString.Replace("|", vbCrLf), AttrFontWhite))
                            End If
                            ' don't look ahead if we are the last row!

                            Dim spanCells As Integer = 1
                            Dim checkDone As Boolean = False
                            checkCell = curCell + 1

                            Col8RowSpan = False
                            If curRow < maxRows - 1 Then
                                Dim nextRow As String = GridRows(curRow + 1).ChildNodes(curCell).InnerText
                                If nextRow = "" Then
                                    hlCell.Rowspan = 2
                                    Col8RowSpan = True
                                End If
                            End If

                            'If curRow < 3 Then
                            If Not (checkDone) AndAlso Not (Col9RowSpan) AndAlso checkCell <= maxCols AndAlso _col9val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            If Not (checkDone) AndAlso checkCell <= maxCols AndAlso _col10val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            'End If

                            ' default to the one below us
                            Dim nextRowCol As Integer = curCell
                            Dim SpannedCol As Boolean = False
                            If Col8RowSpan Then
                                ' check in the next row to see what the 1st column is that has a value
                                For i = curCell To maxCols
                                    Dim nextColCheck As String = GridRows(curRow - 1).ChildNodes(i).InnerText
                                    If nextColCheck <> "" Then
                                        nextRowCol = i

                                        If spanCells >= nextRowCol Then
                                            spanCells = nextRowCol - curCell
                                            SpannedCol = True
                                        End If
                                        Exit For
                                    End If
                                Next
                            End If

                            hlCell.Colspan = spanCells

                            If curRow > 0 Then
                                AddCell(aTable, hlCell, bColor, fColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                            Else
                                AddCell(aTable, hlCell, hdrbColor, hdrfColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                            End If
                            colsRendered += spanCells
                            If SpannedCol Then : colsRendered += 1 : End If
                            checkCell = curCell + spanCells - 1
                        Else
                            If Col8RowSpan Then
                                colsRendered += 1
                            End If
                        End If

                        '#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#
                        '# Row 9
                        '#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#
                        curCell = 8
                        If emptyrow Then
                            ' skip
                        ElseIf colsRendered < maxCols AndAlso checkCell < curCell AndAlso _col9val <> "" Then
                            If curRow > 0 Then
                                hlCell = New PdfPCell(New Phrase(_col9val.ToString.Replace("|", vbCrLf), AttrFont))
                            Else
                                hlCell = New PdfPCell(New Phrase(_col9val.ToString.Replace("|", vbCrLf), AttrFontWhite))
                            End If
                            ' don't look ahead if we are the last row!

                            Dim spanCells As Integer = 1
                            Dim checkDone As Boolean = False
                            checkCell = curCell + 1

                            Col9RowSpan = False
                            If curRow < maxRows - 1 Then
                                Dim nextRow As String = GridRows(curRow + 1).ChildNodes(curCell).InnerText
                                If nextRow = "" Then
                                    hlCell.Rowspan = 2
                                    Col9RowSpan = True
                                End If
                            End If

                            'If curRow < 3 Then
                            If Not (checkDone) AndAlso Not (Col10RowSpan) AndAlso checkCell <= maxCols AndAlso _col10val = "" Then : spanCells += 1 : checkCell += 1 : Else : checkDone = True : End If
                            'End If

                            ' default to the one below us
                            Dim nextRowCol As Integer = curCell
                            Dim SpannedCol As Boolean = False
                            If Col9RowSpan Then
                                ' check in the next row to see what the 1st column is that has a value
                                For i = curCell To maxCols
                                    Dim nextColCheck As String = GridRows(curRow - 1).ChildNodes(i).InnerText
                                    If nextColCheck <> "" Then
                                        nextRowCol = i

                                        If spanCells >= nextRowCol Then
                                            spanCells = nextRowCol - curCell
                                            SpannedCol = True
                                        End If
                                        Exit For
                                    End If
                                Next
                            End If

                            hlCell.Colspan = spanCells

                            If curRow > 0 Then
                                AddCell(aTable, hlCell, bColor, fColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                            Else
                                AddCell(aTable, hlCell, hdrbColor, hdrfColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                            End If
                            colsRendered += spanCells
                            If SpannedCol Then : colsRendered += 1 : End If
                            checkCell = curCell + spanCells - 1
                        Else
                            If Col9RowSpan Then
                                colsRendered += 1
                            End If
                        End If

                        '#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#
                        '# Row 10
                        '#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#=#
                        curCell = 9
                        If emptyrow Then
                            ' skip
                        ElseIf colsRendered < maxCols AndAlso checkCell < curCell AndAlso _col10val <> "" Then
                            If curRow > 0 Then
                                hlCell = New PdfPCell(New Phrase(_col10val.ToString.Replace("|", vbCrLf), AttrFont))
                            Else
                                hlCell = New PdfPCell(New Phrase(_col10val.ToString.Replace("|", vbCrLf), AttrFontWhite))
                            End If
                            ' don't look ahead if we are the last row!

                            Dim spanCells As Integer = 1
                            Dim checkDone As Boolean = False
                            checkCell = curCell + 1

                            Col10RowSpan = False
                            If curRow < maxRows - 1 Then
                                Dim nextRow As String = GridRows(curRow + 1).ChildNodes(curCell).InnerText
                                If nextRow = "" Then
                                    hlCell.Rowspan = 2
                                    Col10RowSpan = True
                                End If
                            End If

                            ' default to the one below us
                            Dim nextRowCol As Integer = curCell
                            Dim SpannedCol As Boolean = False
                            If Col10RowSpan Then
                                ' check in the next row to see what the 1st column is that has a value
                                For i = curCell To maxCols
                                    Dim nextColCheck As String = GridRows(curRow - 1).ChildNodes(i).InnerText
                                    If nextColCheck <> "" Then
                                        nextRowCol = i

                                        If spanCells >= nextRowCol Then
                                            spanCells = nextRowCol - curCell
                                            SpannedCol = True
                                        End If
                                        Exit For
                                    End If
                                Next
                            End If

                            hlCell.Colspan = spanCells

                            If curRow > 0 Then
                                AddCell(aTable, hlCell, bColor, fColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                            Else
                                AddCell(aTable, hlCell, hdrbColor, hdrfColor, cells, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
                            End If
                            colsRendered += spanCells
                            If SpannedCol Then : colsRendered += 1 : End If
                            checkCell = curCell + spanCells - 1
                        Else
                            If Col10RowSpan Then
                                colsRendered += 1
                            End If
                        End If

                        YLine -= cells

                        If colsRendered < maxCols AndAlso checkCell < maxCols Then
                            FillRowEmpty(aTable, colsRendered, maxCols, bColor, fColor, cells)
                        End If

                    End If

                    curRow += 1
                Next

                ' Writing to document
                Select Case maxCols
                    Case 1, 2, 3, 4, 5
                        WriteAttribute(doc, ct, column, aTable, tableAttrFont9)
                    Case 6, 7, 8
                        WriteAttribute(doc, ct, column, aTable, tableAttrFont8)
                    Case 9, 10
                        WriteAttribute(doc, ct, column, aTable, tableAttrFont7)
                End Select

            End If

        End If
    End Sub

    Private Sub FillRowEmpty(ByRef aTable As PdfPTable, dCtr As Integer, maxCols As Integer, cellbColor As BaseColor, cellfColor As BaseColor, cellsize As Single)

        ' fill out the rest of the table
        For ictr As Integer = dCtr To maxCols Step 1
            Dim hlCell = New PdfPCell(New Phrase("", tableAttrFont9))
            AddCell(aTable, hlCell, cellbColor, cellfColor, cellsize, Element.ALIGN_CENTER, Element.ALIGN_MIDDLE)
        Next

    End Sub

End Class


// Copyright (c) Bodoconsult EDV-Dienstleistungen GmbH. All rights reserved.


using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Bodoconsult.Core.Pdf.Helpers;
using Bodoconsult.Core.Pdf.Stylesets;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Shapes.Charts;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;

namespace Bodoconsult.Core.Pdf.PdfSharp
{
    /// <summary>
    /// Class representing a PDF document and basic functionality to add content to it
    /// </summary>
    public class PdfCreator : IDisposable
    {

        #region Constructors

        /// <summary>
        /// Default ctor
        /// </summary>
        public PdfCreator()
        {

            LoadDefaults();

            // Get the predefined style Normal.
            var style = _document.Styles["Normal"];
            // Because all styles are derived from Normal, the next line changes the 
            // font of the whole document. Or, more exactly, it changes the font of
            // all styles and paragraphs that do not redefine the font.
            style.Font.Name = "Arial";
            style.Font.Size = 11;

        }

        /// <summary>
        /// Ctor with basic style settings
        /// </summary>
        /// <param name="fontName">Default font name to use</param>
        /// <param name="fontSize">Default font size to use</param>
        public PdfCreator(string fontName, Unit fontSize)
        {
            LoadDefaults();

            // Get the predefined style Normal.
            var style = _document.Styles["Normal"];
            // Because all styles are derived from Normal, the next line changes the 
            // font of the whole document. Or, more exactly, it changes the font of
            // all styles and paragraphs that do not redefine the font.
            style.Font.Name = fontName;
            style.Font.Size = fontSize;

        }

        /// <summary>
        /// Ctor to load a complete styleset
        /// </summary>
        /// <param name="styleSet">Styleset to use</param>
        public PdfCreator(IStyleSet styleSet)
        {
            LoadDefaults();

            _ps = styleSet.PageSetup;

            ObjectHelper.MapProperties(_ps, _document.DefaultPageSetup);

            var style = _document.Styles["Normal"];

            // ToDo complete clone of NORMAL style
            //style.Font.Name = styleSet.Normal.Font.Name;
            //style.Font.Size = styleSet.Normal.Font.Size;
            //style.ParagraphFormat.SpaceBefore = styleSet.Normal.ParagraphFormat.SpaceBefore;
            //style.ParagraphFormat.SpaceAfter = styleSet.Normal.ParagraphFormat.SpaceAfter ;
            //style.ParagraphFormat.PageBreakBefore = styleSet.Normal.ParagraphFormat.PageBreakBefore;
            //style.ParagraphFormat.Alignment = styleSet.Normal.ParagraphFormat.Alignment;

            ObjectHelper.MapProperties(styleSet.Normal, style);
            ObjectHelper.MapProperties(styleSet.Normal.Font, style.Font);
            ObjectHelper.MapProperties(styleSet.Normal.ParagraphFormat, style.ParagraphFormat);
            ObjectHelper.MapProperties(styleSet.Normal.ParagraphFormat.Borders.Left, style.ParagraphFormat.Borders.Left);
            ObjectHelper.MapProperties(styleSet.Normal.ParagraphFormat.Borders.Right, style.ParagraphFormat.Borders.Right);
            ObjectHelper.MapProperties(styleSet.Normal.ParagraphFormat.Borders.Top, style.ParagraphFormat.Borders.Top);
            ObjectHelper.MapProperties(styleSet.Normal.ParagraphFormat.Borders.Bottom, style.ParagraphFormat.Borders.Bottom);

            //AddStyle(styleSet.Normal);
            AddStyle(styleSet.Footer);

            AddStyle(styleSet.NormalTable);
            AddStyle(styleSet.Bullet1);
            AddStyle(styleSet.ChartCell);
            AddStyle(styleSet.ChartTitle);
            AddStyle(styleSet.ChartYLabel);
            AddStyle(styleSet.Code);
            AddStyle(styleSet.Details);

            AddStyle(styleSet.Header);
            AddStyle(styleSet.Heading1);
            AddStyle(styleSet.Heading2);
            AddStyle(styleSet.Heading3);
            AddStyle(styleSet.Heading4);
            AddStyle(styleSet.NoHeading1);
            AddStyle(styleSet.Table);
            AddStyle(styleSet.Title);
            AddStyle(styleSet.Toc1);
            AddStyle(styleSet.Toc2);
            AddStyle(styleSet.Toc3);
            AddStyle(styleSet.Toc4);
            AddStyle(styleSet.TocHeading1);



        }

        #endregion


        #region Private properties

        private Document _document = new Document();
        private Section _content;
        private Section _toc;

        private PageSetup _ps;

        private const double WidthDateTime = 8;
        private const double WidthDouble = 10;
        private const double WidthInteger = 8;

        #endregion



        public Table Table { get; private set; }





        public int Increment { get; set; }




        public Color AlternateBackColor { get; set; }
        public Color BackColor { get; set; }
        public Color TableBorderColor { get; set; }

        // Farben für Stylesheets wie "wr_cell_h1"

        public Color ShadingRisk2Color { get; set; }
        public Color ShadingRisk1Color { get; set; }
        public Color ShadingH3Color { get; set; }
        public Color ShadingH2Color { get; set; }
        public Color ShadingH1Color { get; set; }



        public bool AddPageBreakIfNecessary { get; set; }

        public string BackgroundImagePath { get; set; }


        

        private void LoadDefaults()
        {
            AddPageBreakIfNecessary = false;

            BackColor = Colors.LightSteelBlue;
            TableBorderColor = Colors.DarkGray;
            AlternateBackColor = Colors.White;
            ShadingH1Color = Colors.GreenYellow;
            ShadingH2Color = Colors.YellowGreen;
            ShadingH3Color = Colors.Gold;
            ShadingRisk1Color = Colors.Red;
            ShadingRisk2Color = Colors.Orange;
            Increment = 21;
        }



        /// <summary>
        /// Get the current width of the page.
        /// </summary>
        public double Width
        {
            get
            {
                double w;

                if (_content.PageSetup.Orientation == Orientation.Landscape)
                {
                    w = _content.PageSetup.PageHeight.Centimeter - _content.PageSetup.RightMargin.Centimeter -
                            _content.PageSetup.LeftMargin.Centimeter;

                }
                else
                {
                    w = _content.PageSetup.PageWidth.Centimeter - _content.PageSetup.RightMargin.Centimeter -
                            _content.PageSetup.LeftMargin.Centimeter;
                }

                return w;
            }
        }


        /// <summary>
        /// Save Pdf to a file
        /// </summary>
        /// <param name="fileName">Full path for pdf file's destination</param>
        /// <param name="showPdfFile">Show Pdf-File in a viewer</param>
        public void RenderToPdf(string fileName, bool showPdfFile)
        {

            var renderer = new PdfDocumentRenderer(true) { Document = _document };
            renderer.RenderDocument();

            // Save the document...
            renderer.PdfDocument.Save(fileName);
            
            if (!showPdfFile) return;

            // ...and start a viewer.
            var p = new Process
            {
                StartInfo = new ProcessStartInfo(fileName)
                {
                    UseShellExecute = true
                }
            };
            p.Start();

        }

        /// <summary>
        /// Save Pdf to a stream
        /// </summary>
        /// <param name="stream">Stream</param>
        public void RenderToPdf(Stream stream)
        {

            var renderer = new PdfDocumentRenderer(true) { Document = _document };
            renderer.RenderDocument();
            renderer.PdfDocument.Save(stream);
        }

        /// <summary>
        /// Set general document information
        /// </summary>
        /// <param name="title">Title of the file</param>
        /// <param name="subject">Subject of the file</param>
        /// <param name="author">Author of the file</param>
        public void SetDocInfo(string title, string subject, string author)
        {
            _document.Info.Title = title;
            _document.Info.Subject = subject;
            _document.Info.Author = author;
        }

        /// <summary>
        /// Page setup for content section
        /// </summary>
        /// <param name="pageFormat">page format</param>
        /// <param name="orientation">Orientation</param>
        /// <param name="pageHeight"></param>
        /// <param name="leftMargin">Left margin</param>
        /// <param name="topMargin">Top margin</param>
        /// <param name="rightMargin">Right margin</param>
        /// <param name="bottomMargin">Bottom margin</param>
        /// <param name="pageWidth"></param>
        /// <returns>Page setup for individual settings</returns>
        public PageSetup SetPage(PageFormat pageFormat, Orientation orientation, Unit pageWidth, Unit pageHeight, Unit leftMargin, Unit topMargin,
                                 Unit rightMargin, Unit bottomMargin)
        {
            _ps = new PageSetup
            {
                Orientation = orientation,
                PageFormat = pageFormat,
                LeftMargin = leftMargin,
                TopMargin = topMargin,
                RightMargin = rightMargin,
                BottomMargin = bottomMargin,
                PageHeight = pageHeight,
                PageWidth = pageWidth
            };

            return _ps;
        }



        public void CreateDocument()
        {
            // Create a new MigraDoc document


            //Styles.DefineStyles(document);

            //Cover.DefineCover(document);
            //TableOfContents.DefineTableOfContents(document);

            //DefineContentSection(document);

            //Paragraphs.DefineParagraphs(document);
            //Tables.DefineTables(document);
            //Charts.DefineCharts(document);


        }

        /// <summary>
        /// Add a ney style based on style "Normal"
        /// </summary>
        /// <param name="styleName">Name of the new style</param>
        /// <returns>New style object</returns>
        public Style AddStyle(string styleName)
        {
            // Check, if style already exists
            var i = _document.Styles.GetIndex(styleName);
            return i >= 0 ? _document.Styles[styleName] : _document.Styles.AddStyle(styleName, "Normal");

            // Create new style
        }

        /// <summary>
        /// Add a style to document
        /// </summary>
        /// <param name="style"></param>
        /// <returns></returns>
        public Style AddStyle(Style style)
        {
            if (style == null) return null;

            _document.Styles.Add(style);

            return style;
        }



        /// <summary>
        /// Add a ney style based on an other style
        /// </summary>
        /// <param name="styleName">Style name</param>
        /// <param name="baseStyleName">name of the style, the new one is based on</param>
        /// <returns></returns>
        public Style AddStyle(string styleName, string baseStyleName)
        {
            // Check, if style already exists
            var i = _document.Styles.GetIndex(styleName);
            return i >= 0 ? _document.Styles[styleName] : _document.Styles.AddStyle(styleName, baseStyleName);
        }

        /// <summary>
        /// Add a content section to the document
        /// </summary>
        public void CreateContentSection()
        {
            //if (_content == null)
            //{
            _content = _document.AddSection();
            _content.PageSetup = _ps.Clone();

            AddHeaderInternal(_content);
            AddFooterInternal(_content);

            //}
        }


        /// <summary>
        /// Add a toc section to the document
        /// </summary>
        public void CreateTocSection(string heading)
        {
            _toc = _document.AddSection();

            AddHeaderInternal(_toc);
            AddFooterInternal(_toc);

            Paragraph p;

            if (!string.IsNullOrEmpty(heading))
            {
                p = _toc.AddParagraph(heading, "Title");
            }

            p = _toc.AddParagraph("Inhaltsverzeichnis", "TocHeading1");
            p.AddBookmark("Inhalt");

            _content = _toc;
            _toc.PageSetup = _ps.Clone();
        }

        /// <summary>
        /// Add a paragraph to the content section
        /// </summary>
        /// <param name="text"></param>
        public void AddParagraph(string text)
        {
            var paragraph = new Paragraph();
            paragraph.AddText(text ?? "");
            _content.Add(paragraph);


        }

        /// <summary>
        /// Add a paragraph to the content section
        /// </summary>
        /// <param name="text"></param>
        /// <param name="styleName"></param>
        public void AddParagraph(string text, string styleName)
        {
            var paragraph = new Paragraph();

            paragraph.AddText(text ?? "");

            paragraph.Style = styleName;
            _content.Add(paragraph);

            if (_toc == null || !styleName.ToLower().StartsWith("heading")) return;


            var p = _toc.AddParagraph();

            switch (styleName.ToLower())
            {
                case "heading1":
                    p.Style = "TOC1";
                    break;
                case "heading2":
                    p.Style = "TOC2";
                    break;
                case "heading3":
                    p.Style = "TOC3";
                    break;
                case "heading4":
                    p.Style = "TOC4";
                    break;
                default:
                    return;
            }

            paragraph.AddBookmark(text);

            var hyperlink = p.AddHyperlink(text);
            hyperlink.AddText(text + "\t");
            hyperlink.AddPageRefField(text);
        }

        /// <summary>
        /// Add a paragraph object to the content
        /// </summary>
        /// <param name="paragraph"></param>
        public void AddParagraph(Paragraph paragraph)
        {
            _content.Add(paragraph);
        }

        private string _footerText;
        private string _footerStyleName;

        /// <summary>
        /// Set a footer text for content and toc
        /// Use &#60;&#60;page&#62;&#62; and &#60;&#60;pages&#62;&#62; fur current page number and number of pages in document
        /// </summary>
        /// <param name="text"></param>
        /// <param name="styleName"></param>
        public void SetFooter(string text, string styleName = "Footer")
        {
            _footerText = text;
            _footerStyleName = styleName;
        }

        private void AddFooterInternal(Section section)
        {
            if (section == null) return;

            var paragraph = new Paragraph();

            var text = _footerText;

            if (text.Contains("<<page>>"))
            {

                var vorher = text.Substring(0, text.IndexOf("<<page>>", StringComparison.Ordinal));
                var nachher = text.Substring(text.IndexOf("<<page>>", StringComparison.Ordinal) + 8,
                    text.Length - text.IndexOf("<<page>>", StringComparison.Ordinal) - 8);
                paragraph.AddText(vorher);
                paragraph.AddPageField();

                if (nachher.Contains("<<pages>>"))
                {
                    vorher = nachher.Substring(0, nachher.IndexOf("<<pages>>", StringComparison.Ordinal));
                    nachher = nachher.Substring(nachher.IndexOf("<<pages>>", StringComparison.Ordinal) + 9,
                        nachher.Length - nachher.IndexOf("<<pages>>", StringComparison.Ordinal) - 9);

                    paragraph.AddText(vorher);
                    paragraph.AddNumPagesField();
                }
                paragraph.AddText(nachher);


            }
            else
            {
                paragraph.AddText(_footerText);
            }

            paragraph.Style = _footerStyleName;
            section.Footers.Primary.Add(paragraph);
        }

        public void SetHeader(string text, string styleName = "Header")
        {

            SetHeader(text, styleName, null);

        }


        private string _headerText;
        private string _headerStyleName;
        private string _headerLogoPath;


        public void SetHeader(string text, string styleName, string logoPath)
        {
            _headerText = text;
            _headerStyleName = styleName;
            _headerLogoPath = logoPath;
        }

        private void AddHeaderInternal(Section section)
        {
            if (section == null) return;

            if (!string.IsNullOrEmpty(BackgroundImagePath) && File.Exists(BackgroundImagePath))
            {
                var image = section.Headers.Primary.AddImage(BackgroundImagePath);
                image.Height = "21cm";
                image.Width = "29.7cm";
                image.RelativeVertical = RelativeVertical.Page;
                image.RelativeHorizontal = RelativeHorizontal.Page;
                image.WrapFormat.Style = WrapStyle.Through;
            }

            var paragraph = new Paragraph
            {
                Format =
                {
                    Alignment = ParagraphAlignment.Left
                }

            };

            paragraph.Format.TabStops.ClearAll();

            var width = (_ps.Orientation== Orientation.Landscape) ? Unit.FromCentimeter(_ps.PageHeight.Centimeter - 
                        _ps.LeftMargin.Centimeter - 
                        _ps.RightMargin.Centimeter) :
                         Unit.FromCentimeter(_ps.PageWidth.Centimeter -
                                    _ps.LeftMargin.Centimeter -
                                    _ps.RightMargin.Centimeter);

            paragraph.Format.AddTabStop(width, TabAlignment.Right);


            if (!string.IsNullOrEmpty(_headerLogoPath))
            {
                var image = paragraph.AddImage(_headerLogoPath);
                image.Height = Unit.FromCentimeter(0.5);
                image.RelativeVertical = RelativeVertical.Line;
                image.RelativeHorizontal = RelativeHorizontal.Margin;
                image.Left = ShapePosition.Left;
                image.Top = ShapePosition.Center;
                image.LockAspectRatio = true;
                image.WrapFormat.Style = WrapStyle.Through;
            }

            paragraph.AddText("\t"+_headerText);

            paragraph.Style = _headerStyleName;
            section.Headers.Primary.Add(paragraph);



        }


        /// <summary>
        /// Add a table to the content
        /// </summary>
        /// <param name="dt">Data to show in the table</param>
        /// <param name="heading">Heading for the table</param>
        /// <param name="headingStyleName">Style name for the heading</param>
        /// <param name="additionalInfos"></param>
        /// <param name="additionalInfosStyleName"></param>
        /// <param name="width"></param>
        /// <param name="tableStyle">Name of the style to use for table formatting (not all properties supported)</param>
        public void AddTable(DataTable dt, string heading, string headingStyleName, string additionalInfos, string additionalInfosStyleName, double width = 0, string tableStyle = "NormalTable")
        {

            if (Math.Abs(width) < 0.000001) width = Width;

            if (!string.IsNullOrEmpty(heading))
            {
                AddParagraph(heading, headingStyleName);
            }

            if (!string.IsNullOrEmpty(additionalInfos))
            {
                AddParagraph(additionalInfos, additionalInfosStyleName);
            }



            var style = _document.Styles[tableStyle];

            //frame.FillFormat.Color = Colors.White;
            var table = _content.AddTable();
            table.LeftPadding = 2;
            table.Borders.Width = 0.5;
            table.Borders.Color = TableBorderColor;

            var colCount = dt.Columns.Count;


            var startCol = 1;
            var format = new string[dt.Columns.Count];

            var usedWidth = 0D;
            var colCountNotUsed = 0;

            var fontSize = Unit.FromCentimeter(style.Font.Size.Centimeter / 40.0);

            // Ermittle Breite der Nicht-Text-Spalten und Anzahl der Text-Spalten
            for (var i = 1; i <= colCount; i++)
            {
                var col = dt.Columns[i - 1];

                if (col.ColumnName.ToLower() == "cssstyle") continue;

                var t = col.DataType.ToString().Replace("System.", "").ToLower();

                switch (t)
                {
                    case "datetime":
                        usedWidth += WidthDateTime * fontSize;
                        break;
                    case "decimal":
                    case "double":
                    case "single":
                    case "float":
                        usedWidth += WidthDouble * fontSize;
                        break;
                    case "int":
                    case "int16":
                    case "int32":
                    case "int64":
                        usedWidth += WidthInteger * fontSize;
                        break;
                    default:
                        colCountNotUsed++;
                        break;
                }
            }

            // Errechne dann die zur Verfügung stehende maximale Breite der Text-Spalten
            var widthText = colCountNotUsed > 0 ? Math.Round((Width - usedWidth) / colCountNotUsed, 1) - 0.1 : 2.0;

            if (widthText > 7.0) widthText = 7.0;

            for (var i = 1; i <= colCount; i++)
            {
                var col = dt.Columns[i - 1];

                if (col.ColumnName.ToLower() == "cssstyle")
                {
                    startCol = 2;
                    continue;
                }

                double colWidth;

                var column = table.AddColumn();
                column.Borders.Color = TableBorderColor;

                var t = col.DataType.ToString().Replace("System.", "").ToLower();
                switch (t)
                {
                    case "datetime":
                        colWidth = WidthDateTime*fontSize;
                        format[i - 1] = @"dd.MM.yyyy";
                        break;
                    case "decimal":
                    case "double":
                    case "single":
                        colWidth = WidthDouble * fontSize;
                        column.Format.Alignment = ParagraphAlignment.Right;
                        format[i - 1] = @"#,##0.00";
                        break;
                    case "int":
                    case "int16":
                    case "int32":
                    case "int64":
                        colWidth = WidthInteger * fontSize;
                        column.Format.Alignment = ParagraphAlignment.Right;
                        format[i - 1] = @"#,##0";
                        break;
                    default:
                        colWidth = widthText;
                        column.Format.Alignment = ParagraphAlignment.Left;
                        break;
                }

                column.Width = Unit.FromCentimeter(colWidth);
            }



            var korr = startCol == 2 ? 1 : 0;

            // Kopfzeile schreiben
            var header = table.AddRow();
            header.Shading.Color = BackColor;
            header.Format.Font.Color = Colors.Black;
            header.Format.Font.Size = style.Font.Size;
            header.Format.Font.Name = style.Font.Name;

            for (var i = 1; i <= table.Columns.Count; i++)
            {
                var cell = header.Cells[i - 1];
                var p = cell.AddParagraph(dt.Columns[i - 1 + korr].ColumnName);
                p.Format.Font.Size = style.Font.Size;
                p.Format.Font.Name = style.Font.Name;
                p.Format.Font.Bold = true;
            }

            // Inhaltszeilen schreiben
            var shadow = false;

            foreach (DataRow r in dt.Rows)

            //for (var zeile = schleife * Increment; zeile < (schleife + 1) * Increment; zeile++)
            {
                var row = table.AddRow();
                //row.KeepWith = 2;
                var css = "";
                if (startCol == 2) css = r[0].ToString();

                Color shadingColor; 

                if (string.IsNullOrEmpty(css))
                {
                    shadingColor = shadow ? BackColor : AlternateBackColor;
                }
                else
                {
                    switch (css.ToLower())
                    {
                        case "wr_cell_h1":
                            shadingColor = ShadingH1Color;
                            break;
                        case "wr_cell_h2":
                            shadingColor = ShadingH2Color;
                            break;
                        case "wr_cell_h3":
                            shadingColor = ShadingH3Color;
                            break;
                        case "wr_cell_risk1":
                            shadingColor = ShadingRisk1Color;
                            break;
                        case "wr_cell_risk2":
                            shadingColor = ShadingRisk2Color;
                            break;
                        default:
                            shadingColor = shadow ? BackColor : AlternateBackColor;
                            break;
                    }
                }

                row.Shading.Color = shadingColor;

                for (var i = 0; i < table.Columns.Count; i++)
                {
                    var cell = row.Cells[i];

                    if (string.IsNullOrEmpty(format[i + korr]))
                    {
                        var p = cell.AddParagraph(r[i + korr].ToString());
                        p.Format.Font.Size = style.Font.Size;
                        p.Format.Font.Name = style.Font.Name;
                        //p.Format.Shading.Color = shadingColor;
                    }
                    else
                    {
                        if (format[i + korr].ToLower().Contains("yy"))
                        {
                            var z = r[i + korr].ToString();
                            if (!string.IsNullOrEmpty(z))
                            {
                                var p = cell.AddParagraph(Convert.ToDateTime(z).ToString(format[i + korr]));
                                p.Format.Font.Size = style.Font.Size;
                                p.Format.Font.Name = style.Font.Name;
                                //p.Format.Shading.Color = shadingColor;
                            }
                        }
                        else
                        {
                            var z = r[i + korr].ToString();
                            if (!string.IsNullOrEmpty(z))
                            {
                                var p = cell.AddParagraph(Convert.ToDouble(z).ToString(format[i + korr]));
                                p.Format.Font.Size = style.Font.Size;
                                p.Format.Font.Name = style.Font.Name;
                               // p.Format.Shading.Color = shadingColor;

                            }
                        }
                    }
                }

                shadow = !shadow;
            }

            //var widthTable = table.Columns.Cast<Column>().Aggregate(0D, (current, t) => current + t.Width.Centimeter);
            //frame.Width = Unit.FromCentimeter(widthTable);

        }


        /// <summary>
        /// Add a table in a separate frame
        /// </summary>
        /// <param name="dt">Data to show in the table</param>
        /// <param name="heading">Heading for the table</param>
        /// <param name="headingStyleName">Style name for the heading</param>
        /// <param name="additionalInfos"></param>
        /// <param name="additionalInfosStyleName"></param>
        /// <param name="width"></param>
        public void AddTableFrame(DataTable dt, string heading, string headingStyleName, string additionalInfos = null, string additionalInfosStyleName = null, double width = 0)
        {

            if (width < 0.000001) width = Width;

            if (Math.Abs(width) < 0.000001) width = Width;

            var anzahlSchleifen = dt.Rows.Count / Increment + 1;

            if (dt.Rows.Count > Increment / 2 && AddPageBreakIfNecessary) _content.AddPageBreak();

            if (!string.IsNullOrEmpty(heading))
            {
                AddParagraph(heading, headingStyleName);
            }

            if (!string.IsNullOrEmpty(additionalInfos))
            {
                AddParagraph(additionalInfos, additionalInfosStyleName);
            }

            for (var schleife = 0; schleife < anzahlSchleifen; schleife++)
            {
                if (schleife > 0)
                {
                    _content.AddPageBreak();
                    if (!string.IsNullOrEmpty(heading))
                    {
                        var p = _content.AddParagraph(heading + " (Forts.)");
                        p.Style = headingStyleName;
                    }
                }

                var frame = _content.AddTextFrame();
                frame.Height = Unit.FromCentimeter(6F);
                frame.Width = Unit.FromCentimeter(width);
                frame.Left = ShapePosition.Center;

                CreateTable(dt, schleife, frame);
            }
        }

        /// <summary>
        /// Add a definition list with left and right column
        /// </summary>
        /// <param name="dt">DataTable with two columns</param>
        /// <param name="style1">Name of the style to use for left column</param>
        /// <param name="style2">Name of the style to use for right column</param>
        /// <param name="columnWidth1">Column width column 1 in percent</param>
        public void AddDefinitionList(DataTable dt, string style1 = null, string style2 = null, double columnWidth1 = 0.3)
        {

            const double borderWidth = 0;

            var table = _content.AddTable();


            table.TopPadding = 4;
            table.Borders.Width = borderWidth;
            table.BottomPadding = 4;

            var column1 = table.AddColumn(Unit.FromCentimeter(columnWidth1 * Width));
            column1.Format.Alignment = ParagraphAlignment.Left;
            column1.RightPadding = 2;

            var column2 = table.AddColumn(Unit.FromCentimeter((1 - columnWidth1) * Width));
            column2.Format.Alignment = ParagraphAlignment.Left;
            //column2.LeftPadding = 2;

            foreach (DataRow r in dt.Rows)
            {
                var row = table.AddRow();

                row.Borders.Width = borderWidth;
                row.BottomPadding = 2;

                var cell1 = row.Cells[0];
                cell1.Borders.Width = borderWidth;
                var p1 = cell1.AddParagraph(r[0].ToString());

                if (!string.IsNullOrEmpty(style1)) p1.Style = style1;


                var cell2 = row.Cells[1];
                cell2.Borders.Width = borderWidth;


                var p2 = cell2.AddParagraph(r[1].ToString());
                if (!string.IsNullOrEmpty(style2)) p2.Style = style2;
            }
        }


        private void CreateTable(DataTable dt, int schleife, TextFrame frame, double borderWidth = 0.5, string tableStyle = "NormalTable")
        {

            var style = _document.Styles[tableStyle];



            //frame.FillFormat.Color = Colors.White;
            var table = frame.AddTable();

            table.Borders.Width = borderWidth;
            table.BottomPadding = 0;
            table.TopPadding = 0;

            if (borderWidth > 0) table.Borders.Color = TableBorderColor;

            var colCount = dt.Columns.Count;


            var startCol = 1;
            var format = new string[dt.Columns.Count];

            var usedWidth = 0D;
            var colCountNotUsed = 0;

            var fontSize = Unit.FromCentimeter(style.Font.Size.Centimeter / 40.0);

            // Ermittle Breite der Nicht-Text-Spalten und Anzahl der Text-Spalten
            for (var i = 1; i <= colCount; i++)
            {
                var col = dt.Columns[i - 1];

                if (col.ColumnName.ToLower() == "cssstyle") continue;

                var t = col.DataType.ToString().Replace("System.", "").ToLower();

                switch (t)
                {
                    case "datetime":
                        usedWidth += WidthDateTime * fontSize;
                        break;
                    case "decimal":
                    case "double":
                    case "single":
                    case "float":
                        usedWidth += WidthDouble * fontSize;
                        break;
                    case "int":
                    case "int16":
                    case "int32":
                    case "int64":
                        usedWidth += WidthInteger * fontSize;
                        break;
                    default:
                        colCountNotUsed++;
                        break;
                }
            }

            // Errechne dann die zur Verfügung stehende maximale Breite der Text-Spalten
            var widthText = colCountNotUsed > 0 ? Math.Round((frame.Width.Centimeter - usedWidth) / colCountNotUsed, 1) - 0.1 : 2.0;

            if (widthText > 7.0) widthText = 7.0;

            for (var i = 1; i <= colCount; i++)
            {
                var col = dt.Columns[i - 1];

                if (col.ColumnName.ToLower() == "cssstyle")
                {
                    startCol = 2;
                    continue;
                }

                double width;

                var column = table.AddColumn();
                column.Borders.Color = TableBorderColor;

                var t = col.DataType.ToString().Replace("System.", "").ToLower();
                switch (t)
                {
                    case "datetime":
                        width = WidthDateTime * fontSize;
                        format[i - 1] = @"dd.MM.yyyy";
                        break;
                    case "decimal":
                    case "double":
                    case "single":
                        width = WidthDouble * fontSize;
                        column.Format.Alignment = ParagraphAlignment.Right;
                        format[i - 1] = @"#,##0.00";
                        break;
                    case "int":
                    case "int16":
                    case "int32":
                    case "int64":
                        width = WidthInteger * fontSize;
                        column.Format.Alignment = ParagraphAlignment.Right;
                        format[i - 1] = @"#,##0";
                        break;
                    default:
                        width = widthText;
                        column.Format.Alignment = ParagraphAlignment.Left;
                        break;
                }

                column.Width = Unit.FromCentimeter(width);
            }



            var korr = startCol == 2 ? 1 : 0;

            // Kopfzeile schreiben
            var header = table.AddRow();
            header.Shading.Color = BackColor;
            header.Format.Font.Color = Colors.Black;
            header.Format.Font.Size = style.Font.Size-0.5;
            header.Format.Font.Name = style.Font.Name;
            header.Format.Font.Bold = true;
            header.Format.SpaceAfter = style.ParagraphFormat.SpaceAfter;
            header.Format.SpaceBefore = style.ParagraphFormat.SpaceBefore;

            for (var i = 1; i <= table.Columns.Count; i++)
            {
                var cell = header.Cells[i - 1];
                cell.AddParagraph(dt.Columns[i - 1 + korr].ColumnName);

            }

            // Inhaltszeilen schreiben
            var shadow = false;
            for (var zeile = schleife * Increment; zeile < (schleife + 1) * Increment; zeile++)
            {
                if (zeile >= dt.Rows.Count) break;

                var r = dt.Rows[zeile];
                var row = table.AddRow();

                row.BottomPadding = 0;
                row.TopPadding = 0;

                var css = "";
                if (startCol == 2) css = r[0].ToString();

                Color shadingColor;

                if (string.IsNullOrEmpty(css))
                {
                    shadingColor = shadow ? BackColor : AlternateBackColor;
                }
                else
                {
                    switch (css.ToLower())
                    {
                        case "wr_cell_h1":
                            shadingColor = ShadingH1Color;
                            break;
                        case "wr_cell_h2":
                            shadingColor = ShadingH2Color;
                            break;
                        case "wr_cell_h3":
                            shadingColor = ShadingH3Color;
                            break;
                        case "wr_cell_risk1":
                            shadingColor = ShadingRisk1Color;
                            break;
                        case "wr_cell_risk2":
                            shadingColor = ShadingRisk2Color;
                            break;
                        default:
                            shadingColor = shadow ? BackColor : AlternateBackColor;
                            break;
                    }
                }

                row.Shading.Color = shadingColor;

                for (var i = 0; i < table.Columns.Count; i++)
                {
                    var cell = row.Cells[i];
                    cell.Format.SpaceAfter = 0;
                    cell.Format.SpaceBefore = 0;
                    if (string.IsNullOrEmpty(format[i + korr]))
                    {
                        var p = cell.AddParagraph(r[i + korr].ToString().Trim());
                        p.Format.Font.Size = style.Font.Size;
                        p.Format.Font.Name = style.Font.Name;
                        p.Format.SpaceAfter = style.ParagraphFormat.SpaceAfter;
                        p.Format.SpaceBefore = style.ParagraphFormat.SpaceBefore;
                        p.Format.Shading.Color = shadingColor;
                    }
                    else
                    {
                        if (format[i + korr].ToLower().Contains("yy"))
                        {
                            var z = r[i + korr].ToString().Trim();
                            if (!string.IsNullOrEmpty(z))
                            {
                                var p = cell.AddParagraph(Convert.ToDateTime(z).ToString(format[i + korr]));
                                p.Format.Font.Size = style.Font.Size;
                                p.Format.Font.Name = style.Font.Name;
                                p.Format.SpaceAfter = style.ParagraphFormat.SpaceAfter;
                                p.Format.SpaceBefore = style.ParagraphFormat.SpaceBefore;
                                p.Format.Shading.Color = shadingColor;
                            }
                        }
                        else
                        {
                            var z = r[i + korr].ToString();
                            if (!string.IsNullOrEmpty(z))
                            {
                                var p = cell.AddParagraph(Convert.ToDouble(z).ToString(format[i + korr]));
                                p.Format.Font.Size = style.Font.Size;
                                p.Format.Font.Name = style.Font.Name;
                                p.Format.SpaceAfter = style.ParagraphFormat.SpaceAfter;
                                p.Format.SpaceBefore = style.ParagraphFormat.SpaceBefore;
                                p.Format.Shading.Color = shadingColor;
                            }
                        }
                    }
                }

                shadow = !shadow;
            }

            var widthTable = table.Columns.Cast<Column>().Aggregate(0D, (current, t) => current + t.Width.Centimeter);
            frame.Width = Unit.FromCentimeter(widthTable);
        }





        /// <summary>
        /// Seitenumbruch in Text einfügen
        /// </summary>
        public void NewPage()
        {
            _content.AddPageBreak();
        }


        /// <summary>
        /// Defines the styles used in the document.
        /// </summary>
        public void DefineStyles()
        {
            //// Get the predefined style Normal.
            //Style style = document.Styles["Normal"];
            //// Because all styles are derived from Normal, the next line changes the 
            //// font of the whole document. Or, more exactly, it changes the font of
            //// all styles and paragraphs that do not redefine the font.
            //style.Font.Name = "Times New Roman";

            //// Heading1 to Heading9 are predefined styles with an outline level. An outline level
            //// other than OutlineLevel.BodyText automatically creates the outline (or bookmarks) 
            //// in PDF.

            //style = document.Styles["Heading1"];
            //style.Font.Name = "Tahoma";
            //style.Font.Size = 14;
            //style.Font.Bold = true;
            //style.Font.Color = Colors.DarkBlue;
            //style.ParagraphFormat.PageBreakBefore = true;
            //style.ParagraphFormat.SpaceAfter = 6;

            //style = document.Styles["Heading2"];
            //style.Font.Size = 12;
            //style.Font.Bold = true;
            //style.ParagraphFormat.PageBreakBefore = false;
            //style.ParagraphFormat.SpaceBefore = 6;
            //style.ParagraphFormat.SpaceAfter = 6;

            //style = document.Styles["Heading3"];
            //style.Font.Size = 10;
            //style.Font.Bold = true;
            //style.Font.Italic = true;
            //style.ParagraphFormat.SpaceBefore = 6;
            //style.ParagraphFormat.SpaceAfter = 3;

            //style = document.Styles[StyleNames.Header];
            //style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right);

            //style = document.Styles[StyleNames.Footer];
            //style.ParagraphFormat.AddTabStop("8cm", TabAlignment.Center);

            //// Create a new style called TextBox based on style Normal
            //style = document.Styles.AddStyle("TextBox", "Normal");
            //style.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
            //style.ParagraphFormat.Borders.Width = 2.5;
            //style.ParagraphFormat.Borders.Distance = "3pt";
            //style.ParagraphFormat.Shading.Color = Colors.SkyBlue;

            //// Create a new style called TOC based on style Normal
            //style = document.Styles.AddStyle("TOC", "Normal");
            //style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right, TabLeader.Dots);
            //style.ParagraphFormat.Font.Color = Colors.Blue;
        }





        ///// <summary>
        ///// Defines the cover page.
        ///// </summary>
        //public void DefineCover()
        //{
        //    //Section section = document.AddSection();

        //    //Paragraph paragraph = section.AddParagraph();
        //    //paragraph.Format.SpaceAfter = "3cm";

        //    _content.AddImage()

        //    Image image = _content.AddImage("../../images/Logo landscape.png");
        //    image.Width = "10cm";

        //    //paragraph = section.AddParagraph("A sample document that demonstrates the\ncapabilities of MigraDoc");
        //    //paragraph.Format.Font.Size = 16;
        //    //paragraph.Format.Font.Color = Colors.DarkRed;
        //    //paragraph.Format.SpaceBefore = "8cm";
        //    //paragraph.Format.SpaceAfter = "3cm";

        //    //paragraph = section.AddParagraph("Rendering date: ");
        //    //paragraph.AddDateField();
        //}

        /// <summary>
        /// Defines page setup, headers, and footers.
        /// </summary>
        public void DefineContentSection()
        {
            var section = _document.AddSection();
            section.PageSetup.OddAndEvenPagesHeaderFooter = false;
            section.PageSetup.StartingNumber = 1;

            var paragraph = new Paragraph();
            paragraph.AddText("Hallo Welt");

            section.Add(paragraph);

            //var header = section.Headers.Primary;
            //header.AddParagraph("\tOdd Page Header");

            //header = section.Headers.EvenPage;
            //header.AddParagraph("Even Page Header");

            //// Create a paragraph with centered page number. See definition of style "Footer".
            //var paragraph = new Paragraph();
            //paragraph.AddTab();
            //paragraph.AddPageField();

            //// Add paragraph to footer for odd pages.
            //section.Footers.Primary.Add(paragraph);
            //// Add clone of paragraph to footer for odd pages. Cloning is necessary because an object must
            //// not belong to more than one other object. If you forget cloning an exception is thrown.
            //section.Footers.EvenPage.Add(paragraph.Clone());
        }

        /// <summary>
        /// Defines the cover page.
        /// </summary>
        public static void DefineTableOfContents(Document document)
        {


            var section = document.LastSection;

            section.AddPageBreak();
            var paragraph = section.AddParagraph("Table of Contents");
            paragraph.Format.Font.Size = 14;
            paragraph.Format.Font.Bold = true;
            paragraph.Format.SpaceAfter = 24;
            paragraph.Format.OutlineLevel = OutlineLevel.Level1;

            paragraph = section.AddParagraph();
            paragraph.Style = "TOC";
            var hyperlink = paragraph.AddHyperlink("Paragraphs");
            hyperlink.AddText("Paragraphs\t");
            hyperlink.AddPageRefField("Paragraphs");

            paragraph = section.AddParagraph();
            paragraph.Style = "TOC";
            hyperlink = paragraph.AddHyperlink("Tables");
            hyperlink.AddText("Tables\t");
            hyperlink.AddPageRefField("Tables");

            paragraph = section.AddParagraph();
            paragraph.Style = "TOC";
            hyperlink = paragraph.AddHyperlink("Charts");
            hyperlink.AddText("Charts\t");
            hyperlink.AddPageRefField("Charts");

        }


        //public void DefineParagraphs()
        //{
        ////{
        ////    Paragraph paragraph = document.LastSection.AddParagraph("Paragraph Layout Overview", "Heading1");
        ////    paragraph.AddBookmark("Paragraphs");

        ////    DemonstrateAlignment(document);
        ////    DemonstrateIndent(document);
        ////    DemonstrateFormattedText(document);
        ////    DemonstrateBordersAndShading(document);
        //}

        //static void DemonstrateAlignment(Document document)
        //{
        //    //document.LastSection.AddParagraph("Alignment", "Heading2");

        //    //document.LastSection.AddParagraph("Left Aligned", "Heading3");

        //    //Paragraph paragraph = document.LastSection.AddParagraph();
        //    //paragraph.Format.Alignment = ParagraphAlignment.Left;
        //    //paragraph.AddText(FillerText.Text);

        //    //document.LastSection.AddParagraph("Right Aligned", "Heading3");

        //    //paragraph = document.LastSection.AddParagraph();
        //    //paragraph.Format.Alignment = ParagraphAlignment.Right;
        //    //paragraph.AddText(FillerText.Text);

        //    //document.LastSection.AddParagraph("Centered", "Heading3");

        //    //paragraph = document.LastSection.AddParagraph();
        //    //paragraph.Format.Alignment = ParagraphAlignment.Center;
        //    //paragraph.AddText(FillerText.Text);

        //    //document.LastSection.AddParagraph("Justified", "Heading3");

        //    //paragraph = document.LastSection.AddParagraph();
        //    //paragraph.Format.Alignment = ParagraphAlignment.Justify;
        //    //paragraph.AddText(FillerText.MediumText);
        //}

        //static void DemonstrateIndent(Document document)
        //{
        //    //document.LastSection.AddParagraph("Indent", "Heading2");

        //    //document.LastSection.AddParagraph("Left Indent", "Heading3");

        //    //Paragraph paragraph = document.LastSection.AddParagraph();
        //    //paragraph.Format.LeftIndent = "2cm";
        //    //paragraph.AddText(FillerText.Text);

        //    //document.LastSection.AddParagraph("Right Indent", "Heading3");

        //    //paragraph = document.LastSection.AddParagraph();
        //    //paragraph.Format.RightIndent = "1in";
        //    //paragraph.AddText(FillerText.Text);

        //    //document.LastSection.AddParagraph("First Line Indent", "Heading3");

        //    //paragraph = document.LastSection.AddParagraph();
        //    //paragraph.Format.FirstLineIndent = "12mm";
        //    //paragraph.AddText(FillerText.Text);

        //    //document.LastSection.AddParagraph("First Line Negative Indent", "Heading3");

        //    //paragraph = document.LastSection.AddParagraph();
        //    //paragraph.Format.LeftIndent = "1.5cm";
        //    //paragraph.Format.FirstLineIndent = "-1.5cm";
        //    //paragraph.AddText(FillerText.Text);
        //}

        //static void DemonstrateFormattedText(Document document)
        //{
        //    //document.LastSection.AddParagraph("Formatted Text", "Heading2");

        //    ////document.LastSection.AddParagraph("Left Aligned", "Heading3");

        //    //Paragraph paragraph = document.LastSection.AddParagraph();
        //    //paragraph.AddText("Text can be formatted ");
        //    //paragraph.AddFormattedText("bold", TextFormat.Bold);
        //    //paragraph.AddText(", ");
        //    //paragraph.AddFormattedText("italic", TextFormat.Italic);
        //    //paragraph.AddText(", or ");
        //    //paragraph.AddFormattedText("bold & italic", TextFormat.Bold | TextFormat.Italic);
        //    //paragraph.AddText(".");
        //    //paragraph.AddLineBreak();
        //    //paragraph.AddText("You can set the ");
        //    //FormattedText formattedText = paragraph.AddFormattedText("size ");
        //    //formattedText.Size = 15;
        //    //paragraph.AddText("the ");
        //    //formattedText = paragraph.AddFormattedText("color ");
        //    //formattedText.Color = Colors.Firebrick;
        //    //paragraph.AddText("the ");
        //    //formattedText = paragraph.AddFormattedText("font", new Font("Verdana"));
        //    //paragraph.AddText(".");
        //    //paragraph.AddLineBreak();
        //    //paragraph.AddText("You can set the ");
        //    //formattedText = paragraph.AddFormattedText("subscript");
        //    //formattedText.Subscript = true;
        //    //paragraph.AddText(" or ");
        //    //formattedText = paragraph.AddFormattedText("superscript");
        //    //formattedText.Superscript = true;
        //    //paragraph.AddText(".");
        //}

        //static void DemonstrateBordersAndShading(Document document)
        //{
        //    document.LastSection.AddPageBreak();
        //    document.LastSection.AddParagraph("Borders and Shading", "Heading2");

        //    document.LastSection.AddParagraph("Border around Paragraph", "Heading3");

        //    Paragraph paragraph = document.LastSection.AddParagraph();
        //    paragraph.Format.Borders.Width = 2.5;
        //    paragraph.Format.Borders.Color = Colors.Navy;
        //    paragraph.Format.Borders.Distance = 3;
        //    paragraph.AddText(FillerText.MediumText);

        //    document.LastSection.AddParagraph("Shading", "Heading3");

        //    paragraph = document.LastSection.AddParagraph();
        //    paragraph.Format.Shading.Color = Colors.LightCoral;
        //    paragraph.AddText(FillerText.Text);

        //    document.LastSection.AddParagraph("Borders & Shading", "Heading3");

        //    paragraph = document.LastSection.AddParagraph();
        //    paragraph.Style = "TextBox";
        //    paragraph.AddText(FillerText.MediumText);
        //}


        //public static void DefineTables(Document document)
        //{
        //    Paragraph paragraph = document.LastSection.AddParagraph("Table Overview", "Heading1");
        //    paragraph.AddBookmark("Tables");

        //    DemonstrateSimpleTable(document);
        //    DemonstrateAlignment(document);
        //    DemonstrateCellMerge(document);
        //}

        //public static void DemonstrateSimpleTable(Document document)
        //{
        //    document.LastSection.AddParagraph("Simple Tables", "Heading2");

        //    Table table = new Table();
        //    table.Borders.Width = 0.75;

        //    Column column = table.AddColumn(Unit.FromCentimeter(2));
        //    column.Format.Alignment = ParagraphAlignment.Center;

        //    table.AddColumn(Unit.FromCentimeter(5));

        //    Row row = table.AddRow();
        //    row.Shading.Color = Colors.PaleGoldenrod;
        //    Cell cell = row.Cells[0];
        //    cell.AddParagraph("Itemus");
        //    cell = row.Cells[1];
        //    cell.AddParagraph("Descriptum");

        //    row = table.AddRow();
        //    cell = row.Cells[0];
        //    cell.AddParagraph("1");
        //    cell = row.Cells[1];
        //    cell.AddParagraph(FillerText.ShortText);

        //    row = table.AddRow();
        //    cell = row.Cells[0];
        //    cell.AddParagraph("2");
        //    cell = row.Cells[1];
        //    cell.AddParagraph(FillerText.Text);

        //    table.SetEdge(0, 0, 2, 3, Edge.Box, BorderStyle.Single, 1.5, Colors.Black);

        //    document.LastSection.Add(table);
        //}

        //public static void DemonstrateAlignment(Document document)
        //{
        //    document.LastSection.AddParagraph("Cell Alignment", "Heading2");

        //    Table table = document.LastSection.AddTable();
        //    table.Borders.Visible = true;
        //    table.Format.Shading.Color = Colors.LavenderBlush;
        //    table.Shading.Color = Colors.Salmon;
        //    table.TopPadding = 5;
        //    table.BottomPadding = 5;

        //    Column column = table.AddColumn();
        //    column.Format.Alignment = ParagraphAlignment.Left;

        //    column = table.AddColumn();
        //    column.Format.Alignment = ParagraphAlignment.Center;

        //    column = table.AddColumn();
        //    column.Format.Alignment = ParagraphAlignment.Right;

        //    table.Rows.Height = 35;

        //    Row row = table.AddRow();
        //    row.VerticalAlignment = VerticalAlignment.Top;
        //    row.Cells[0].AddParagraph("Text");
        //    row.Cells[1].AddParagraph("Text");
        //    row.Cells[2].AddParagraph("Text");

        //    row = table.AddRow();
        //    row.VerticalAlignment = VerticalAlignment.Center;
        //    row.Cells[0].AddParagraph("Text");
        //    row.Cells[1].AddParagraph("Text");
        //    row.Cells[2].AddParagraph("Text");

        //    row = table.AddRow();
        //    row.VerticalAlignment = VerticalAlignment.Bottom;
        //    row.Cells[0].AddParagraph("Text");
        //    row.Cells[1].AddParagraph("Text");
        //    row.Cells[2].AddParagraph("Text");
        //}

        //public static void DemonstrateCellMerge(Document document)
        //{
        //    document.LastSection.AddParagraph("Cell Merge", "Heading2");

        //    Table table = document.LastSection.AddTable();
        //    table.Borders.Visible = true;
        //    table.TopPadding = 5;
        //    table.BottomPadding = 5;

        //    Column column = table.AddColumn();
        //    column.Format.Alignment = ParagraphAlignment.Left;

        //    column = table.AddColumn();
        //    column.Format.Alignment = ParagraphAlignment.Center;

        //    column = table.AddColumn();
        //    column.Format.Alignment = ParagraphAlignment.Right;

        //    table.Rows.Height = 35;

        //    Row row = table.AddRow();
        //    row.Cells[0].AddParagraph("Merge Right");
        //    row.Cells[0].MergeRight = 1;

        //    row = table.AddRow();
        //    row.VerticalAlignment = VerticalAlignment.Bottom;
        //    row.Cells[0].MergeDown = 1;
        //    row.Cells[0].VerticalAlignment = VerticalAlignment.Bottom;
        //    row.Cells[0].AddParagraph("Merge Down");

        //    table.AddRow();
        //}


        //public static void DefineCharts(Document document)
        //{
        //    Paragraph paragraph = document.LastSection.AddParagraph("Chart Overview", "Heading1");
        //    paragraph.AddBookmark("Charts");

        //    document.LastSection.AddParagraph("Sample Chart", "Heading2");

        //    Chart chart = new Chart();
        //    chart.Left = 0;

        //    chart.Width = Unit.FromCentimeter(16);
        //    chart.Height = Unit.FromCentimeter(12);
        //    Series series = chart.SeriesCollection.AddSeries();
        //    series.ChartType = ChartType.Column2D;
        //    series.Add(new double[] { 1, 17, 45, 5, 3, 20, 11, 23, 8, 19 });
        //    series.HasDataLabel = true;

        //    series = chart.SeriesCollection.AddSeries();
        //    series.ChartType = ChartType.Line;
        //    series.Add(new double[] { 41, 7, 5, 45, 13, 10, 21, 13, 18, 9 });

        //    XSeries xseries = chart.XValues.AddXSeries();
        //    xseries.Add("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N");

        //    chart.XAxis.MajorTickMark = TickMarkType.Outside;
        //    chart.XAxis.Title.Caption = "X-Axis";

        //    chart.YAxis.MajorTickMark = TickMarkType.Outside;
        //    chart.YAxis.HasMajorGridlines = true;

        //    chart.PlotArea.LineFormat.Color = Colors.DarkGray;
        //    chart.PlotArea.LineFormat.Width = 1;

        //    document.LastSection.Add(chart);
        //}



        //public void DefineCharts()
        //{
        //    _document.LastSection.AddParagraph("Sample Chart", "Heading2");

        //    var w = new PdfChart();
        //    w.Left = 5;
        //    w.TypeOfChart = ChartType.Line;
        //    w.AddSeries(new double[] { 1, 17, 45, 5, 3, 20, 11, 23, 8, 19 }, "Reihe 1");
        //    w.AddSeries(new double[] { 41, 7, 5, 45, 13, 10, 21, 13, 18, 9 }, "Reihe 2");
        //    w.XSeries("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N");
        //    _document.LastSection.Add(w.Draw());

        //}

        public void AddChart(Chart chart)
        {
            _content.Add(chart);
        }

        public void TableStart()
        {
            TableStart("NormalTable");
        }

        public void TableStart(string style)
        {
            var p = _content.AddParagraph();
            p.Style = style;
            p.AddText(" \t\t\t");
            Table = _content.AddTable();
            Table.Borders.Visible = false;
            Table.BottomPadding = 0.3F;
            //_table.Style = "ChartTable";
            Table.TopPadding = 9;

            _content.AddParagraph();

        }

        public void TableAddColumn(ParagraphAlignment alignment, double width)
        {
            var column = Table.AddColumn();
            column.Format.Alignment = alignment;
            column.Width = Unit.FromCentimeter(width);
        }


        public void TableEnd()
        {


        }

        public Row TableAddRow()
        {
            return Table.AddRow();
        }

        public void TableSetContent(int column, int row, string content)
        {

            if (string.IsNullOrEmpty(content)) content = "";

            var cell = Table.Rows[row].Cells[column];
            cell.AddParagraph(content);
        }

        public void TableSetContent(int column, int row, Chart chart)
        {
            Table.Rows[row][column].Add(chart);

        }



        public void TableSetContent(int column, int row, string chartImagePath, double width, double height)
        {

            var t = Table.Rows[row][column];

            var image = t.AddImage(chartImagePath);

            image.Width = Unit.FromCentimeter(width);
            image.Height = Unit.FromCentimeter(height);
            //image.Left = 0;

        }

        public void AddImage(string fileName, double width, double height)
        {
            var frame = _content.AddTextFrame();
            frame.Height = Unit.FromCentimeter(height);
            frame.Width = Unit.FromCentimeter(width);
            frame.Left = ShapePosition.Center;
            var image = frame.AddImage(fileName);
            image.Width = frame.Width;
            image.Height = frame.Height;
            image = null;
        }

        public void TableSetContentSmallTable(int column, int row, DataTable data, string heading, double width, double height = 6F)
        {
            var t = Table.Rows[row][column];

            if (!string.IsNullOrEmpty(heading))
            {
                var p = t.AddParagraph(heading);
                p.Style = "NoHeading1";
            }

            var frame = t.AddTextFrame();
            frame.Height = Unit.FromCentimeter(height);
            frame.Left = ShapePosition.Center;
            frame.Width = Unit.FromCentimeter(width);

            CreateTable(data, 0, frame);
        }

        public void AddTable<T>(IList<T> items, IList<PdfTableField> listFields, bool showheader)
        {

            var myObjectType = typeof(T);

            var fieldInfo = myObjectType.GetProperties().ToList();



            // Header

            var p = _content.AddParagraph();
            p.Style = "NormalTable";
            Table = _content.AddTable();
            Table.Borders.Visible = false;
            //_table.Style = "ChartTable";
            Table.TopPadding = 9;
            Table.Rows.LeftIndent = Unit.FromCentimeter(1);

            foreach (var field in listFields)
            {
                var columnDef = Table.AddColumn(Unit.FromCentimeter(field.Width));
                switch (field.TextAlign)
                {
                    case PdfTextAlignment.Center:
                        columnDef.Format.Alignment = ParagraphAlignment.Center;
                        break;
                    case PdfTextAlignment.Right:
                        columnDef.Format.Alignment = ParagraphAlignment.Right;
                        break;
                    default:
                        columnDef.Format.Alignment = ParagraphAlignment.Left;
                        break;
                }
            }


            // header anzeigen
            var column = 0;
            if (showheader)
            {
                var header = Table.AddRow();

                foreach (var field in listFields)
                {
                    var cell = header.Cells[column];
                    cell.AddParagraph(field.Header);
                    column++;
                }
            }

            // Show data
            foreach (var item in items)
            {
                var row = Table.AddRow();

                column = 0;
                foreach (var field in listFields)
                {
                    var f = field;
                    var info = fieldInfo.FirstOrDefault(x => x.Name == f.Name);

                    if (info == null) continue;

                    string value;

                    if (string.IsNullOrEmpty(field.Format))
                    {
                        value = info.GetValue(item, null).ToString();
                    }
                    else
                    {
                        var x = "{0:" + field.Format + "}";
                        value = string.Format(x, info.GetValue(item, null));

                    }

                    var cell = row.Cells[column];
                    cell.AddParagraph(value);
                    column++;
                }
            }

        }

        /// <summary>
        /// Add a HTML code
        /// </summary>
        /// <param name="html"></param>
        public void AddHtml(string html)
        {
            if (string.IsNullOrEmpty(html)) html = "";
            if (!html.Contains("<"))
            {
                AddParagraph(html);
                return;
            }


            html = html.Replace("&nbsp;", " ").Replace("<br />", "\r\n").Replace("<br/>", "\r\n").Replace("\r\n\r\n", "\r\n");

            var startTag = html.IndexOf("<", StringComparison.InvariantCultureIgnoreCase);


            while (startTag > -1 && startTag < html.Length - 1)
            {
                var nextLetter = html.Substring(startTag + 1, 2);

                switch (nextLetter.ToLower())
                {
                    case "p ":
                    case "p>":
                    case "h1":
                    case "h2":
                    case "h3":
                    case "h4":
                    case "h5":

                        var endTag = html.IndexOf("<", startTag + 2, StringComparison.InvariantCultureIgnoreCase);
                        var endTagStartTag = html.IndexOf(">", startTag + 1, StringComparison.InvariantCultureIgnoreCase);

                        var content = html.Substring(endTagStartTag + 1, endTag - endTagStartTag - 1);

                        switch (nextLetter.ToLower())
                        {
                            case "p ":
                            case "p>":
                                AddParagraph(content);
                                break;
                            case "h1":
                                AddParagraph(content, "Heading1");
                                break;
                            case "h2":
                                AddParagraph(content, "Heading2");
                                break;
                            case "h3":
                                AddParagraph(content, "Heading3");
                                break;
                            case "h4":
                                AddParagraph(content, "Heading4");
                                break;
                            case "h5":
                                AddParagraph(content, "Heading5");
                                break;
                        }

                        startTag = endTag + 1;
                        break;
                    //default:
                    //    break;
                }



                startTag = html.IndexOf("<", startTag + 1, StringComparison.InvariantCultureIgnoreCase);
            }

        }


        public void CreateFooter3(string footerLeft, string footerMiddle, string footerRight, string styleName)
        {


            var table = _content.Footers.Primary.AddTable();
            table.Borders.Visible = false;
            table.TopPadding = 9;
            //table.Rows.LeftIndent = Unit.FromCentimeter(1);

            var w = Unit.FromCentimeter(Width / 3);

            var col = table.AddColumn();
            col.Width = w;

            col = table.AddColumn();
            col.Format.Alignment = ParagraphAlignment.Center;
            col.Width = w;

            col = table.AddColumn();
            col.Format.Alignment = ParagraphAlignment.Right;
            col.Width = w;

            table.AddRow();

            var cell = table.Rows[0][0];
            var p = cell.AddParagraph(footerLeft);
            p.Style = styleName;

            cell = table.Rows[0][1];
            p = cell.AddParagraph(footerMiddle);
            p.Style = styleName;


            cell = table.Rows[0][2];
            p = cell.AddParagraph(footerRight);
            p.Style = styleName;


        }


        /// <summary>
        /// Add a pagebreak to the content section
        /// </summary>
        public void AddPageBreak()
        {
            _content.AddPageBreak();
        }


        public void Dispose()
        {
            _content = null;
            _toc = null;
            _document = null;
        }
    }
}

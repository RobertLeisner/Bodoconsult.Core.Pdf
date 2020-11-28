using MigraDoc.DocumentObjectModel;

namespace Bodoconsult.Core.Pdf.Stylesets
{

    /// <summary>
    /// Defines a default style set.Use subclassing for creating advanced stylesets
    /// </summary>
    public class DefaultStyleSet : IStyleSet
    {

        public DefaultStyleSet()
        {

            double verticalMargin = 3;

            PageSetup = new PageSetup
            {
                Orientation = Orientation.Portrait,
                PageFormat = PageFormat.A4,
                PageWidth = Unit.FromCentimeter(21.0),
                PageHeight = Unit.FromCentimeter(29.7),
                LeftMargin = Unit.FromCentimeter(2.5),
                TopMargin = Unit.FromCentimeter(1.9),
                RightMargin = Unit.FromCentimeter(2.5),
                BottomMargin = Unit.FromCentimeter(1.9)
            };


            var width = PageSetup.PageWidth - PageSetup.LeftMargin - PageSetup.RightMargin;

            Normal = new Style("Normal", null)
            {
                Font =
                {
                    Name = "Arial",
                    Size = 10,
                    Color = Colors.Black
                },
                ParagraphFormat =
                {
                    SpaceBefore = verticalMargin,
                    SpaceAfter = verticalMargin,
                    PageBreakBefore = false,
                    Alignment = ParagraphAlignment.Left,
                    
                }
            };

            //Normal.ParagraphFormat.Borders.Left.Width = 1;
            //Normal.ParagraphFormat.Borders.Left.Color = Colors.Black;

            NormalTable = new Style("NormalTable", "Normal")
            {
                Font =
                {
                    Name = "Arial Narrow",
                    Size = 8
                },
                ParagraphFormat =
                {
                    SpaceBefore = 0,
                    SpaceAfter = 0,
                    PageBreakBefore = false,
                    Alignment = ParagraphAlignment.Center
                }
            };

            Heading1 = new Style("Heading1", "Normal")
            {
                Font =
                {
                    Name = "Arial Black",
                    Size = 16,
                    Color = Colors.Black
                },
                ParagraphFormat =
                {
                    PageBreakBefore = true,
                    KeepTogether = true,
                    KeepWithNext = true,
                    SpaceAfter = verticalMargin,
                    SpaceBefore = 3 * verticalMargin,
                    Alignment = ParagraphAlignment.Left
                }
            };


            Heading2 = new Style("Heading2", "Normal")
            {
                Font =
                {
                    Name = "Arial Black",
                    Size = 14,
                    Color = Colors.Black
                },
                ParagraphFormat =
                {
                    KeepTogether = true,
                    KeepWithNext = true,
                    SpaceAfter = verticalMargin,
                    SpaceBefore = 2 * verticalMargin,
                    Alignment = ParagraphAlignment.Left
                }
            };

            Heading3 = new Style("Heading3", "Normal")
            {
                Font =
                {
                    Name = "Arial",
                    Size = 12,
                    Color = Colors.Black,
                    Bold = true
                },
                ParagraphFormat =
                {
                    KeepTogether = true,
                    KeepWithNext = true,
                    SpaceAfter = verticalMargin,
                    SpaceBefore = 2 * verticalMargin,
                    Alignment = ParagraphAlignment.Left
                }
            };


            Heading4 = new Style("Heading4", "Normal")
            {
                Font =
                {
                    Name = "Arial",
                    Size = 12,
                    Color = Colors.Black,
                    Bold = true
                },
                ParagraphFormat =
                {
                    KeepTogether = true,
                    KeepWithNext = true,
                    SpaceAfter = verticalMargin,
                    SpaceBefore = verticalMargin,
                    Alignment = ParagraphAlignment.Left
                }
            };


            Title = new Style("Title", "Normal")
            {
                Font =
                {
                    Name = "Arial Black",
                    Size = 20,
                    Color = Colors.Black
                },
                ParagraphFormat =
                {
                    SpaceAfter = 4 * verticalMargin,
                    SpaceBefore = 5 * verticalMargin,
                    Alignment = ParagraphAlignment.Center
                }
            };


            NoHeading1 = new Style("NoHeading1", "Normal")
            {
                Font =
                {
                    Name = "Arial Black",
                    Size = 10,
                    Color = Colors.Black
                },
                ParagraphFormat =
                {
                    SpaceAfter = verticalMargin,
                    SpaceBefore = 2 * verticalMargin,
                    Alignment = ParagraphAlignment.Center
                }
            };

            //style.ParagraphFormat.
            //style.ParagraphFormat.Shading.Color = Colors.Orange;


            ChartTitle = new Style("ChartTitle", "Normal")
            {
                Font =
                {
                    Name = "Arial Narrow",
                    Size = 10,
                    Bold = true,
                    Color = Colors.Red
                },
                ParagraphFormat =
                {
                    SpaceBefore = 1.5 * verticalMargin,
                    Alignment = ParagraphAlignment.Center,
                    SpaceAfter = 1.5 * verticalMargin
                }
            };

            ChartYLabel = new Style("ChartYLabel", "Normal")
            {
                Font =
                {
                    Name = "Arial Narrow",
                    Size = 8,
                    Color = Colors.Black
                }
            };

            Toc1 = new Style("TOC1", "Normal")
            {
                Font =
                {
                    Name = "Arial Black",
                    Size = 10,
                    Color = Colors.Black
                },
                ParagraphFormat =
                {
                    SpaceBefore = 2 * verticalMargin,
                    SpaceAfter = verticalMargin
                }
            };

            Toc1.ParagraphFormat.Borders.Bottom.Width = 0;
            Toc1.ParagraphFormat.Borders.Top.Width = 0;
            Toc1.ParagraphFormat.Borders.Right.Width = 0;
            Toc1.ParagraphFormat.Borders.Left.Width = 0;
            Toc1.ParagraphFormat.TabStops.AddTabStop(width, TabAlignment.Right);

            Toc2 = new Style("TOC2", "Normal")
            {
                Font =
                {
                    Name = "Arial",
                    Size = 10,
                    Color = Colors.Black
                },
                ParagraphFormat =
                {
                    SpaceBefore = 2 * verticalMargin,
                    SpaceAfter = verticalMargin
                }
            };

            Toc2.ParagraphFormat.Borders.Bottom.Width = 0;
            Toc2.ParagraphFormat.Borders.Top.Width = 0;
            Toc2.ParagraphFormat.Borders.Right.Width = 0;
            Toc2.ParagraphFormat.Borders.Left.Width = 0;
            Toc2.ParagraphFormat.LeftIndent = Unit.FromCentimeter(1);
            Toc2.ParagraphFormat.TabStops.AddTabStop(width, TabAlignment.Right);

            Toc3 = new Style("TOC3", "Normal")
            {
                Font =
                {
                    Name = "Arial",
                    Size = 10,
                    Color = Colors.Black
                },
                ParagraphFormat =
                {
                    SpaceBefore = 2 * verticalMargin,
                    SpaceAfter = verticalMargin
                }
            };

            Toc3.ParagraphFormat.Borders.Bottom.Width = 0;
            Toc3.ParagraphFormat.Borders.Top.Width = 0;
            Toc3.ParagraphFormat.Borders.Right.Width = 0;
            Toc3.ParagraphFormat.Borders.Left.Width = 0;
            Toc3.ParagraphFormat.LeftIndent = Unit.FromCentimeter(2);
            Toc3.ParagraphFormat.TabStops.AddTabStop(width, TabAlignment.Right);

            Toc4 = new Style("TOC4", "Normal")
            {
                Font =
                {
                    Name = "Arial",
                    Size = 10,
                    Color = Colors.Black
                },
                ParagraphFormat =
                {
                    SpaceBefore = 2 * verticalMargin,
                    SpaceAfter = verticalMargin
                }
            };

            Toc4.ParagraphFormat.Borders.Bottom.Width = 0;
            Toc4.ParagraphFormat.Borders.Top.Width = 0;
            Toc4.ParagraphFormat.Borders.Right.Width = 0;
            Toc4.ParagraphFormat.Borders.Left.Width = 0;
            Toc4.ParagraphFormat.LeftIndent = Unit.FromCentimeter(3);
            Toc4.ParagraphFormat.TabStops.AddTabStop(width, TabAlignment.Right);


            TocHeading1 = new Style("TocHeading1", "Normal")
            {
                Font =
                {
                    Name = "Arial Black",
                    Size = 12,
                    Color = Colors.Black
                },
                ParagraphFormat = { SpaceAfter = verticalMargin }
            };
            TocHeading1.ParagraphFormat.SpaceAfter = verticalMargin;
            TocHeading1.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Kopfzeile
            Header = new Style("Header", "Normal")
            {
                Font =
                {
                    Name = "Arial Narrow",
                    Size = 8,
                    Color = Colors.Black
                },
                ParagraphFormat =
                {
                    Alignment = ParagraphAlignment.Right,
                    SpaceAfter = 4 * verticalMargin,
                    SpaceBefore = 0
                }
            };
            Header.ParagraphFormat.Borders.Bottom.Width = 1;
            Header.ParagraphFormat.Borders.Bottom.Color = Colors.Black;


            Details = new Style("Details", "Normal")
            {
                Font =
                {
                    Name = "Arial Narrow",
                    Size = 10,
                    Color = Colors.Black
                },
                ParagraphFormat =
                {
                    SpaceAfter = 3 * verticalMargin,
                    Alignment = ParagraphAlignment.Center
                }
            };



            Table = new Style("Table", "Normal")
            {
                Font =
                {
                    Name = "Arial Narrow",
                    Size = 9,
                    Color = Colors.Black
                },
                ParagraphFormat =
                {
                    SpaceAfter = 3 * verticalMargin,
                    Alignment = ParagraphAlignment.Center
                }
            };



            ChartCell = new Style("ChartCell", "Normal")
            {
                ParagraphFormat = { Alignment = ParagraphAlignment.Center }
            };

            Footer = new Style("Footer", "Normal")
            {
                Font =
                {
                    Name = "Arial Narrow",
                    Size = 8,
                    Color = Colors.Black
                },
                ParagraphFormat =
                {
                    Alignment = ParagraphAlignment.Left,
                    SpaceAfter = 0
                }
            };

            Footer.ParagraphFormat.Borders.Top.Visible = true;
            Footer.ParagraphFormat.Borders.Top.Width = 0.5;
            Footer.ParagraphFormat.Borders.Top.Color = Colors.Black;
            Footer.ParagraphFormat.AddTabStop(width, TabAlignment.Right);

            Code = new Style("Code", "Normal")
            {
                Font =
                {
                    Name = "Courier New",
                    Size = 10,
                    Color = Colors.Black,
                    Italic = true
                },
                ParagraphFormat =
                {
                    Alignment = ParagraphAlignment.Left,
                    SpaceBefore = 2 * verticalMargin,
                    SpaceAfter = 2 * verticalMargin
                }
            };
            Code.ParagraphFormat.TabStops.ClearAll();
            Code.ParagraphFormat.TabStops.AddTabStop("1cm");
            Code.ParagraphFormat.TabStops.AddTabStop("2cm");
            Code.ParagraphFormat.TabStops.AddTabStop("3cm");
            Code.ParagraphFormat.TabStops.AddTabStop("4cm");
            Code.ParagraphFormat.TabStops.AddTabStop("5cm");
            Code.ParagraphFormat.TabStops.AddTabStop("6cm");
            Code.ParagraphFormat.TabStops.AddTabStop("7cm");
            Code.ParagraphFormat.TabStops.AddTabStop("8cm");
            Code.ParagraphFormat.TabStops.AddTabStop("9cm");
            Code.ParagraphFormat.TabStops.AddTabStop("10cm");
            Code.ParagraphFormat.TabStops.AddTabStop("11cm");
            Code.ParagraphFormat.TabStops.AddTabStop("12cm");
            Code.ParagraphFormat.TabStops.AddTabStop("13cm");
            Code.ParagraphFormat.TabStops.AddTabStop("14cm");
            Code.ParagraphFormat.TabStops.AddTabStop("26.7cm", TabAlignment.Right);


            Bullet1 = new Style("Bullet1", "Normal")
            {
                Font =
                {
                    Name = "Arial",
                    Size = 10,
                    Color = Colors.Black
                },
                ParagraphFormat =
                {
                    Alignment = ParagraphAlignment.Left,
                    SpaceBefore = verticalMargin,
                    SpaceAfter = verticalMargin
                }
            };
            Bullet1.ParagraphFormat.AddTabStop("26.7cm", TabAlignment.Right);
            Bullet1.ParagraphFormat.LeftIndent = Unit.FromCentimeter(1);
            var li = new ListInfo { ListType = ListType.BulletList1 };
            Bullet1.ParagraphFormat.ListInfo = li;

        }

        /// <summary>
        /// Normal paragraphs (default style)
        /// </summary>
        public PageSetup PageSetup { get; set; }


        /// <summary>
        /// Normal paragraphs (default style)
        /// </summary>
        public Style Normal { get; }

        /// <summary>
        /// Spezielles Format für Tabellenbasis (nicht ändern!)
        /// </summary>
        public Style NormalTable { get; }


        /// <summary>
        /// Table format
        /// </summary>
        public Style Table { get; }


        /// <summary>
        /// Title
        /// </summary>
        public Style Title { get; }

        /// <summary>
        /// Heading level 1
        /// </summary>
        public Style Heading1 { get; }

        /// <summary>
        /// Heading level 2
        /// </summary>
        public Style Heading2 { get; }

        /// <summary>
        /// Heading level 3
        /// </summary>
        public Style Heading3 { get; }


        /// <summary>
        /// Heading level 4
        /// </summary>
        public Style Heading4 { get; }


        public Style NoHeading1 { get; }


        public Style ChartTitle { get; }


        public Style ChartYLabel { get; }


        public Style Toc1 { get; }


        public Style Toc2 { get; }

        public Style Toc3 { get; }


        public Style Toc4 { get; }



        public Style TocHeading1 { get; }


        public Style Header { get; }



        public Style Details { get; }



        public Style ChartCell { get; }



        public Style Footer { get; }



        public Style Code { get; }



        public Style Bullet1 { get; }


    }
}

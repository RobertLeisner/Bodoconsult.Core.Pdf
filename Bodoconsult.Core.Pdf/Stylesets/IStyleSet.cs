// Copyright (c) Bodoconsult EDV-Dienstleistungen GmbH. All rights reserved.

using MigraDoc.DocumentObjectModel;

namespace Bodoconsult.Core.Pdf.Stylesets
{

    /// <summary>
    /// Interface for stylesets used to format PDF files created by PdfSharp and MigraDoc
    /// </summary>
    public interface IStyleSet
    {
        /// <summary>
        /// Normal paragraphs (default style)
        /// </summary>
        Style Normal { get; }

        /// <summary>
        /// Spezielles Format für Tabellenbasis (nicht ändern!)
        /// </summary>
        Style NormalTable { get; }

        /// <summary>
        /// Title
        /// </summary>
        Style Title { get; }

        /// <summary>
        /// Heading level 1
        /// </summary>
        Style Heading1 { get; }

        /// <summary>
        /// Heading level 2
        /// </summary>
        Style Heading2 { get; }

        /// <summary>
        /// Heading level 3
        /// </summary>
        Style Heading3 { get; }

        /// <summary>
        /// Heading level 4
        /// </summary>
        Style Heading4 { get; }

        Style NoHeading1 { get; }
        Style ChartTitle { get; }
        Style ChartYLabel { get; }

        Style Table { get; }

        Style Toc1 { get; }

        Style Toc2 { get; }

        Style Toc3 { get; }

        Style Toc4 { get; }


        Style TocHeading1 { get; }

        Style Header { get; }
        Style Details { get; }
        Style ChartCell { get; }
        Style Footer { get; }
        Style Code { get; }
        Style Bullet1 { get; }

        PageSetup PageSetup { get; set; }
    }
}
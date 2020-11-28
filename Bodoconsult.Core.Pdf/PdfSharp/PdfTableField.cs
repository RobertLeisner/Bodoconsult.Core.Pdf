namespace Bodoconsult.Core.Pdf.PdfSharp
{
    /// <summary>
    /// Collects data for presentation of a object field in a pdf table
    /// </summary>
    public struct PdfTableField
    {
        public string Header;

        public string Name;

        public double Width;

        public string Format;

        public PdfTextAlignment TextAlign;
    }
}
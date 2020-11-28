using System;
using System.Data;
using System.Linq;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes.Charts;

namespace Bodoconsult.Core.Pdf.PdfSharp
{
    public class PdfChart
    {
        readonly Chart _chart = new Chart();


        public PdfChart()
        {
            HasDataLabel = false;
            Width = 13;
            Height = 8.5;
            Left = 0;
        }


        public double Left; // { get; set; }


        public double Width; // { get; set; }

        public double Height; // { get; set; }

        public ChartType TypeOfChart; // { get; set; }

        public bool HasDataLabel; // { get; set; }

        public string XAxisLabel; // { get; set; }
        public string YAxisLabel; // { get; set; }

        public int LegendLength; // { get; set; }

        private int _xDimension;

        public string Title; // { get; set; }

        public Chart Draw()
        {
            _chart.Width = Unit.FromCentimeter(Width);
            _chart.Height = Unit.FromCentimeter(Height);

            _chart.XAxis.MajorTickMark = TickMarkType.Outside;
            if (!string.IsNullOrEmpty(XAxisLabel)) _chart.XAxis.Title.Caption = XAxisLabel;

            _chart.YAxis.MajorTickMark = TickMarkType.Outside;
            _chart.YAxis.HasMajorGridlines = true;

            _chart.PlotArea.LineFormat.Color = Colors.SteelBlue;
            _chart.PlotArea.LineFormat.Width = 1;
            _chart.PlotArea.BottomPadding = Unit.FromCentimeter(0.2);
            _chart.PlotArea.LeftPadding = Unit.FromCentimeter(0.2);
            
            _chart.Format.Font.Size = 8;
            _chart.HeaderArea.Style = "ChartTitle";
            _chart.HeaderArea.AddParagraph(Title);
            _chart.FillFormat.Color = Colors.LightSteelBlue;
            _chart.PlotArea.FillFormat.Color = Colors.White;
            _chart.LineFormat.Color = Colors.SteelBlue;



            if (_chart.SeriesCollection.Count > 1)
            {

                _chart.SeriesCollection[0].MarkerBackgroundColor = Colors.YellowGreen;
                _chart.SeriesCollection[0].MarkerSize = 2;
                _chart.SeriesCollection[0].MarkerStyle = MarkerStyle.Plus;
                _chart.SeriesCollection[0].FillFormat.Color = Colors.YellowGreen;
                _chart.SeriesCollection[0].LineFormat.Color = Colors.YellowGreen;
                _chart.SeriesCollection[0].LineFormat.Width = 2;

                _chart.SeriesCollection[1].FillFormat.Color = Colors.Orange;
                _chart.SeriesCollection[1].MarkerBackgroundColor = Colors.Orange;
                _chart.SeriesCollection[1].MarkerSize = 2;
                _chart.SeriesCollection[1].MarkerStyle = MarkerStyle.Plus;
                _chart.SeriesCollection[1].LineFormat.Color = Colors.Orange;
                _chart.SeriesCollection[1].LineFormat.Width = 2;
            }
            else
            {
                _chart.SeriesCollection[0].FillFormat.Color = Colors.Orange;
                _chart.SeriesCollection[0].MarkerBackgroundColor = Colors.Orange;
                _chart.SeriesCollection[0].MarkerSize = 2;
                _chart.SeriesCollection[0].MarkerStyle = MarkerStyle.Plus;
                _chart.SeriesCollection[0].LineFormat.Color = Colors.Orange;
                _chart.SeriesCollection[0].LineFormat.Width = 2;       
            }

            if (!string.IsNullOrEmpty(YAxisLabel))
            {
                _chart.LeftArea.Style = "ChartYLabel";
                _chart.LeftArea.AddParagraph(YAxisLabel);
                _chart.LeftArea.LeftPadding = Unit.FromCentimeter(0.2);
            }


            _chart.RightArea.AddLegend();

            return _chart;

        }


        public void AddSeries(double[] data, string name)
        {
            var series = _chart.SeriesCollection.AddSeries();
            series.ChartType = TypeOfChart;
            series.Add(data);
            series.Name = name;
            series.MarkerSize = Unit.FromCentimeter(0.1);
            series.HasDataLabel = HasDataLabel;

        }


        public void AddSeries(DataTable dataTable, int columnId, string name)
        {

            var x = new double[_xDimension];

            for (var i = 0; i < _xDimension; i++)
            {
                var v = dataTable.Rows[i][columnId].ToString();

                if (string.IsNullOrEmpty(v)) continue;
                x[i] = Convert.ToDouble(v);
            }



            //var query = from mycolumn in dataTable.AsEnumerable()
            //            where mycolumn.Field<decimal>(columnId) != null 
            //            select mycolumn;

            //var x = query.Select(i => i.Field<Double>(columnId)).ToArray(); 

            AddSeries(x, name);

        }


        public void XSeries(params string[] s)
        {
            _xDimension = s.Length;
            var xseries = _chart.XValues.AddXSeries();
            xseries.Add(s);
        }

        public void XSeries(DataTable dataTable, int columnId)
        {
            if (LegendLength > 15) LegendLength = 15;
            
            var query = from mycolumn in dataTable.AsEnumerable()
                        where mycolumn.Field<string>(columnId) != "Gesamt"
                        select mycolumn;


            var x = query.Select(i => i.Field<string>(columnId).Substring(0, LegendLength)).ToArray();
            _xDimension = x.Length;

            var xseries = _chart.XValues.AddXSeries();
            xseries.Add(x);


        }

    }
}

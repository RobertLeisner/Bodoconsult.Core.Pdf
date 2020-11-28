using System;
using System.Data;
using System.Diagnostics;

namespace Bodoconsult.Core.Pdf.Test.Helpers
{
    internal class TestHelper
    {


        public const string Masstext1 =
            "Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet. Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet.";


        /// <summary>
        /// Opens a file in a separate process. Executable not required
        /// </summary>
        /// <param name="path"></param>
        public static void OpenFile(string path)
        {
            var p = new Process
            {
                StartInfo = new ProcessStartInfo(path)
                {
                    UseShellExecute = true
                }
            };
            p.Start();
        }


        public static DataTable GetDefinitionList()
        {
            var dt = new DataTable();
            dt.Columns.Add("Key");
            dt.Columns.Add("Value");


            var item = dt.NewRow();
            item["Key"] = "Content";
            item["Value"] = Masstext1;
            dt.Rows.Add(item);

            item = dt.NewRow();
            item["Key"] = "Info";
            item["Value"] = Masstext1;
            dt.Rows.Add(item);

            item = dt.NewRow();
            item["Key"] = "Info";
            item["Value"] = "500";
            dt.Rows.Add(item);

            item = dt.NewRow();
            item["Key"] = "Info";
            item["Value"] = Masstext1;
            dt.Rows.Add(item);

            item = dt.NewRow();
            item["Key"] = "Info";
            item["Value"] = Masstext1;
            dt.Rows.Add(item);

            item = dt.NewRow();
            item["Key"] = "Info";
            item["Value"] = Masstext1;
            dt.Rows.Add(item);

            item = dt.NewRow();
            item["Key"] = "Info";
            item["Value"] = Masstext1;
            dt.Rows.Add(item);

            return dt;
        }

        public static DataTable GetDataTable()
        {
            var dt = new DataTable();
            dt.Columns.Add("Key");
            dt.Columns.Add("Value1").DataType = typeof(double);
            dt.Columns.Add("Value2").DataType = typeof(DateTime);
            dt.Columns.Add("Value3").DataType = typeof(bool);

            var item = dt.NewRow();
            item["Key"] = "Content";
            item["Value1"] = 5.97;
            item["Value2"] = DateTime.Now;
            item["Value3"] = true;
            dt.Rows.Add(item);

            item = dt.NewRow();
            item["Key"] = "Website";
            item["Value1"] = 65.97;
            item["Value2"] = DateTime.Now.AddDays(-25);
            item["Value3"] = false;
            dt.Rows.Add(item);

            item = dt.NewRow();
            item["Key"] = "Marketing";
            item["Value1"] = 45.97;
            item["Value2"] = DateTime.Now;
            item["Value3"] = true;
            dt.Rows.Add(item);

            return dt;
        }


    }
}

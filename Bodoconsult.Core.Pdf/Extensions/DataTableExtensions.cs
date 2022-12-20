// Copyright (c) Bodoconsult EDV-Dienstleistungen GmbH. All rights reserved.


using System.Collections.Generic;
using System.Data;

namespace Bodoconsult.Core.Pdf.Extensions
{
    public static class DataTableExtensions
    {

        public static IEnumerable<DataRow> EnumerateRows(this DataTable table)
        {
            foreach (DataRow row in table.Rows)
            {
                yield return row;
            }
        }

        public static IEnumerable<DataColumn> EnumerateColumns(this DataTable table)
        {

            

            foreach (DataColumn column in table.Columns)
            {
                yield return column;
            }
        }
    }
}

/* Author: John Shield
 * Date: 2014 October
 */

using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using ICSharpCode.SharpZipLib.Zip;
using System.Diagnostics;
using XlsxReadWrite.Properties;
using System.Text.RegularExpressions;
using System;

namespace XlsxReadWrite
{
    internal static class XlsxRW
    {
        private static double [] columnSizes = new double[80];
        private static int [] filter_A_col;
        private static bool [] filter_A_matchtype; 
        private static string [] filter_A_string;
        private static int[] filter_R_col;
        private static bool[] filter_R_matchtype;
        private static string[] filter_R_string;

        public static void DeleteDirectoryContents(string directory)
        {
            var info = new DirectoryInfo(directory);
            try
            {
                foreach (var file in info.GetFiles())
                    file.Delete();

                foreach (var dir in info.GetDirectories())
                    dir.Delete(true);
            }
            catch (DirectoryNotFoundException)
            {
                Directory.CreateDirectory(directory);
            }

        }

        public static bool UnzipFile(string zipFileName, string targetDirectory)
        {
            try
            {
                new FastZip().ExtractZip(zipFileName, targetDirectory, null);
            }
            catch (IOException)
            {
                System.Windows.Forms.MessageBox.Show("File is in use by another program: " + zipFileName,
                    "Read Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation,
                    System.Windows.Forms.MessageBoxDefaultButton.Button1);
                return false;
            }
            return true;
        }

        public static bool ZipDirectory(string sourceDirectory, string zipFileName)
        {
            try
            {
                new FastZip().CreateZip(zipFileName, sourceDirectory, true, null);
            }
            catch (IOException)
            {
                System.Windows.Forms.MessageBox.Show("File is in use by another program: " + zipFileName,
                    "Read Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation,
                    System.Windows.Forms.MessageBoxDefaultButton.Button1);
                return false;
            }
            return true;
        }

        public static IList<string> ReadStringTable(Stream input)
        {
            var stringTable = new List<string>();

            using (var reader = XmlReader.Create(input))
                for (reader.MoveToContent(); reader.Read(); )
                    if (reader.NodeType == XmlNodeType.Element && reader.Name == "t")
                        stringTable.Add(reader.ReadElementString());

            return stringTable;
        }

        public static int[] ToIntArray(this string value, char separator)
        {
            return Array.ConvertAll(value.Split(separator), s => int.Parse(s));
        }

        public static int getDate(Object obj) {
            if (obj == DBNull.Value)
                return 0;
            try {
                DateTime outputDate;
                if (DateTime.TryParse(Convert.ToString(obj), out outputDate))
                {
                    return Convert.ToInt32(outputDate.ToOADate());
                }
                else 
                {
                    //Not  a valid date
                    return 0;
                } 

                //return Convert.ToInt32(DateTime.Parse(Convert.ToString(obj)).ToOADate());
            }
            catch (OverflowException) {
                return 0;
            }   
            catch (FormatException) {
                return 0;
            }

        }

        public static bool Contains(string source, string toCheck, StringComparison comp)
        {
            return source.IndexOf(toCheck, comp) >= 0;
        }

        public static bool filterRow(DataRow row)
        {
          // accept override filters
            for (int ii = 0; filter_A_col != null && ii < filter_A_col.Length; ii++)
            {
              if (filter_A_matchtype[ii]) {
                  if (row[filter_A_col[ii]-1].Equals(filter_A_string[ii]))
                      return false;
              } else {
                  if (Contains(row[filter_A_col[ii]-1].ToString(), filter_A_string[ii], StringComparison.Ordinal))
                      return false;
              }
          }

          // rejection filters
          for (int ii = 0; filter_R_col != null && ii < filter_R_col.Length; ii++)
          {
              if (filter_R_matchtype[ii])
              {
                  if (row[filter_R_col[ii]-1].Equals(filter_R_string[ii]))
                      return true;
              }
              else
              {
                  if (Contains(row[filter_R_col[ii]-1].ToString(), filter_R_string[ii], StringComparison.Ordinal))
                      return true;
              }
          }

          // if there's no rejection filters, but there are acceptance filters, then reject the rest
          if (filter_A_col != null && filter_R_col == null)
              return true;

          return false; 
        }

        public static int insertRow(DataTable data, DataRow row, int rowIndex, bool mainsheet)
        {
            int matchIndex = -1;
            if (row != null)
            {
               if (mainsheet == false) // if not main sheet
                   if (filterRow(row)) // check whether it's filtered
                       return 0;       // return if row is filtered

               int newIndex;
               int rowDate = getDate((Object)row[1]);
               //Debug.WriteLine("ROWDATE:" + rowDate);
               // if there's a row already at the location
               for (newIndex = 0; data.Rows.Count > newIndex; newIndex++)
               {
                   //Debug.WriteLine("DATECOMPAR:" + data.Rows[newIndex][1] + ":" + getDate(data.Rows[newIndex][1]));
                   // TODO add these as options
                   // This does same line matching
                   if (data.Rows[newIndex][1].Equals(row[1]))
                   {
                       if (data.Rows[newIndex][0].Equals(row[0]) && // same ID
                           data.Rows[newIndex][2].Equals(row[2]) && // same start time
                           data.Rows[newIndex][3].Equals(row[3])) // same end time
                       {
                           if (Contains(row[4].ToString(), data.Rows[newIndex][4].ToString(), StringComparison.Ordinal))
                           {
                               matchIndex = newIndex;
                               break;
                           }
                       }
                   }
                   // if not found
                   else if (getDate(data.Rows[newIndex][1]) > rowDate)
                   {
                       //Debug.WriteLine("Compare:" + getDate(data.Rows[newIndex][1]) + "> " + rowDate);
                       rowIndex = newIndex;
                       break;
                   }
               }
               // if the end of the database was reached without a match or a larger date
               if (data.Rows.Count == newIndex)
               {
                   rowIndex = newIndex;
               }

                if (matchIndex == -1)
                {
                    if (row[1] == DBNull.Value) // delete rows without a date
                    {
                        //data.Rows.Add(row);
                    }
                    else
                    {
                        //if (data.Rows.Count > rowIndex)
                        //{Debug.WriteLine("ROWDATE:" + rowDate);}
                        data.Rows.InsertAt(row, rowIndex);
                    }
                } else {
                    // only overwrite when it's the priority sheet (local copy)
                    if (mainsheet)
                    {
                        data.Rows.RemoveAt(matchIndex);
                        data.Rows.InsertAt(row, matchIndex);
                    }
                    rowIndex = matchIndex;
                }
                rowIndex++;
            }
            return rowIndex;
        }

        public static void generate_filters()
        {
            if (!string.IsNullOrEmpty(Settings.Default.IncludeMatchArray_col) && !string.IsNullOrEmpty(Settings.Default.IncludeMatchArray_exact)
                  && !string.IsNullOrEmpty(Settings.Default.IncludeMatchArray_str))
            {
                filter_A_string = Settings.Default.IncludeMatchArray_str.Split(',');
                filter_A_col = Array.ConvertAll(Settings.Default.IncludeMatchArray_col.Split(','), s => int.Parse(s));
                filter_A_matchtype = Array.ConvertAll(Settings.Default.IncludeMatchArray_exact.Split(','), s => bool.Parse(s));
            }
            if (!string.IsNullOrEmpty(Settings.Default.RemoveMatchArray_col) && !string.IsNullOrEmpty(Settings.Default.RemoveMatchArray_exact)
                && !string.IsNullOrEmpty(Settings.Default.RemoveMatchArray_str))
            {
                filter_R_string = Settings.Default.RemoveMatchArray_str.Split(',');
                filter_R_col = Array.ConvertAll(Settings.Default.RemoveMatchArray_col.Split(','), s => int.Parse(s));
                filter_R_matchtype = Array.ConvertAll(Settings.Default.RemoveMatchArray_exact.Split(','), s => bool.Parse(s));
            }
        }

        // This is the State worksheet
        public static void ReadWorksheet(Stream input, IList<string> stringTable, DataTable data)
        {
            // get the column settings
            //Debug.WriteLine("Info:" + Settings.Default.Input1Columns);
            int[] column_reorder = ToIntArray(Settings.Default.Input1Columns, ',');
            generate_filters(); // these filters are only used here
            ParseWorksheet(input, stringTable, data, column_reorder, false);
        }

        // This is the AMC worksheet
        public static void ReadWorksheet2(Stream input, IList<string> stringTable, DataTable data)
        {
            // get the column settings
            //Debug.WriteLine("Info:" + Settings.Default.Input2Columns);
            int[] column_reorder = ToIntArray(Settings.Default.Input2Columns, ',');
            ParseWorksheet(input, stringTable, data, column_reorder, true);
        }

        public static void ParseWorksheet(Stream input, IList<string> stringTable, DataTable data, int[] column_reorder, bool main_sheet)
        {
            using (var reader = XmlReader.Create(input))
            {
                DataRow row = null;
                int columnIndex = 0;
                int readColumnIndex = 0;
                int rowIndex = 0;
                int test = 0;
                string type;
                string style;
                int value;
                string hidden = ""; // skip the hidden rows

                for (reader.MoveToContent(); reader.Read(); )
                    if (reader.NodeType == XmlNodeType.Element)
                        switch (reader.Name)
                        {
                            case "cols":
                                // this saves the column widths for later
                                if (main_sheet)
                                    GetCols(reader, data);
                                break;
                            case "row":
                                // insert the row if required
                                rowIndex = insertRow(data, row, rowIndex, main_sheet);
                                row = null;
                                // check if the row is hidden (remove hidden rows for input sheets)
                                if (Settings.Default.FilterHidden)
                                {
                                    hidden = reader.GetAttribute("hidden");
                                    if (hidden == "1")
                                        break;
                                }
                                row = data.NewRow();

                                columnIndex = 0;
                                test = 0;
                                break;

                            case "c":
                                if (hidden == "1")
                                    break;


                                readColumnIndex = getCol(reader.GetAttribute("r"));
                                //Debug.WriteLine("test:" + test + " col:" + readColumnIndex);

                                if (readColumnIndex >= column_reorder.Length) {
                                    // out of bounds for listed remapping
                                    break;
                                } else if (column_reorder[readColumnIndex] == 0) {
                                    // ignore this column
                                    test++;
                                    break;
                                } else {
                                    // set the column remapping
                                    columnIndex = column_reorder[readColumnIndex] - 1;
                                    test++;
                                }

                                // sheet is larger than list boundaries
                                if (data.Columns.Count < columnIndex)
                                    break;

                                // when type is missing, but its not an empty square, it's probably date format
                                //Debug.WriteLine("type:" + type + " style:" + style);
                                type = reader.GetAttribute("t");
                                style = reader.GetAttribute("s");

                                // Debug.WriteLine("reader.Read:" + returnval);
                                if (reader.IsStartElement() == false || reader.IsEmptyElement)
                                {// don't do more if cell is empty
                                    break;
                                }

                                // load the value
                                bool returnval = reader.Read();
                                string cellstring = reader.ReadElementString();

                                try
                                {
                                    //strings and integers are stored as ints
                                    value = int.Parse(cellstring, CultureInfo.InvariantCulture);
                                }
                                catch (FormatException)
                                {
                                    // if it's not an INT value, assume it might be a double (hence time value)
                                    double time = double.Parse(cellstring, CultureInfo.InvariantCulture);

                                    row[columnIndex] = TimeSpan.FromDays(time).ToString(@"hh\:mm");
                                    break;
                                }

                                //Debug.WriteLine("value:" + value + " " + type);

                                //Turn numbers in string format into numbers
                                if (columnIndex == 0 && type == "s")
                                {
                                    int result;
                                    if (System.Int32.TryParse(stringTable[value], out result))
                                    {
                                        row[columnIndex] = result;
                                        break;
                                    }
                                }

 
                                //Special handling for column 2 (index 1), Time Column
                                if (columnIndex == 1)
                                {
                                    
                                    if (type == "s")
                                    {
                                        try
                                        {
                                            //Debug.WriteLine("value:" + DateTime.Parse(stringTable[value]));
                                            //row[columnIndex] = DateTime.Parse(stringTable[value]).ToOADate();
                                            row[columnIndex] = DateTime.Parse(stringTable[value]).ToString("ddd dd MMM yy", CultureInfo.InvariantCulture);
                                            //Debug.WriteLine("value:" + row[columnIndex]);
                                            break;
                                        }
                                        catch (FormatException)
                                        { }
                                    }
                                    else {
                                        row[columnIndex] = DateTime.FromOADate(value).ToString("ddd dd MMM yy", CultureInfo.InvariantCulture);
                                        //Debug.WriteLine("value:" + row[columnIndex]);
                                        break;
                                    }
                                    
                                }

                                if (type == "s")
                                {
                                    // Appears to be a problem with "row" missing column information
                                    // the original datatable (data) doesn't seem to have any column information
                                    //  http://stackoverflow.com/questions/3795340/run-time-error-cannot-find-column-0
                                    //Debug.WriteLine("READ:" + stringTable[value]);
                                    //if (value > 0)
                                    //{
                                    row[columnIndex] = stringTable[value];
                                    //}
                                    
                                }
                                else
                                    row[columnIndex] = value;

                                break;
                        }

                // after parsing is done insert the last row
                insertRow(data, row, rowIndex, main_sheet);
                row = null;
            }
        }

        public static void GetCols(XmlReader reader, DataTable data)
        {
            for (int ii=0; ii < columnSizes.Length; ii++)
            { columnSizes[ii] = 0; }

            for (reader.MoveToContent(); reader.Read(); )
                if (reader.NodeType == XmlNodeType.Element)
                    switch (reader.Name)
                    {
                        case "col":
                            int columnNum = Convert.ToInt32(reader.GetAttribute("min"));
                            double columnsize = Convert.ToDouble(reader.GetAttribute("width"));
                            if (columnNum < columnSizes.Length)
                                columnSizes[columnNum] = columnsize;
                            break;
                        case "sheetData":
                            return;
                    }
        }

        public static int getCol(string cell)
        {

            if (string.IsNullOrEmpty(cell))
                return 0;
            cell = Regex.Replace(cell, @"[\d-]", string.Empty);

            int sum = 0;

            for (int i = 0; i < cell.Length; i++)
            {
                sum *= 26;
                sum += (cell[i] - 'A'+1);
            }

            return sum-1;
        }

        public static IList<string> CreateStringTables(DataTable data, out IDictionary<string, int> lookupTable)
        {
            var stringTable = new List<string>();
            lookupTable = new Dictionary<string, int>();

            foreach (DataRow row in data.Rows)
                foreach (DataColumn column in data.Columns)
                    if (column.DataType == typeof(string) && row[column] != System.DBNull.Value)
                    {
                        var value = (string)row[column];

                        if (!lookupTable.ContainsKey(value))
                        {
                            lookupTable.Add(value, stringTable.Count);
                            stringTable.Add(value);
                        }
                    }

            return stringTable;
        }

        public static void WriteStringTable(Stream output, IList<string> stringTable)
        {
            using (var writer = XmlWriter.Create(output))
            {
                writer.WriteStartDocument(true);

                writer.WriteStartElement("sst", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                writer.WriteAttributeString("count", stringTable.Count.ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("uniqueCount", stringTable.Count.ToString(CultureInfo.InvariantCulture));

                foreach (var str in stringTable)
                {
                    writer.WriteStartElement("si");
                    writer.WriteElementString("t", str);
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
            }
        }

        public static string RowColumnToPosition(int row, int column)
        {
            return ColumnIndexToName(column) + RowIndexToName(row);
        }

        public static string ColumnIndexToName(int columnIndex)
        {
            var second = (char)(((int)'A') + columnIndex % 26);

            columnIndex /= 26;

            if (columnIndex == 0)
                return second.ToString();
            else
                return ((char)(((int)'A') - 1 + columnIndex)).ToString() + second.ToString();
        }

        public static string RowIndexToName(int rowIndex)
        {
            return (rowIndex + 1).ToString(CultureInfo.InvariantCulture);
        }

        public static void WriteWorksheet(Stream output, DataTable data, IDictionary<string, int> lookupTable)
        {
            using (XmlTextWriter writer = new XmlTextWriter(output, Encoding.UTF8))
            {
                writer.WriteStartDocument(true);

                writer.WriteStartElement("worksheet");
                writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                writer.WriteAttributeString("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                writer.WriteStartElement("dimension");
                var lastCell = RowColumnToPosition(data.Rows.Count - 1, data.Columns.Count - 1);
                writer.WriteAttributeString("ref", "A1:" + lastCell);
                writer.WriteEndElement();

                writer.WriteStartElement("sheetViews");
                writer.WriteStartElement("sheetView");
                writer.WriteAttributeString("tabSelected", "1");
                writer.WriteAttributeString("workbookViewId", "0");
                writer.WriteEndElement();
                writer.WriteEndElement();

                writer.WriteStartElement("sheetFormatPr");
                writer.WriteAttributeString("defaultRowHeight", "15");
                writer.WriteEndElement();

                // check for no column information
                bool columns = false;
                for (int ii = 0; ii < columnSizes.Length; ii++)
                    if (columnSizes[ii] > 0)
                        columns = true;

                if (columns)
                {
                    writer.WriteStartElement("cols");
                    for (int ii = 0; ii < columnSizes.Length; ii++)
                    {
                        //Debug.WriteLine("READ:" + ii + ":" + columnSizes[ii]);
                        if (columnSizes[ii] > 0)
                        {
                            writer.WriteStartElement("col");
                            writer.WriteAttributeString("min", ii.ToString());
                            writer.WriteAttributeString("max", ii.ToString());
                            writer.WriteAttributeString("width", columnSizes[ii].ToString());
                            writer.WriteAttributeString("style", "1");
                            writer.WriteEndElement();
                            //<col min="1" max="1" width="6.42578125" style="1" customWidth="1"/>
                        }
                    }
                    writer.WriteEndElement();
                }

                writer.WriteStartElement("sheetData");
                WriteWorksheetData(writer, data, lookupTable);
                writer.WriteEndElement();

                writer.WriteStartElement("pageMargins");
                writer.WriteAttributeString("left", "0.7");
                writer.WriteAttributeString("right", "0.7");
                writer.WriteAttributeString("top", "0.75");
                writer.WriteAttributeString("bottom", "0.75");
                writer.WriteAttributeString("header", "0.3");
                writer.WriteAttributeString("footer", "0.3");
                writer.WriteEndElement();

                writer.WriteEndElement();
            }
        }

        public static void WriteWorksheetData(XmlTextWriter writer, DataTable data, IDictionary<string, int> lookupTable)
        {
            var rowsCount = data.Rows.Count;
            var columnsCount = data.Columns.Count;
            string relPos;

            for (int row = 0; row < rowsCount; row++)
            {
                writer.WriteStartElement("row");
                relPos = RowIndexToName(row);
                writer.WriteAttributeString("r", relPos);
                writer.WriteAttributeString("spans", "1:" + columnsCount.ToString(CultureInfo.InvariantCulture));

                for (int column = 0; column < columnsCount; column++)
                {
                    object value = data.Rows[row][column];

                    writer.WriteStartElement("c");
                    relPos = RowColumnToPosition(row, column);
                    writer.WriteAttributeString("r", relPos);

                    writer.WriteAttributeString("s", "1");

                    var str = value as string;
                    if (str != null)
                    {
                        writer.WriteAttributeString("t", "s");
                        value = lookupTable[str];
                    }

                    writer.WriteElementString("v", value.ToString());

                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
            }
        }
    }
}

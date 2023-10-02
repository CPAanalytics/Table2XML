using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using ExcelDna.Integration;

namespace Table2XML
{
    public static class XmlConverterFunctions
    {
        [ExcelFunction(Description = "Converts a range to XML")]
        public static string ConvertToXml(string outerTag, string innerTag, [ExcelArgument(AllowReference = true)] object range)
        {
            try
            {
                // Check if the range object is an ExcelReference
                if (range is ExcelReference excelRef)
                {
                    // Get the values from the ExcelReference
                    object[,] rangeValues = (object[,])excelRef.GetValue();

                    return ConvertToXmlImpl(outerTag, innerTag, rangeValues);
                }

                return "Error: Invalid range provided.";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        private static string ConvertToXmlImpl(string outerTag, string innerTag, object[,] rangeValues)
        {
            if (rangeValues == null || rangeValues.Length == 0)
                return "Error: Provided range is empty.";

            StringBuilder xmlStr = new StringBuilder();

            // Start the XML with the outermost tag
            xmlStr.AppendLine($"<{outerTag}>");

            // Extract headers from the first row
            int numRows = rangeValues.GetLength(0);
            int numCols = rangeValues.GetLength(1);

            string[] headers = new string[numCols];
            for (int col = 0; col < numCols; col++)
            {
                if (rangeValues[0, col] == null || string.IsNullOrWhiteSpace(rangeValues[0, col].ToString()))
                    return $"Error: Header in column {col + 1} is empty or invalid.";

                headers[col] = Convert.ToString(rangeValues[0, col]);
            }

            // Loop through each row of the array (skip the header row)
            for (int row = 1; row < numRows; row++)
            {
                xmlStr.AppendLine($"<{innerTag}>");
                for (int col = 0; col < numCols; col++)
                {
                    xmlStr.AppendLine($"<{headers[col]}>{EscapeXML(Convert.ToString(rangeValues[row, col]))}</{headers[col]}>");
                }
                xmlStr.AppendLine($"</{innerTag}>");
            }

            // Close the outermost tag
            xmlStr.AppendLine($"</{outerTag}>");

            return xmlStr.ToString();
        }

        private static string EscapeXML(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;

            value = value.Replace("&", "&amp;");
            value = value.Replace("<", "&lt;");
            value = value.Replace(">", "&gt;");
            value = value.Replace("\"", "&quot;");
            value = value.Replace("'", "&apos;");
            return value;
        }
        [ExcelFunction(Description = "Converts XML in a cell to an Excel table")]
        public static string ConvertXmlToTable([ExcelArgument(AllowReference = true)] object cell)
        {
            try
            {
                if (cell is ExcelReference cellRef)
                {
                    object cellValue = cellRef.GetValue();
                    if (cellValue is string xmlString)
                    {
                        if (string.IsNullOrWhiteSpace(xmlString))
                            return "Error: Cell is empty or contains only whitespace.";

                        XDocument doc;
                        try
                        {
                            doc = XDocument.Parse(xmlString);
                        }
                        catch (Exception)
                        {
                            return "Error: Cell does not contain valid XML.";
                        }

                        return WriteXmlToTable(doc);
                    }
                    return "Error: Cell does not contain a string value.";
                }
                return "Error: Invalid cell reference.";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        private static string WriteXmlToTable(XDocument doc)
        {
            // This function assumes the XML has a consistent structure like:
            // <Items>
            //     <Item><Name>...</Name><Value>...</Value></Item>
            //     ...
            // </Items>

            var rows = doc.Descendants("Item").ToList();
            if (rows.Count == 0)
                return "Error: No 'Item' elements found in XML.";

            // Check for inconsistent XML data
            int expectedColumns = rows.First().Elements().Count();
            foreach (var row in rows)
            {
                if (row.Elements().Count() != expectedColumns)
                    return "Error: Inconsistent column count in XML rows.";
            }

            int currentRow = 1;  // Assuming starting from row 1, adjust if needed
            int currentCol = 1;  // Assuming starting from col 1 (A), adjust if needed

            // Write headers
            foreach (var header in rows.First().Elements())
            {
                ExcelReference headerCell = new ExcelReference(currentRow, currentCol);
                headerCell.SetValue(header.Name.LocalName);
                currentCol++;
            }

            currentRow++;

            // Write data
            foreach (var row in rows)
            {
                currentCol = 1;  // Reset column for each row
                foreach (var column in row.Elements())
                {
                    ExcelReference dataCell = new ExcelReference(currentRow, currentCol);
                    dataCell.SetValue(column.Value);
                    currentCol++;
                }
                currentRow++;
            }

            return "XML data written to table.";
        }

    }
}

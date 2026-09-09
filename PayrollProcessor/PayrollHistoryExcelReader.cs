using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PayrollProcessor
{
    internal static class PayrollHistoryExcelReader
    {
        public static List<string[]> ReadSheet(string path, string? preferredSheetName)
        {
            using SpreadsheetDocument document = SpreadsheetDocument.Open(path, false);
            WorkbookPart workbookPart = document.WorkbookPart
                ?? throw new InvalidOperationException("Workbook is missing in " + path);
            Sheet? sheet = FindSheet(workbookPart, preferredSheetName)
                ?? workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
            if (sheet?.Id == null)
            {
                throw new InvalidOperationException("No worksheet found in " + path);
            }

            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
            string[] sharedStrings = LoadSharedStrings(workbookPart);
            List<string[]> rows = new();
            int maxColumns = 0;
            foreach (Row row in worksheetPart.Worksheet.Descendants<Row>())
            {
                Dictionary<int, string> cells = new();
                foreach (Cell cell in row.Elements<Cell>())
                {
                    int column = ColumnIndex(cell.CellReference?.Value ?? "");
                    if (column <= 0)
                    {
                        continue;
                    }
                    cells[column] = GetCellText(cell, sharedStrings);
                    maxColumns = Math.Max(maxColumns, column);
                }
                if (cells.Count == 0)
                {
                    continue;
                }

                string[] values = new string[Math.Max(maxColumns, cells.Keys.DefaultIfEmpty(0).Max())];
                foreach ((int column, string value) in cells)
                {
                    if (column - 1 >= values.Length)
                    {
                        Array.Resize(ref values, column);
                    }
                    values[column - 1] = value;
                }
                rows.Add(values);
            }

            for (int i = 0; i < rows.Count; i++)
            {
                string[] row = rows[i];
                if (row.Length < maxColumns)
                {
                    Array.Resize(ref row, maxColumns);
                    rows[i] = row;
                }
                for (int c = 0; c < row.Length; c++)
                {
                    rows[i][c] ??= "";
                }
            }
            return rows;
        }

        private static Sheet? FindSheet(WorkbookPart workbookPart, string? preferredSheetName)
        {
            if (string.IsNullOrWhiteSpace(preferredSheetName))
            {
                return null;
            }

            return workbookPart.Workbook.Descendants<Sheet>()
                .FirstOrDefault(sheet => string.Equals(sheet.Name, preferredSheetName, StringComparison.OrdinalIgnoreCase));
        }

        private static string[] LoadSharedStrings(WorkbookPart workbookPart)
        {
            SharedStringTable? table = workbookPart.SharedStringTablePart?.SharedStringTable;
            if (table == null)
            {
                return Array.Empty<string>();
            }

            return table.Elements<SharedStringItem>()
                .Select(item => item.Text?.Text ?? item.InnerText ?? "")
                .ToArray();
        }

        private static string GetCellText(Cell cell, string[] sharedStrings)
        {
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                if (int.TryParse(cell.InnerText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index)
                    && index >= 0 && index < sharedStrings.Length)
                {
                    return sharedStrings[index].Trim();
                }
                return "";
            }

            if (cell.DataType != null && cell.DataType == CellValues.InlineString)
            {
                return (cell.InlineString?.Text?.Text ?? cell.InnerText ?? "").Trim();
            }

            if (cell.DataType != null && cell.DataType == CellValues.Boolean)
            {
                return cell.InnerText == "1" ? "TRUE" : "FALSE";
            }

            string raw = cell.CellValue?.Text ?? cell.InnerText ?? "";
            if (cell.DataType != null && cell.DataType == CellValues.Date
                && double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double oaDate))
            {
                return DateTime.FromOADate(oaDate).ToString("M/d/yyyy", CultureInfo.InvariantCulture);
            }

            return raw.Trim();
        }

        private static int ColumnIndex(string cellReference)
        {
            int column = 0;
            foreach (char character in cellReference)
            {
                if (!char.IsLetter(character))
                {
                    break;
                }
                column = column * 26 + (char.ToUpperInvariant(character) - 'A' + 1);
            }
            return column;
        }
    }

    internal static class PayrollHistoryValueParser
    {
        public static string Normalize(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "";
            }
            return Regex.Replace(value.Trim(), @"\s+", " ");
        }

        public static bool TryGetDate(string? value, out DateTime date)
        {
            date = DateTime.MinValue;
            string text = Normalize(value);
            if (text == "")
            {
                return false;
            }

            string[] formats =
            {
                "M/d/yyyy", "MM/dd/yyyy", "M/d/yy", "yyyy-MM-dd", "M-d-yyyy", "MM-dd-yyyy"
            };
            if (DateTime.TryParseExact(text, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out date)
                || DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out date)
                || DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out date))
            {
                date = date.Date;
                return true;
            }

            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double oaDate)
                && oaDate > 20000 && oaDate < 80000)
            {
                date = DateTime.FromOADate(oaDate).Date;
                return true;
            }

            return false;
        }

        public static bool TryGetInt(string? value, out int number)
        {
            number = 0;
            string text = Normalize(value).TrimStart('0');
            if (text == "")
            {
                return false;
            }
            return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out number);
        }

        public static float GetFloat(string? value)
        {
            string text = Normalize(value).Replace("$", "").Replace(",", "");
            if (text == "")
            {
                return 0f;
            }
            if (float.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out float number)
                || float.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out number))
            {
                return number;
            }
            return 0f;
        }

        public static void ParseName(string name, out string firstName, out string middleName, out string lastName)
        {
            firstName = "";
            middleName = "";
            lastName = "";
            name = Normalize(name);
            if (name == "")
            {
                return;
            }

            int comma = name.IndexOf(',');
            if (comma >= 0)
            {
                lastName = name[..comma].Trim();
                string[] given = name[(comma + 1)..].Split(' ', StringSplitOptions.RemoveEmptyEntries);
                if (given.Length > 0)
                {
                    firstName = given[0];
                }
                if (given.Length > 1)
                {
                    middleName = string.Join(" ", given.Skip(1));
                }
                return;
            }

            string[] parts = name.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 1)
            {
                firstName = parts[0];
            }
            else if (parts.Length >= 2)
            {
                firstName = parts[0];
                lastName = parts[^1];
                if (parts.Length > 2)
                {
                    middleName = string.Join(" ", parts.Skip(1).Take(parts.Length - 2));
                }
            }
        }
    }
}

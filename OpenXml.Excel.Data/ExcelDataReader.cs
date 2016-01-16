using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXml.Excel.Data.Util;
using Tuple = System.Tuple;

namespace OpenXml.Excel.Data
{
    public class ExcelDataReader : IDataReader
    {
        private SpreadsheetDocument _document;
        private OpenXmlReader _reader;
        private OpenXmlElement _currentRow;
        private readonly string[] _headers;
        private readonly IDictionary<int, string> _sharedStringTable;

        public ExcelDataReader(string path, int sheetIndex = 0, bool firstRowAsHeader = true)
        {
            _document = SpreadsheetDocument.Open(path, false);
            _sharedStringTable = _document.WorkbookPart.SharedStringTablePart.SharedStringTable
                .Select((x, i) => Tuple.Create(i, x.InnerText))
                .ToDictionary(x => x.Item1, x => x.Item2);

            var worksheetPart = _document.WorkbookPart.GetPartById(GetSheetByIndex(sheetIndex).Id.Value);
            _reader = OpenXmlReader.Create(worksheetPart);
            _headers = firstRowAsHeader ? GetFirstRowAsHeaders() : GetRangeHeaders(worksheetPart);
        }

        public ExcelDataReader(Stream stream, int sheetIndex = 0, bool firstRowAsHeader = true)
        {
            _document = SpreadsheetDocument.Open(stream, false);
            _sharedStringTable = _document.WorkbookPart.SharedStringTablePart.SharedStringTable
                .Select((x, i) => Tuple.Create(i, x.InnerText))
                .ToDictionary(x => x.Item1, x => x.Item2);

            var worksheetPart = _document.WorkbookPart.GetPartById(GetSheetByIndex(sheetIndex).Id.Value);
            _reader = OpenXmlReader.Create(worksheetPart);
            _headers = firstRowAsHeader ? GetFirstRowAsHeaders() : GetRangeHeaders(worksheetPart);
        }

        public ExcelDataReader(string path, string sheetName, bool firstRowAsHeader = true)
        {
            _document = SpreadsheetDocument.Open(path, false);
            _sharedStringTable = _document.WorkbookPart.SharedStringTablePart.SharedStringTable
                .Select((x, i) => Tuple.Create(i, x.InnerText))
                .ToDictionary(x => x.Item1, x => x.Item2);

            var worksheetPart = _document.WorkbookPart.GetPartById(GetSheetByName(sheetName).Id.Value);
            _reader = OpenXmlReader.Create(worksheetPart);
            _headers = firstRowAsHeader ? GetFirstRowAsHeaders() : GetRangeHeaders(worksheetPart);
        }

        public ExcelDataReader(Stream stream, string sheetName, bool firstRowAsHeader = true)
        {
            _document = SpreadsheetDocument.Open(stream, false);
            _sharedStringTable = _document.WorkbookPart.SharedStringTablePart.SharedStringTable
                .Select((x, i) => Tuple.Create(i, x.InnerText))
                .ToDictionary(x => x.Item1, x => x.Item2);

            var worksheetPart = _document.WorkbookPart.GetPartById(GetSheetByName(sheetName).Id.Value);
            _reader = OpenXmlReader.Create(worksheetPart);
            _headers = firstRowAsHeader ? GetFirstRowAsHeaders() : GetRangeHeaders(worksheetPart);
        }

        #region methods

        private IEnumerable<Sheet> GetSheets()
        {
            return _document.WorkbookPart.Workbook
                .GetFirstChild<Sheets>()
                .Elements<Sheet>();
        }

        private Sheet GetSheetByIndex(int sheetIndex)
        {
            var sheets = GetSheets().ToArray();
            if (sheetIndex < 0 || sheetIndex >= sheets.Count())
            {
                Dispose();
                throw new ApplicationException(Error.NotFoundSheetIndex(sheetIndex));
            }
            return sheets.ElementAt(sheetIndex);
        }

        private Sheet GetSheetByName(string sheetName)
        {
            var sheet = GetSheets().FirstOrDefault(x => x.Name == sheetName);
            if (sheet == null)
            {
                Dispose();
                throw new ApplicationException(Error.NotFoundSheetName(sheetName));
            }
            return sheet;
        }

        private string[] GetFirstRowAsHeaders()
        {
            string[] result;

            if (Read())
            {
                result = _currentRow.Elements<Cell>()
                    .Select(GetCellValue)
                    .ToArray();
            }
            else
                result = new string[] {};

            _currentRow = null;

            return result;
        }

        private static string[] GetRangeHeaders(OpenXmlPart worksheetPart)
        {
            var count = 0;
            using (var reader = OpenXmlReader.Create(worksheetPart))
            {
                while (reader.Read())
                {
                    if (reader.ElementType == typeof (Row))
                    {
                        count = reader.LoadCurrentElement().Elements<Cell>().Count();
                        break;
                    }
                }
            }
            return Enumerable.Range(0, count).Select(x => "col" + x).ToArray();
        }

        private string GetCellValue(CellType cell)
        {
            if (cell.CellValue == null)
                return null;

            var value = cell.CellValue.InnerXml;
            if (value == null)
                return null;

            int index;            
            if (int.TryParse(value, out index) && cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                return _sharedStringTable[index];

            return value;
        }

        private static IEnumerable<Cell> AdjustRow(OpenXmlElement row, int capacity)
        {
            if (row == null)
                return new Cell[] {};

            var cells = row.Elements<Cell>();

            var list = new Cell[capacity];
            foreach (var cell in cells)
                list[GetColumnIndexByName(cell.CellReference.Value)] = cell;

            return list;
        }

        private static int GetColumnIndexByName(string colName)
        {
            var name = Regex.Replace(colName, @"\d", "");

            int number = 0, pow = 1;
            for (var i = name.Length - 1; i >= 0; i--)
            {
                number += (name[i] - 'A' + 1) * pow;
                pow *= 26;
            }
            return number - 1;
        }

        #endregion

        #region IDataReader Members

        public void Close()
        {
            Dispose();
        }

        public int Depth
        {
            get { return 0; }
        }

        public DataTable GetSchemaTable()
        {
            return null;
        }

        public bool IsClosed
        {
            get { return _document == null; }
        }

        public bool NextResult()
        {
            return false;
        }

        public bool Read()
        {
            while (_reader.Read())
            {
                if (_reader.ElementType == typeof (Row))
                {
                    _currentRow = _reader.LoadCurrentElement();
                    break;
                }
            }
            return _currentRow != null && !_reader.EOF;
        }

        public int RecordsAffected
        {
            /*
             * RecordsAffected is only applicable to batch statements
             * that include inserts/updates/deletes. The sample always
             * returns -1.
             */
            get { return -1; }
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            if (_reader != null)
            {
                _reader.Close();
                _reader.Dispose();
                _reader = null;
            }

            if (_document != null)
            {
                _document.Dispose();
                _document = null;
            }
        }

        #endregion

        #region IDataRecord Members

        public int FieldCount
        {
            get { return _headers.Length; }
        }

        public bool GetBoolean(int i)
        {
            return SafeConverter.Convert<bool>(GetValue(i));
        }

        public byte GetByte(int i)
        {
            return SafeConverter.Convert<byte>(GetValue(i));
        }

        public long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferoffset, int length)
        {
            throw new NotImplementedException();
        }

        public char GetChar(int i)
        {
            return SafeConverter.Convert<char>(GetValue(i));
        }

        public long GetChars(int i, long fieldoffset, char[] buffer, int bufferoffset, int length)
        {
            throw new NotImplementedException();
        }

        public IDataReader GetData(int i)
        {
            return null;
        }

        public string GetDataTypeName(int i)
        {
            return typeof (string).Name;
        }

        public DateTime GetDateTime(int i)
        {
            return SafeConverter.Convert<DateTime>(GetValue(i));
        }

        public decimal GetDecimal(int i)
        {
            return SafeConverter.Convert<decimal>(GetValue(i));
        }

        public double GetDouble(int i)
        {
            return SafeConverter.Convert<double>(GetValue(i));
        }

        public Type GetFieldType(int i)
        {
            return typeof (string);
        }

        public float GetFloat(int i)
        {
            return SafeConverter.Convert<float>(GetValue(i));
        }

        public Guid GetGuid(int i)
        {
            return SafeConverter.Convert<Guid>(GetValue(i));
        }

        public short GetInt16(int i)
        {
            return SafeConverter.Convert<short>(GetValue(i));
        }

        public int GetInt32(int i)
        {
            return SafeConverter.Convert<int>(GetValue(i));
        }

        public long GetInt64(int i)
        {
            return SafeConverter.Convert<long>(GetValue(i));
        }

        public string GetName(int i)
        {
            return _headers[i];
        }

        public int GetOrdinal(string name)
        {
            for(var i = 0; i < _headers.Length; i++)
                if (string.Equals(_headers[i], name, StringComparison.InvariantCultureIgnoreCase))
                    return i;

            return -1;
        }

        public string GetString(int i)
        {
            return SafeConverter.Convert<string>(GetValue(i));
        }

        public object GetValue(int i)
        {
            var cell = AdjustRow(_currentRow, _headers.Length).ElementAtOrDefault(i);
            return GetCellValue(cell);
        }

        public int GetValues(object[] values)
        {
            var num = values.Length < _headers.Length ? values.Length : _headers.Length;
            var row = AdjustRow(_currentRow, num)
                .Select(x => x == null ? null : GetCellValue(x))
                .ToArray();

            for (var i = 0; i < num; i++)
                values[i] = row[i];

            return num;
        }

        public bool IsDBNull(int i)
        {
            return Convert.IsDBNull(GetValue(i));
        }

        public object this[string name]
        {
            get { return this[GetOrdinal(name)]; }
        }

        public object this[int i]
        {
            get { return GetValue(i); }
        }

        #endregion
    }
}
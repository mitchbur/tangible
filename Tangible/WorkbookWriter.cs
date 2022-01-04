using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Reflection.Metadata;
using System.Linq;

namespace Tangible
{
    public class WorkbookWriter : IDisposable
    {
        /// <summary>
        /// construct Workbook writer
        /// </summary>
        /// <param name="filepath">path to created file</param>
        /// <param name="tablestyle">spreadsheet table style</param>
        /// <param name="banded_rows">banded row table style</param>
        /// <param name="integer_format">integer column numbering format</param>
        /// <param name="scalar_format">scalar column numbering format</param>
        /// <param name="date_format">date column numbering format</param>
        public WorkbookWriter(string filepath, string tablestyle = "TableStyleLight1", bool banded_rows=true,
            string integer_format = @"#0;[Red]-#0", string scalar_format = @"#0.000",
            string date_format = @"yyyy\-mm\-dd\ hh:mm:ss")
        {
            table_style_ = tablestyle;
            banded_rows_ = banded_rows;

            doc_ = SpreadsheetDocument.
                Create(filepath, SpreadsheetDocumentType.Workbook, autoSave: false);
            PrepareWorkbook(doc_);

            style_codes_ = new Dictionary<TypeCode, Tuple<uint, uint>>();
            CreateUsefulNumberingFormats(integer_format: integer_format, scalar_format: scalar_format,
                date_format: date_format);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public void Close()
        {
            Dispose(true);
        }

        public void Save()
        {
            doc_.Save();
        }

        public void AddWorksheet(DataTable table)
        {
            AppendNewWorksheet(doc_.WorkbookPart, table);
        }

        public static void ExportToWorkbook(string filepath, List<DataTable> tables)
        {
            using (var writer_instance = new WorkbookWriter(filepath))
            {
                foreach (var table in tables)
                {
                    writer_instance.AddWorksheet(table);
                }
            }
        }

        #region private implementation

        protected bool is_disposed_ = false;
        private SpreadsheetDocument doc_;
        private Dictionary<TypeCode, Tuple<uint, uint>> style_codes_;
        protected virtual void Dispose(bool disposing)
        {
            if (is_disposed_) return;
            if (disposing)
            {
                doc_.Save();
                // doc_.WorkbookPart.Workbook.Save();
                doc_.Dispose();
            }
            this.is_disposed_ = true;
        }

        private void CreateUsefulNumberingFormats(string integer_format, string scalar_format, string date_format)
        {
            var integer_style_code = CreateUserNumberingFormat(doc_.WorkbookPart.Workbook, integer_format);
            style_codes_.Add(TypeCode.Int32, integer_style_code);
            var numeric_style_code = CreateUserNumberingFormat(doc_.WorkbookPart.Workbook, scalar_format);
            style_codes_.Add(TypeCode.Double, numeric_style_code);
            var datetime_style_code = CreateUserNumberingFormat(doc_.WorkbookPart.Workbook, date_format);
            style_codes_.Add(TypeCode.DateTime, datetime_style_code);
            var boolean_style_code = CreateUserNumberingFormat(doc_.WorkbookPart.Workbook, "\"TRUE\";\"TRUE\";\"FALSE\"");
            style_codes_.Add(TypeCode.Boolean, boolean_style_code);
        }

        private static void PrepareWorkbook(SpreadsheetDocument doc)
        {
            if (doc.WorkbookPart == null)
            {
                doc.AddWorkbookPart().Workbook = new Workbook();
            }
            var workbook_part = doc.WorkbookPart;

            var shared_strings_table_part = workbook_part.AddNewPart<SharedStringTablePart>();
            shared_strings_table_part.SharedStringTable = new SharedStringTable();
            shared_strings_table_part.SharedStringTable.Save();

            var styles_part = workbook_part.AddNewPart<WorkbookStylesPart>();
            styles_part.Stylesheet = new Stylesheet();
            var number_formats = styles_part.Stylesheet.AppendChild(new NumberingFormats());
            var stylesh_fonts = styles_part.Stylesheet.AppendChild(new Fonts());
            var stylesh_fills = styles_part.Stylesheet.AppendChild(new Fills());
            var stylesh_borders = styles_part.Stylesheet.AppendChild(new Borders());
            var cell_style_formats = styles_part.Stylesheet.AppendChild(new CellStyleFormats());
            var cell_formats = styles_part.Stylesheet.AppendChild(new CellFormats());
            var cell_styles = styles_part.Stylesheet.AppendChild(new CellStyles());
            var differential_formats = styles_part.Stylesheet.AppendChild(new DifferentialFormats());

            var def_font_size = new FontSize();
            def_font_size.Val = 11.0;
            var def_font_name = new FontName();
            def_font_name.Val = "Calibri";
            var def_font_family = new FontFamilyNumbering();
            def_font_family.Val = 2;
            var def_font_scheme = new FontScheme();
            def_font_scheme.Val = FontSchemeValues.Minor;
            var def_font = new Font();
            def_font.FontSize = def_font_size;
            def_font.FontName = def_font_name;
            def_font.FontFamilyNumbering = def_font_family;
            def_font.FontScheme = def_font_scheme;
            stylesh_fonts.Append(def_font);

            var def_fill = new Fill();
            var def_pattern_fill = new PatternFill();
            def_pattern_fill.PatternType = PatternValues.None;
            def_fill.PatternFill = def_pattern_fill;
            stylesh_fills.Append(def_fill);
            var grey_fill = new Fill();
            var grey_pattern_fill = new PatternFill();
            grey_pattern_fill.PatternType = PatternValues.Gray125;
            grey_fill.PatternFill = grey_pattern_fill;
            stylesh_fills.Append(grey_fill);

            var def_border = new Border();
            def_border.LeftBorder = new LeftBorder();
            def_border.RightBorder = new RightBorder();
            def_border.TopBorder = new TopBorder();
            def_border.BottomBorder = new BottomBorder();
            def_border.DiagonalBorder = new DiagonalBorder();
            def_border.DiagonalDown = false;
            def_border.DiagonalUp = false;
            stylesh_borders.Append(def_border);

            var def_cell_format = new CellFormat();
            def_cell_format.NumberFormatId = 0;
            def_cell_format.FontId = 0;
            def_cell_format.FillId = 0;
            def_cell_format.BorderId = 0;
            cell_style_formats.Append(def_cell_format);

            var general_cellfmt = new CellFormat();
            general_cellfmt.NumberFormatId = 0;
            general_cellfmt.FormatId = 0;
            cell_formats.Append(general_cellfmt);

            var def_cell_style = new CellStyle();
            def_cell_style.Name = "Normal";
            def_cell_style.FormatId = 0;
            def_cell_style.BuiltinId = 0;
            cell_styles.Append(def_cell_style);

            var diffr_format = new DifferentialFormat();
            diffr_format.Font = new Font();
            diffr_format.Font.Bold = new Bold() { Val = true };
            diffr_format.Font.Italic = new Italic() { Val = true };
            diffr_format.NumberingFormat = new NumberingFormat()
            {
                NumberFormatId = 0,
                FormatCode = "General"
            };
            differential_formats.Append(diffr_format);

            styles_part.Stylesheet.Save();

            workbook_part.Workbook.AppendChild(new Sheets());
            workbook_part.Workbook.AppendChild(new Tables());
        }

        private static uint NextAvailableSheetId_ = 1;
        private static uint AssignSheetId()
        {
            return NextAvailableSheetId_++;
        }

        // Number Format IDs less than 164 correspond to pre-defined numbering formats.
        private static uint NextAvailableNumberingFormatId_ = 164;
        private StringValue table_style_;
        private BooleanValue banded_rows_ = true;

        private static uint AssignUserNumerbingFormatId()
        {
            return NextAvailableNumberingFormatId_++;
        }

        /// <summary>
        /// create numbering format and cell format elements in stylesheet
        /// </summary>
        /// <param name="workbook">target workbook</param>
        /// <param name="formatcode">numbering format code, e.g. "#0.##"</param>
        /// <returns>Tuple pair:
        /// * Item1 : 0-based index of new cell format element in CellXfs list,
        /// use this as the "style" for a cell or column
        /// * Item2 : 0-based index of new differential format element,
        /// use this as a table column format index
        /// </returns>
        private static System.Tuple<uint, uint> CreateUserNumberingFormat(Workbook workbook, string formatcode)
        {
            var stylesheet = workbook.WorkbookPart.WorkbookStylesPart.Stylesheet;
            var number_formats = stylesheet.GetFirstChild<NumberingFormats>();
            var cell_formats = stylesheet.GetFirstChild<CellFormats>();
            var diffr_formats = stylesheet.GetFirstChild<DifferentialFormats>();
            var numbering_format_id = AssignUserNumerbingFormatId();

            var def_number_numfmt = new NumberingFormat();
            def_number_numfmt.FormatCode = formatcode;
            def_number_numfmt.NumberFormatId = numbering_format_id;
            number_formats.Append(def_number_numfmt);

            uint cell_formats_index = (uint)cell_formats.ChildElements.Count;
            var cell_format = new CellFormat();
            cell_format.NumberFormatId = numbering_format_id;
            cell_format.FormatId = 0;
            cell_format.ApplyNumberFormat = true;
            cell_formats.Append(cell_format);

            uint diffr_formats_index = (uint)diffr_formats.ChildElements.Count;
            var differential_format = new DifferentialFormat();
            differential_format.NumberingFormat = new NumberingFormat()
            {
                NumberFormatId = numbering_format_id,
                FormatCode = formatcode
            };
            diffr_formats.Append(differential_format);
            stylesheet.Save();

            return new Tuple<uint, uint>(cell_formats_index, diffr_formats_index);
        }

        private void AppendNewWorksheet(WorkbookPart workbook_part, DataTable table)
        {
            var shared_string_table = workbook_part.Workbook.GetFirstChild<SharedStringTable>();
            var worksheet_part = workbook_part.AddNewPart<WorksheetPart>();
            worksheet_part.Worksheet = new Worksheet();

            var table_definition_part = worksheet_part.AddNewPart<TableDefinitionPart>();
            table_definition_part.Table = new Table();

            Columns cols = new Columns();
            var table_columns = new TableColumns();

            foreach (DataColumn column in table.Columns)
            {
                uint col_pos = (uint)(column.Ordinal + 1);
                var typecode = Type.GetTypeCode(column.DataType);
                Column col_definition = new Column() { CustomWidth = true };
                col_definition.Min = col_pos;
                col_definition.Max = col_pos;

                var t_col = new TableColumn();
                t_col.Id = col_pos;
                t_col.Name = column.ColumnName;

                switch (typecode)
                {
                    case TypeCode.Int32:
                        col_definition.Width = 12;
                        col_definition.Style = style_codes_[typecode].Item1;
                        t_col.DataFormatId = style_codes_[typecode].Item2;
                        break;
                    case TypeCode.Double:
                        col_definition.Width = 14;
                        col_definition.Style = style_codes_[typecode].Item1;
                        t_col.DataFormatId = style_codes_[typecode].Item2;
                        break;
                    case TypeCode.DateTime:
                        col_definition.Width = 25;
                        col_definition.Style = style_codes_[typecode].Item1;
                        t_col.DataFormatId = style_codes_[typecode].Item2;
                        break;
                    case TypeCode.Boolean:
                        col_definition.Width = 12;
                        col_definition.Style = style_codes_[typecode].Item1;
                        t_col.DataFormatId = style_codes_[typecode].Item2;
                        break;
                    default:
                        col_definition.Width = 15;
                        col_definition.Style = 0;
                        break;
                }

                cols.Append(col_definition);
                table_columns.Append(t_col);
            }
            worksheet_part.Worksheet.Append(cols);

            var sheet_data = TableToSheetData(table);
            worksheet_part.Worksheet.Append(sheet_data);
            var sheet = new Sheet()
            {
                Id = workbook_part.GetIdOfPart(worksheet_part),
                SheetId = AssignSheetId(),
                Name = table.TableName
            };
            workbook_part.Workbook.Sheets.Append(sheet);

            var num_rows = (uint)sheet_data.ChildElements.Count();
            var num_cols = (uint)sheet_data.ChildElements.First<Row>().Count();
            string cell_range_ref = CellRangeReference(num_rows, num_cols);

            table_definition_part.Table.Name = table.TableName;
            table_definition_part.Table.Reference = cell_range_ref;
            table_definition_part.Table.DisplayName = table.TableName;
            table_definition_part.Table.TotalsRowShown = false;
            table_definition_part.Table.Id = sheet.SheetId;
            table_definition_part.Table.TableColumns = table_columns;
            table_definition_part.Table.TableStyleInfo = new TableStyleInfo()
            {
                Name = this.table_style_,
                ShowRowStripes = this.banded_rows_,
                ShowColumnStripes = false,
                ShowFirstColumn = false,
                ShowLastColumn = false
            };
            table_definition_part.Table.HeaderRowFormatId = 0;
            table_definition_part.Table.AutoFilter = new AutoFilter() { Reference = cell_range_ref };
            var wb_tables = workbook_part.Workbook.GetFirstChild<Tables>();
            wb_tables.Append(table_definition_part.Table);
            table_definition_part.Table.Save();

            string table_relationship_id = worksheet_part.GetIdOfPart(table_definition_part);
            var table_parts = new TableParts();
            var table_part = new TablePart() { Id = table_relationship_id };
            table_parts.Append(table_part);
            worksheet_part.Worksheet.Append(table_parts);

            worksheet_part.Worksheet.SheetDimension = new SheetDimension() { Reference = cell_range_ref };

            worksheet_part.Worksheet.Save();
        }

        private static readonly uint num_column_letters_ = (uint)'Z' - (uint)'A' + 1;
        private static string ColumnReference(uint col)
        {
            var dividend = col - 1;

            if (dividend < num_column_letters_)
            {
                // single letter representation
                return ((char)((uint)'A' + dividend)).ToString();
            }

            var rems = new List<uint>();
            while (dividend >= num_column_letters_)
            {
                var remainder = dividend % num_column_letters_;
                dividend = dividend / num_column_letters_;
                rems.Insert(0, remainder);
            }
            rems.Insert(0, dividend);
            var letters = rems.Select(x => (char)((uint)'A' + x));
            var colref = string.Concat(letters);
            return colref;
        }

        private static string CellRangeReference(uint row, uint col, uint num_rows, uint num_cols)
        {
            var last_row = row + num_rows - 1;
            var last_col = col + num_cols - 1;
            return ColumnReference(col) + row.ToString() + ":" +
                ColumnReference(last_col) + last_row.ToString();
        }
        private static string CellRangeReference(uint num_rows, uint num_cols)
        {
            return "A1:" + ColumnReference(num_cols) + num_rows.ToString();
        }

        private SheetData TableToSheetData(DataTable table)
        {
            var sheet_data = new SheetData();
            var header_row = new Row();

            var col_type_codes = new List<TypeCode>();
            foreach (System.Data.DataColumn column in table.Columns)
            {
                col_type_codes.Add(Type.GetTypeCode(column.DataType));
                var cell = new Cell();
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(column.ColumnName);
                header_row.AppendChild(cell);
            }

            var col_cell_handlers = new List<IColumnTypeHandler>();
            foreach (System.Data.DataColumn column in table.Columns)
            {
                TypeCode type_code = Type.GetTypeCode(column.DataType);
                switch (type_code)
                {
                    case TypeCode.String:
                        col_cell_handlers.Add(new StringColumnTypeHandler());
                        break;
                    case TypeCode.Boolean:
                        col_cell_handlers.Add(new BooleanColumnTypeHandler(style_codes_[type_code].Item1));
                        break;
                    case TypeCode.Int32:
                        col_cell_handlers.Add(new IntegerColumnTypeHandler(style_codes_[type_code].Item1));
                        break;
                    case TypeCode.Double:
                        col_cell_handlers.Add(new ScalarColumnTypeHandler(style_codes_[type_code].Item1));
                        break;
                    case TypeCode.DateTime:
                        col_cell_handlers.Add(new DateColumnTypeHandler(style_codes_[type_code].Item1));
                        break;
                }
            }

            sheet_data.AppendChild(header_row);

            foreach (DataRow tbl_data_row in table.Rows)
            {
                var sheet_row = new Row();
                var col_handler_enum = col_cell_handlers.GetEnumerator();
                foreach (var item in tbl_data_row.ItemArray)
                {
                    col_handler_enum.MoveNext();
                    var cell = col_handler_enum.Current.ConstructCell(item);
                    sheet_row.AppendChild(cell);
                }
                sheet_data.AppendChild(sheet_row);
            }

            return sheet_data;
        }

    }

    interface IColumnTypeHandler
    {
        public Cell ConstructCell(object item);
    }

    class StringColumnTypeHandler : IColumnTypeHandler
    {
        public StringColumnTypeHandler()
        {
        }

        public Cell ConstructCell(object item)
        {
            var cell = new Cell();
            cell.DataType = CellValues.String;
            // cell.StyleIndex is not set, use default XML attribute
            if ((item.GetType() != typeof(DBNull)))
            {
                cell.CellValue = new CellValue((string)item);
            }
            return cell;
        }
    }

    class ScalarColumnTypeHandler : IColumnTypeHandler
    {
        public ScalarColumnTypeHandler(uint styleindex)
        {
            styleindex_ = styleindex;
        }

        public Cell ConstructCell(object item)
        {
            var cell = new Cell();
            cell.DataType = CellValues.Number;
            cell.StyleIndex = styleindex_;
            if ((item.GetType() != typeof(DBNull)))
            {
                cell.CellValue = new CellValue((double)item);
            }
            return cell;
        }

        private uint styleindex_;
    }

    class IntegerColumnTypeHandler : IColumnTypeHandler
    {
        public IntegerColumnTypeHandler(uint styleindex)
        {
            styleindex_ = styleindex;
        }

        public Cell ConstructCell(object item)
        {
            var cell = new Cell();
            cell.DataType = CellValues.Number;
            cell.StyleIndex = styleindex_;
            if ((item.GetType() != typeof(DBNull)))
            {
                cell.CellValue = new CellValue((int)item);
            }
            return cell;
        }

        private uint styleindex_;
    }

    class BooleanColumnTypeHandler : IColumnTypeHandler
    {
        public BooleanColumnTypeHandler(uint styleindex)
        {
            styleindex_ = styleindex;
        }

        public Cell ConstructCell(object item)
        {
            var cell = new Cell();
            cell.DataType = CellValues.Boolean;
            cell.StyleIndex = styleindex_;
            if ((item.GetType() != typeof(DBNull)))
            {
                int value = (bool)item ? 1 : 0;
                cell.CellValue = new CellValue(value);
            }
            return cell;
        }

        private uint styleindex_;
    }

    class DateColumnTypeHandler : IColumnTypeHandler
    {
        public DateColumnTypeHandler(uint styleindex)
        {
            styleindex_ = styleindex;
        }

        public Cell ConstructCell(object item)
        {
            var cell = new Cell();
            cell.DataType = CellValues.Date;
            cell.StyleIndex = styleindex_;
            if ((item.GetType() != typeof(DBNull)))
            {
                cell.CellValue = new CellValue((DateTime)item);
            }
            return cell;
        }

        private uint styleindex_;
    }


    #endregion
}

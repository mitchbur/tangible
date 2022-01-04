using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using LumenWorks.Framework.IO.Csv;

namespace Tangible
{
    public class TsvReader
    {
        public static DataTable ReadTsvFile(string filepath)
        {
            // use root file name as table name
            string table_name = Path.ChangeExtension(Path.GetFileName(filepath), "").TrimEnd('.');
            var table = new DataTable(table_name);
            using (var txt_reader = new System.IO.StreamReader(filepath, encoding: System.Text.Encoding.UTF8))
            {
                var tsv_reader = new CsvReader(txt_reader, hasHeaders: true, delimiter: '\t');
                table.Load(tsv_reader);
            }

            CheckColumnNames(table);

            // all columns are read in as 'string' type
            // convert columns to various data types
            var out_table = ConvertColumnTypes(table);
            return out_table;
        }

        private static void CheckColumnNames( DataTable table )
        {
            foreach ( var col in table.Columns )
            {
                var col_name_good = col.ToString().Select(c => c >= (uint)7).All(b => b);
                if ( ! col_name_good )
                {
                    throw new InvalidDataException("column name contains invalid characters: '" +
                        col.ToString() + "'");
                }
            }
        }

        private static DataTable ConvertColumnTypes(DataTable table)
        {
            var column_types = GuessColumnTypes(table);

            DataTable cnv_table = table.Clone();
            foreach (var col_idx in Enumerable.Range(0, cnv_table.Columns.Count))
            {
                cnv_table.Columns[col_idx].DataType = column_types[col_idx];
            }
            foreach (DataRow row in table.Rows)
            {
                cnv_table.ImportRow(row);
            }
            cnv_table.AcceptChanges();
            return cnv_table;
        }

        private static Type[] GuessColumnTypes(DataTable table, int max_check_depth = 65)
        {
            var col_types = new Type[table.Columns.Count];
            // var col_type_codes = new TypeCode[table.Columns.Count];
            Array.Fill(col_types, typeof(Object));
            var col_values_dictionaries = new List<Dictionary<string, uint>>();
            foreach (var j in Enumerable.Range(0, table.Columns.Count))
            {
                col_values_dictionaries.Add(new Dictionary<string, uint>());
            }
            var selected_row_indexes = SampleChunkedRaggedGaps(table.Rows.Count, max_check_depth);
            foreach (var row_index in selected_row_indexes)
            {
                var data_row = table.Rows[row_index];
                foreach (var j in Enumerable.Range(0, col_types.Length))
                {
                    string cell_val = data_row[j].ToString().Trim().ToUpper();
                    if (IsNA(cell_val)) continue;

                    var value_dictionary = col_values_dictionaries[j];
                    if (value_dictionary.ContainsKey(cell_val))
                    {
                        value_dictionary[cell_val] += 1;
                    }
                    else
                    {
                        value_dictionary.Add(cell_val, 1);
                    }
                }
            }

            foreach (var j in Enumerable.Range(0, col_values_dictionaries.Count))
            {
                var value_dictionary = col_values_dictionaries[j];

                if (value_dictionary.Count == 0)
                {
                    continue;
                }

                if (value_dictionary.Count <= 2)
                {
                    if (value_dictionary.Keys.Select(x => IsBoolean(x)).All(x => x))
                    {
                        col_types[j] = typeof(bool);
                    }
                }

                foreach (var item in value_dictionary)
                {
                    switch (Type.GetTypeCode(col_types[j]))
                    {
                        case TypeCode.Boolean:
                        case TypeCode.String:
                            break;
                        case TypeCode.Int32:
                            if ( !IsInteger(item.Key))
                            {
                                if (IsDouble(item.Key))
                                {
                                    col_types[j] = typeof(double);
                                }
                                else
                                {
                                    col_types[j] = typeof(string);
                                }
                            }
                            break;
                        case TypeCode.Double:
                            if (!IsDouble(item.Key))
                            {
                                // inconsistency so column must be string
                                col_types[j] = typeof(string);
                            }
                            break;
                        case TypeCode.DateTime:
                            if (!IsDate(item.Key))
                            {
                                // inconsistency so column must be string
                                col_types[j] = typeof(string);
                            }
                            break;
                        case TypeCode.Object:
                            // not yet determined any type for column
                            if (IsDate(item.Key))
                            {
                                col_types[j] = typeof(DateTime);
                            }
                            else if (IsDouble(item.Key))
                            {
                                if ( IsInteger(item.Key))
                                {
                                    col_types[j] = typeof(int);
                                }
                                else
                                {
                                    col_types[j] = typeof(double);
                                }
                            }
                            else
                            {
                                col_types[j] = typeof(string);
                            }
                            break;
                    }
                }
            }
            return col_types;
        }

        private static bool IsNA(string s)
        {
            return ( s.Length == 0 );
        }
        private static bool IsDouble(string s )
        {
            double dummy_value;
            return Double.TryParse(s, out dummy_value);
        }

        private static bool IsInteger(string s)
        {
            int dummy_value;
            return Int32.TryParse(s, out dummy_value);
        }


        private static readonly CultureInfo datetime_culture = CultureInfo.InvariantCulture;
        private static readonly DateTimeStyles datetime_style = DateTimeStyles.AdjustToUniversal | 
            DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal;
        private static bool IsDate(string s)
        {
            DateTime dummy_datetime;
            return DateTime.TryParse(s, datetime_culture, datetime_style, out dummy_datetime);
        }

        private static bool IsBoolean(string s)
        {
            bool dummy_value;
            return Boolean.TryParse(s, out dummy_value);
        }

        /// <summary>
        /// sample a range, in chunks with ragged gaps between chunks
        /// </summary>
        /// <param name="num_items">range length</param>
        /// <param name="num_samples">number of samples to generate</param>
        /// <param name="chunk_size">chunk size</param>
        /// <returns>selections from range [0,num_items), in increasing order</returns>
        public static int[] SampleChunkedRaggedGaps( int num_items, int num_samples, int chunk_size=5 )
        {
            if ( ( num_items <= 0 ) || ( num_samples <= 0 ) || ( chunk_size <= 0 ) )
            {
                return new int[0];
            }
            var num_ignored = num_items - num_samples;
            if ( num_ignored < 1 )
            {
                // more samples requested than available, return the full range
                return Enumerable.Range(0, (int)num_items).ToArray();
            }
            if (num_samples <= chunk_size)
            {
                // number of samples is only one chunk
                // select from the beginning of the range
                return Enumerable.Range(0, (int)num_samples).ToArray();
            }

            // select chunks in range with jittered gap sizes
            var num_chunks = ( num_samples + chunk_size - 1 ) / chunk_size;
            var num_gaps = num_chunks - 1;
            var mid_spacing = num_ignored / num_gaps;
            var jitter = ( mid_spacing + 1 ) / 2;
            int[] jittered_spacing = new int[3] { mid_spacing, mid_spacing - jitter, mid_spacing + jitter };

            int row_index = 0;
            int sarr_index = 0;
            var selected = new int[num_samples];
            List<int> chunk_start = new List<int>();
            for ( int space_idx = 0; space_idx < num_gaps; ++space_idx )
            {
                for ( int k = 0; k < chunk_size; ++k )
                {
                    selected[sarr_index] = row_index + k;
                    ++sarr_index;
                }
                chunk_start.Add(row_index);
                var space_width = jittered_spacing[space_idx % 3];
                row_index += chunk_size + space_width;
            }
            var rem_sel = num_samples - sarr_index;
            row_index = num_items - rem_sel;
            for ( int k = 0; k < rem_sel; ++k )
            {
                selected[sarr_index] = row_index + k;
                ++sarr_index;
            }
            return selected;
        }
    }
}

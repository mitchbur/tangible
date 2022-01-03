using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace Tangible
{
    /// <summary>
    /// tab-separated-variable text file writer
    /// </summary>
    public class TsvWriter : IDisposable
    {
        /// <summary>
        /// constructor
        /// </summary>
        /// <param name="delimiter">character separating row items</param>
        /// <param name="quote">character used to quote string values</param>
        /// <param name="escape">character used to escape special characters in strings</param>
        /// <param name="nullvalue">value written for null row items</param>
        /// <param name="encoding">output stream text encoding</param>
        public TsvWriter(char delimiter = '\t', char quote = '\'', char escape = '\\', string nullvalue = "",
            Encoding encoding = null)
        {
            delimiter_ = delimiter;
            quote_ = quote;
            escape_ = escape;
            nullvalue_ = nullvalue;
            if (encoding == null)
            {
                encoding_ = System.Text.Encoding.UTF8;
            }
            else
            {
                encoding_ = encoding;
            }
        }

        public void Dispose()
        {
            // nothing special needed here
        }

        /// <summary>
        /// export DataTable to TSV file
        /// </summary>
        /// <param name="table">source data table</param>
        /// <param name="filepath">path to new file</param>
        /// <param name="header">if true, header row is written</param>
        public void Write(DataTable table, string filepath, bool header = true)
        {
            using (var txt_writer = new System.IO.StreamWriter(filepath,
                append: false, encoding: encoding_))
            {
                if (header)
                {
                    bool need_separator = false;
                    foreach (DataColumn col in table.Columns)
                    {
                        if (need_separator)
                        {
                            txt_writer.Write(delimiter_);
                        }
                        need_separator = true;

                        txt_writer.Write(col.ColumnName);
                    }
                    txt_writer.WriteLine();
                }

                var col_types = new List<Type>();
                foreach (System.Data.DataColumn column in table.Columns)
                {
                    col_types.Add(column.DataType);
                }

                foreach (DataRow row in table.Rows)
                {
                    bool need_separator = false;
                    for (int j = 0; j < col_types.Count; ++j)
                    {
                        var item = row.ItemArray[j];
                        if (need_separator)
                        {
                            txt_writer.Write(delimiter_);
                        }
                        need_separator = true;

                        if (item == DBNull.Value)
                        {
                            txt_writer.Write(nullvalue_);
                        }
                        else if (col_types[j] == typeof(string))
                        {
                            txt_writer.Write(EnquoteString(item as string));
                        }
                        else if (col_types[j] == typeof(DateTime))
                        {
                            DateTime date = ((DateTime)item).ToUniversalTime();
                            txt_writer.Write(date.ToString("u"));
                        }
                        else
                        {
                            txt_writer.Write(item.ToString());
                        }
                    }
                    txt_writer.WriteLine();
                }
                txt_writer.Close();
            }
        }

        private char delimiter_;
        private char quote_;
        private char escape_;
        private string nullvalue_;
        private System.Text.Encoding encoding_;

        private string EnquoteString(string value)
        {
            // TODO
            return value;
        }
    }
}

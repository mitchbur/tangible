using NUnit.Framework;
using Tangible;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace TestTangible
{
    class TestsWorkbookWriter
    {
        [SetUp]
        public void Setup()
        {
        }

        public static string MakeRandomString( int length )
        {
            var random_generator = new System.Random();
            var char_seq = Enumerable.Range(0, length).
                Select(k => (char)('a' + random_generator.Next(0, 25)) ).
                ToArray();
            return string.Concat(char_seq);
        }

        [Test]
        public void TestAddWorksheet()
        {
            var in_filepath = @"..\..\..\..\data\interesting_data.txt";
            var out_filepath = MakeRandomString(4) + ".xlsx";
            var table = ReadTestData(in_filepath);

            var writer = new WorkbookWriter(out_filepath);
            writer.AddWorksheet(table);
            writer.Close();

            var info = new System.IO.FileInfo(out_filepath);
            Assert.True(info.Exists, "output file exists");
            Assert.Greater(info.Length, 0, "output file not empty");
        }

        [Test]
        public void TestExport()
        {
            var in_filepath = @"..\..\..\..\data\interesting_data.txt";
            var out_filepath = MakeRandomString(4) + ".xlsx";
            var table = ReadTestData(in_filepath);
            WorkbookWriter.ExportToWorkbook(out_filepath, new List<DataTable> { table });

            var info = new System.IO.FileInfo(out_filepath);
            Assert.True(info.Exists, "output file exists");
            Assert.Greater(info.Length, 0, "output file not empty");
        }

        public static DataTable ReadTestData( string filepath )
        {
            DataTable test_datatable = null;

            try
            {
                test_datatable = TsvReader.ReadTsvFile(filepath);
            }
            catch (Exception ex)
            {
                Assert.Fail("ReadTestData failed: " + ex.Message);
            }

            return test_datatable;
        }
    }
}

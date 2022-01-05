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
        private string input_tsv_filepath_ = null;

        private void CreateInputDataFile()
        {
            this.input_tsv_filepath_ = DataGenerator.MakeRandomString(4) + ".txt";
            var table = DataGenerator.GenerateExampleTableOne();
            using (var writer = new TsvWriter())
            {
                writer.Write(table, input_tsv_filepath_);
            }
        }

        [SetUp]
        public void Setup()
        {
        }

        [OneTimeSetUp]
        public void OneTimeSetup()
        {
            CreateInputDataFile();
        }

        [Test]
        public void TestAddWorksheet()
        {
            var out_filepath = DataGenerator.MakeRandomString(4) + ".xlsx";
            var table = ReadTestData(this.input_tsv_filepath_);

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
            var out_filepath = DataGenerator.MakeRandomString(4) + ".xlsx";
            var table = ReadTestData(this.input_tsv_filepath_);
            WorkbookWriter.ExportToWorkbook(out_filepath, new List<DataTable> { table });

            var info = new System.IO.FileInfo(out_filepath);
            Assert.True(info.Exists, "output file exists");
            Assert.Greater(info.Length, 0, "output file not empty");
        }

        [Test]
        public void TestColumnReference( )
        {
            TestColumnReferenceValue(1, "A");
            TestColumnReferenceValue(2, "B");
            TestColumnReferenceValue(25, "Y");
            TestColumnReferenceValue(26, "Z");
            TestColumnReferenceValue(27, "AA");     //          1*26 + 1
            TestColumnReferenceValue(28, "AB");     //          1*26 + 2
            TestColumnReferenceValue(51, "AY");     //          1*26 + 25
            TestColumnReferenceValue(52, "AZ");     //          1*26 + 26
            TestColumnReferenceValue(53, "BA");     //          2*26 + 1
            TestColumnReferenceValue(54, "BB");     //          2*26 + 2
            TestColumnReferenceValue(77, "BY");     //          2*26 + 25
            TestColumnReferenceValue(78, "BZ");     //          2*26 + 26
            TestColumnReferenceValue(651, "YA");    //         25*26 + 1
            TestColumnReferenceValue(652, "YB");    //         25*26 + 2
            TestColumnReferenceValue(675, "YY");    //         25*26 + 25
            TestColumnReferenceValue(676, "YZ");    //         25*26 + 26
            TestColumnReferenceValue(677, "ZA");    //         26*26 + 1
            TestColumnReferenceValue(678, "ZB");    //         26*26 + 2
            TestColumnReferenceValue(701, "ZY");    //         26*26 + 25
            TestColumnReferenceValue(702, "ZZ");    //         26*26 + 26
            TestColumnReferenceValue(703, "AAA");   // 1*676 +  1*26 + 1
            TestColumnReferenceValue(704, "AAB");   // 1*676 +  1*26 + 2
            TestColumnReferenceValue(727, "AAY");   // 1*676 +  1*26 + 25
            TestColumnReferenceValue(728, "AAZ");   // 1*676 +  1*26 + 26
            TestColumnReferenceValue(729, "ABA");   // 1*676 +  2*26 + 1
            TestColumnReferenceValue(730, "ABB");   // 1*676 +  2*26 + 2
            TestColumnReferenceValue(753, "ABY");   // 1*676 +  2*26 + 25
            TestColumnReferenceValue(754, "ABZ");   // 1*676 +  2*26 + 26
            TestColumnReferenceValue(1353, "AZA");  // 1*676 + 26*26 + 1
            TestColumnReferenceValue(1354, "AZB");  // 1*676 + 26*26 + 2
            TestColumnReferenceValue(1377, "AZY");  // 1*676 + 26*26 + 25
            TestColumnReferenceValue(1378, "AZZ");  // 1*676 + 26*26 + 26
            TestColumnReferenceValue(1379, "BAA");  // 2*676 +  1*26 + 1
            TestColumnReferenceValue(1380, "BAB");  // 2*676 +  1*26 + 2
            TestColumnReferenceValue(1403, "BAY");  // 2*676 +  1*26 + 25
            TestColumnReferenceValue(1404, "BAZ");  // 2*676 +  1*26 + 26
        }

        public static void TestColumnReferenceValue(uint value, string expected)
        {
            string result = WorkbookWriter.ColumnReference(value);
            Assert.AreEqual(expected, result, "ColumnReference( {0} )", value );
        }

        public static DataTable ReadTestData(string filepath)
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

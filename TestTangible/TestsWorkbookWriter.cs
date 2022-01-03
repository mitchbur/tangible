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

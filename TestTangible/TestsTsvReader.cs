using NUnit.Framework;
using Tangible;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace TestTangible
{
    public class TestsTsvReader
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void ReadExample()
        {
            string example_filepath = DataGenerator.MakeRandomString(4) + ".txt";
            var expected_table = DataGenerator.GenerateExampleTableOne();
            using (var writer = new TsvWriter())
            {
                writer.Write(expected_table, example_filepath);
            }

            var readback_table = TsvReader.ReadTsvFile(example_filepath);
        }

        [Test]
        public void Test_SampleChunkedRaggedGaps_ZeroAvailable()
        {
            int num_items = 0;
            int expected_length = 0;

            for (int requested_samples = 0; requested_samples <= 2; ++requested_samples)
            {
                var result = TsvReader.SampleChunkedRaggedGaps(num_items, requested_samples);
                Assert.AreEqual(expected_length, result.Length, "zero available items");
            }
        }

        [Test]
        public void Test_SampleChunkedRaggedGaps_ZeroChunkSize()
        {
            int chunk_size = 0;
            int num_items = 100;

            {
                int expected_length = 0;
                var result = TsvReader.SampleChunkedRaggedGaps(num_items, expected_length, chunk_size);
                Assert.AreEqual(expected_length, result.Length, "zero samples requested");
            }

            for (int requested_length = 1; requested_length <= num_items; ++requested_length)
            {
                int expected_length = 0;
                var result = TsvReader.SampleChunkedRaggedGaps(num_items, requested_length, chunk_size);
                Assert.AreEqual(expected_length, result.Length, "chunk size is zero" );
            }
        }

        [Test]
        public void Test_SampleChunkedRaggedGaps()
        {
            int chunk_size = 5;
            int num_items = 5 * chunk_size;

            {
                int expected_length = 0;
                var result = TsvReader.SampleChunkedRaggedGaps(num_items, expected_length, chunk_size);
                Assert.AreEqual(expected_length, result.Length, "zero samples requested");
            }

            for ( int expected_length = 1; expected_length <= num_items; ++expected_length )
            {
                var result = TsvReader.SampleChunkedRaggedGaps(num_items, expected_length, chunk_size);
                Assert.AreEqual(expected_length, result.Length, "single chunk sample length");
                CheckChunksAndGaps(result, num_items, chunk_size);
            }

            {
                int requested_length = num_items + 1;
                int expected_length = num_items;
                var result = TsvReader.SampleChunkedRaggedGaps(num_items, requested_length, chunk_size);
                Assert.AreEqual(expected_length, result.Length, "limit to number of available items");
            }
        }

        [Test]
        public void Test_SampleChunkedRaggedGaps_Size1()
        {
            int chunk_size = 1;
            int num_items = 100;

            {
                int expected_length = 0;
                var result = TsvReader.SampleChunkedRaggedGaps(num_items, expected_length, chunk_size);
                Assert.AreEqual(expected_length, result.Length, "zero samples requested");
            }

            for (int expected_length = 1; expected_length <= num_items; ++expected_length)
            {
                var result = TsvReader.SampleChunkedRaggedGaps(num_items, expected_length, chunk_size);
                Assert.AreEqual(expected_length, result.Length, "single chunk sample length");
                CheckChunksAndGaps(result, num_items, chunk_size);
            }

            {
                int requested_length = num_items + 1;
                int expected_length = num_items;
                var result = TsvReader.SampleChunkedRaggedGaps(num_items, requested_length, chunk_size);
                Assert.AreEqual(expected_length, result.Length, "limit to number of available items");
            }
        }

        private void CheckChunksAndGaps( int[] seq, int n, int expected_chunk_size )
        {
            var chunks = new List<int>();
            var gaps = new List<int>();

            var ignored_count = n - seq.Length;
            var expected_num_chunks = ( seq.Length + expected_chunk_size - 1 )/ expected_chunk_size;
            var min_gap = Math.Max(0, ignored_count / expected_num_chunks / 2);

            int current_index = 0;
            int chunk_width = 0;
            int gap_width = 0;
            foreach ( var val in seq)
            {
                Assert.Less(val, n, "sequence value exceeds upper limit");

                Assert.GreaterOrEqual(val, current_index, "sequence not increasing");
                gap_width = val - current_index;
                if ( gap_width > 0 )
                {
                    if ( chunk_width > 0 )
                    {
                        chunks.Add(chunk_width);
                    }
                    Assert.GreaterOrEqual(gap_width, min_gap, "gap too small" );
                    gaps.Add(gap_width);
                    chunk_width = 1;
                    current_index = val + 1;
                }
                else // gap is 0
                {
                    ++chunk_width;
                    ++current_index;
                    if (chunk_width > expected_chunk_size)
                    {
                        // degenerate zero width gap
                        Assert.GreaterOrEqual(gap_width, min_gap, "0 width gap not expected");
                        gaps.Add(gap_width);
                        chunks.Add(expected_chunk_size);
                        chunk_width -= expected_chunk_size;
                    }
                }
            }
            gap_width = n - current_index;
            if ( chunk_width > 0 )
            {
                chunks.Add(chunk_width);
            }
            gaps.Add(gap_width);

            int num_chunks = chunks.Count();
            Assert.AreEqual(expected_num_chunks, num_chunks);

            int num_gaps = gaps.Count();
            Assert.AreEqual(num_chunks, num_gaps, "mis-count of gaps");

            if ( num_chunks > 1 )
            {
                Assert.AreEqual(0, gaps.Last(), 
                    "more than one chunk and last chunk was not at end of sequence");
            }

            // all except the last chunk should be the chunk size
            foreach ( var chunk_size in chunks.Take(num_chunks-1) )
            {
                Assert.AreEqual(expected_chunk_size, chunk_size, "incorrect chunk size");
            }

            return;
        }
    }
}
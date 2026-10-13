using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// Pins the pair-range flattening rules and the size guard
    /// (PLAN.md §1.3, S3.3): a 2-column range is read row-wise, a 2-row range
    /// column-wise, anything else throws the frozen shape error; a count
    /// mismatch between keys and values throws its own frozen error; an empty
    /// result throws "no entries to write"; an oversized range is rejected
    /// before allocation.
    /// </summary>
    public class NatsRangeHelperTests
    {
        private static object[,] Range(int rows, int cols)
        {
            return new object[rows, cols];
        }

        // ------------------------------------------------------------------
        // FlattenPairRange(range): 2 columns = vertical pairs.
        // ------------------------------------------------------------------

        [Fact]
        public void FlattenPairRange_TwoColumnsIsReadRowWise()
        {
            // A 2x2 range with cols == 2: rows are the pairs.
            object[,] range = { { "a", "1" }, { "b", "2" } };

            List<KeyValuePair<string, string>> pairs = NatsUdfCommon.FlattenPairRange(range);

            Assert.Equal(2, pairs.Count);
            Assert.Equal(new KeyValuePair<string, string>("a", "1"), pairs[0]);
            Assert.Equal(new KeyValuePair<string, string>("b", "2"), pairs[1]);
        }

        [Fact]
        public void FlattenPairRange_TwoRowsIsReadColumnWise()
        {
            // A 2x3 range (rows == 2, cols == 3): columns are the pairs.
            object[,] range = { { "a", "b", "c" }, { "1", "2", "3" } };

            List<KeyValuePair<string, string>> pairs = NatsUdfCommon.FlattenPairRange(range);

            Assert.Equal(3, pairs.Count);
            Assert.Equal(new KeyValuePair<string, string>("a", "1"), pairs[0]);
            Assert.Equal(new KeyValuePair<string, string>("b", "2"), pairs[1]);
            Assert.Equal(new KeyValuePair<string, string>("c", "3"), pairs[2]);
        }

        [Fact]
        public void FlattenPairRange_AColumnRangeIsReadRowWise()
        {
            // 3 rows, 2 columns.
            object[,] range = { { "a", "1" }, { "b", "2" }, { "c", "3" } };

            List<KeyValuePair<string, string>> pairs = NatsUdfCommon.FlattenPairRange(range);

            Assert.Equal(new[] { "a", "b", "c" }, pairs.Select(p => p.Key).ToArray());
        }

        [Fact]
        public void FlattenPairRange_WrongShapeThrowsTheFrozenError()
        {
            var ex = Assert.Throws<ArgumentException>(
                () => NatsUdfCommon.FlattenPairRange(Range(3, 3)));
            Assert.Equal(NatsUdfCommon.ErrorRangeShape, ex.Message);
        }

        [Fact]
        public void FlattenPairRange_NullRangeIsRequired()
        {
            var ex = Assert.Throws<ArgumentException>(
                () => NatsUdfCommon.FlattenPairRange((object[,])null));
            Assert.Equal(NatsUdfCommon.ErrorRangeRequired, ex.Message);
        }

        [Fact]
        public void FlattenPairRange_EmptyResultThrowsNoEntries()
        {
            // A 2-column range whose keys are all truly missing flattens to
            // nothing and must report "no entries to write".
            object[,] range = { { null, "1" }, { ExcelMissing.Value, "2" } };

            var ex = Assert.Throws<ArgumentException>(() => NatsUdfCommon.FlattenPairRange(range));
            Assert.Equal(NatsUdfCommon.ErrorNoEntries, ex.Message);
        }

        [Fact]
        public void FlattenPairRange_MissingValueBecomesEmptyText()
        {
            object[,] range = { { "a", null }, { "b", ExcelEmpty.Value } };

            List<KeyValuePair<string, string>> pairs = NatsUdfCommon.FlattenPairRange(range);

            Assert.Equal(2, pairs.Count);
            Assert.Equal("", pairs[0].Value);
            Assert.Equal("", pairs[1].Value);
        }

        [Fact]
        public void FlattenPairRange_ExcelErrorCellThrowsTheFrozenError()
        {
            object[,] range = { { "a", ExcelError.ExcelErrorValue } };

            var ex = Assert.Throws<ArgumentException>(() => NatsUdfCommon.FlattenPairRange(range));
            Assert.Equal(NatsUdfCommon.ErrorExcelErrorCell, ex.Message);
        }

        // ------------------------------------------------------------------
        // FlattenPairRange(keys, values): two parallel ranges.
        // ------------------------------------------------------------------

        [Fact]
        public void FlattenPairRange_TwoRangesIsRowMajor()
        {
            object[,] keys = { { "a", "b" }, { "c", "d" } };
            object[,] values = { { "1", "2" }, { "3", "4" } };

            List<KeyValuePair<string, string>> pairs = NatsUdfCommon.FlattenPairRange(keys, values);

            Assert.Equal(4, pairs.Count);
            Assert.Equal(new KeyValuePair<string, string>("a", "1"), pairs[0]);
            Assert.Equal(new KeyValuePair<string, string>("b", "2"), pairs[1]);
            Assert.Equal(new KeyValuePair<string, string>("c", "3"), pairs[2]);
            Assert.Equal(new KeyValuePair<string, string>("d", "4"), pairs[3]);
        }

        [Fact]
        public void FlattenPairRange_CountMismatchThrowsTheFrozenError()
        {
            object[,] keys = { { "a", "b" } };      // 2 cells
            object[,] values = { { "1", "2", "3" } }; // 3 cells

            var ex = Assert.Throws<ArgumentException>(() => NatsUdfCommon.FlattenPairRange(keys, values));
            Assert.Equal(NatsUdfCommon.ErrorCellCountMismatch, ex.Message);
        }

        [Fact]
        public void FlattenPairRange_NullKeysOrValuesIsRequired()
        {
            object[,] keys = { { "a" } };
            Assert.Equal(
                NatsUdfCommon.ErrorRangeRequired,
                Assert.Throws<ArgumentException>(
                    () => NatsUdfCommon.FlattenPairRange((object[,])null, keys)).Message);
            Assert.Equal(
                NatsUdfCommon.ErrorRangeRequired,
                Assert.Throws<ArgumentException>(
                    () => NatsUdfCommon.FlattenPairRange(keys, null)).Message);
        }

        // ------------------------------------------------------------------
        // EnsureRangeFits: reject before allocating.
        // ------------------------------------------------------------------

        [Fact]
        public void EnsureRangeFits_AcceptsAReasonableRange()
        {
            NatsUdfCommon.EnsureRangeFits(Range(10, 2));
            NatsUdfCommon.EnsureRangeFits(1000, 1000);
        }

        [Fact]
        public void EnsureRangeFits_RejectsAnOversizedRange()
        {
            // Beyond the worksheet row limit.
            Assert.Equal(
                NatsUdfCommon.ErrorRangeTooLarge,
                Assert.Throws<ArgumentException>(
                    () => NatsUdfCommon.EnsureRangeFits(2_000_000, 2)).Message);

            // Beyond the worksheet column limit.
            Assert.Equal(
                NatsUdfCommon.ErrorRangeTooLarge,
                Assert.Throws<ArgumentException>(
                    () => NatsUdfCommon.EnsureRangeFits(2, 100_000)).Message);

            // Within each dimension but over the total-cell budget
            // (10_000 x 10_000 = 100M > the 50M total-cell cap).
            Assert.Equal(
                NatsUdfCommon.ErrorRangeTooLarge,
                Assert.Throws<ArgumentException>(
                    () => NatsUdfCommon.EnsureRangeFits(10_000, 10_000)).Message);
        }

        [Fact]
        public void EnsureRangeFits_NullIsANoOp()
        {
            NatsUdfCommon.EnsureRangeFits((object[,])null);
        }

        // ------------------------------------------------------------------
        // RequireRange: presence, size and Excel-error rejection.
        // ------------------------------------------------------------------

        [Fact]
        public void RequireRange_AcceptsAUsableRange()
        {
            Assert.Null(NatsUdfCommon.RequireRange(Range(2, 2)));
        }

        [Fact]
        public void RequireRange_NullOrEmptyIsRequired()
        {
            Assert.Equal(
                NatsUdfCommon.ErrorPrefix + "a range is required",
                NatsUdfCommon.RequireRange(null));
            Assert.Equal(
                NatsUdfCommon.ErrorPrefix + "a range is required",
                NatsUdfCommon.RequireRange(new object[0, 0]));
        }

        [Fact]
        public void RequireRange_RejectsAnExcelErrorCell()
        {
            object[,] range = { { ExcelError.ExcelErrorNA } };
            Assert.Equal(
                NatsUdfCommon.ErrorPrefix + NatsUdfCommon.ErrorExcelErrorCell,
                NatsUdfCommon.RequireRange(range));
        }
    }
}

using System;
using ExcelDna.Integration;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// <see cref="NatsJson"/> matrix converters. The shape contract is pinned by
    /// the XML doc: a 1x1 range serializes as a scalar, a single row as a flat
    /// array, anything else as an array of row arrays, and the shape round-trips
    /// through <see cref="NatsJson.JsonToMatrix"/>.
    /// </summary>
    public class NatsJsonTests
    {
        // ------------------------------------------------------------------
        // MatrixToJson shape.
        // ------------------------------------------------------------------

        [Fact]
        public void MatrixToJson_OneByOneRange_SerializesAsAScalar()
        {
            Assert.Equal("5", NatsJson.MatrixToJson(new object[,] { { 5.0 } }));
            Assert.Equal("\"hi\"", NatsJson.MatrixToJson(new object[,] { { "hi" } }));
        }

        [Fact]
        public void MatrixToJson_SingleRow_SerializesAsAFlatArray()
        {
            Assert.Equal("[1,2,3]", NatsJson.MatrixToJson(new object[,] { { 1.0, 2.0, 3.0 } }));
        }

        [Fact]
        public void MatrixToJson_MultipleRows_SerializesAsRowArrays()
        {
            Assert.Equal("[[1,2],[3,4]]",
                NatsJson.MatrixToJson(new object[,] { { 1.0, 2.0 }, { 3.0, 4.0 } }));
        }

        [Fact]
        public void MatrixToJson_SingleColumnMultipleRows_SerializesAsRowArrays()
        {
            // Not a "single row", so it keeps the row-array shape.
            Assert.Equal("[[1],[2]]",
                NatsJson.MatrixToJson(new object[,] { { 1.0 }, { 2.0 } }));
        }

        [Fact]
        public void MatrixToJson_NullRange_ReturnsAnErrorCell()
        {
            string json = NatsJson.MatrixToJson(null);
            Assert.StartsWith("Error:", json);
        }

        [Fact]
        public void MatrixToJson_NonFiniteDouble_IsRejectedWithAClearError()
        {
            string nan = NatsJson.MatrixToJson(new object[,] { { double.NaN } });
            string infinity = NatsJson.MatrixToJson(new object[,] { { double.PositiveInfinity } });
            Assert.StartsWith("Error:", nan);
            Assert.StartsWith("Error:", infinity);
        }

        // ------------------------------------------------------------------
        // JsonToMatrix shape.
        // ------------------------------------------------------------------

        [Fact]
        public void JsonToMatrix_Scalar_ProducesAOneByOneMatrix()
        {
            object[,] matrix = NatsJson.JsonToMatrix("5");
            Assert.Equal(1, matrix.GetLength(0));
            Assert.Equal(1, matrix.GetLength(1));
            Assert.Equal(5L, matrix[0, 0]);
        }

        [Fact]
        public void JsonToMatrix_FlatArray_ProducesOneRow()
        {
            object[,] matrix = NatsJson.JsonToMatrix("[1,2,3]");
            Assert.Equal(1, matrix.GetLength(0));
            Assert.Equal(3, matrix.GetLength(1));
            Assert.Equal(1L, matrix[0, 0]);
            Assert.Equal(3L, matrix[0, 2]);
        }

        [Fact]
        public void JsonToMatrix_ArrayOfArrays_ProducesRows()
        {
            object[,] matrix = NatsJson.JsonToMatrix("[[1,2],[3,4]]");
            Assert.Equal(2, matrix.GetLength(0));
            Assert.Equal(2, matrix.GetLength(1));
            Assert.Equal(1L, matrix[0, 0]);
            Assert.Equal(4L, matrix[1, 1]);
        }

        [Fact]
        public void JsonToMatrix_JsonNullElement_MapsToTheNullValueArgument()
        {
            object[,] matrix = NatsJson.JsonToMatrix("[1,null,3]", "N/A");
            Assert.Equal("N/A", matrix[0, 1]);
            Assert.Equal(1L, matrix[0, 0]);
            Assert.Equal(3L, matrix[0, 2]);
        }

        [Fact]
        public void JsonToMatrix_NullValueDefaultsToBlankText()
        {
            object[,] matrix = NatsJson.JsonToMatrix("[1,null,3]");
            Assert.Equal(string.Empty, matrix[0, 1]);
        }

        [Fact]
        public void JsonToMatrix_NonFiniteDouble_IsRejectedWithAClearError()
        {
            // Newtonsoft accepts the non-standard literal NaN; the converter
            // rejects it so a cell never receives a #NUM!.
            object[,] matrix = NatsJson.JsonToMatrix("NaN");
            Assert.Equal(1, matrix.GetLength(0));
            Assert.Equal(1, matrix.GetLength(1));
            var text = matrix[0, 0] as string;
            Assert.NotNull(text);
            Assert.StartsWith("Error:", text);
        }

        [Fact]
        public void JsonToMatrix_InvalidJson_ReturnsAnErrorCell()
        {
            object[,] matrix = NatsJson.JsonToMatrix("{not valid");
            Assert.StartsWith("Error:", (string)matrix[0, 0]);
        }

        [Fact]
        public void JsonToMatrix_NullDashWhitespaceInput_ReturnsTheFillValue()
        {
            Assert.Equal(string.Empty, NatsJson.JsonToMatrix(null)[0, 0]);
            Assert.Equal(string.Empty, NatsJson.JsonToMatrix("   ")[0, 0]);
            Assert.Equal("X", NatsJson.JsonToMatrix("\uFEFF  ", "X")[0, 0]);
        }

        // ------------------------------------------------------------------
        // Round-trip.
        // ------------------------------------------------------------------

        [Fact]
        public void RoundTrip_OneByOneStaysScalar()
        {
            object[,] matrix = NatsJson.JsonToMatrix(NatsJson.MatrixToJson(new object[,] { { 7.0 } }));
            Assert.Equal(1, matrix.GetLength(0));
            Assert.Equal(1, matrix.GetLength(1));
            Assert.Equal(7L, matrix[0, 0]);
        }

        [Fact]
        public void RoundTrip_SingleRowStaysOneRow()
        {
            object[,] source = { { 1.0, 2.0, 3.0 } };
            object[,] matrix = NatsJson.JsonToMatrix(NatsJson.MatrixToJson(source));
            Assert.Equal(1, matrix.GetLength(0));
            Assert.Equal(3, matrix.GetLength(1));
        }

        [Fact]
        public void RoundTrip_MultipleRowsStayRowArrays()
        {
            object[,] source = { { 1.0, 2.0 }, { 3.0, 4.0 } };
            object[,] matrix = NatsJson.JsonToMatrix(NatsJson.MatrixToJson(source));
            Assert.Equal(2, matrix.GetLength(0));
            Assert.Equal(2, matrix.GetLength(1));
            Assert.Equal(1L, matrix[0, 0]);
            Assert.Equal(4L, matrix[1, 1]);
        }
    }
}

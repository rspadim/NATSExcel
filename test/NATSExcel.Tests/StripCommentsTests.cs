using System;
using Newtonsoft.Json.Linq;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// <see cref="AppConfig.StripComments"/> must be string-aware: a <c>//</c> or
    /// <c>/*</c> that sits inside a JSON string is content, not a comment, which
    /// is what makes <c>{"u":"nats://h:4222"}</c> survive. These are pure tests
    /// over the input text, so no config fixture is needed.
    /// </summary>
    public class StripCommentsTests
    {
        [Fact]
        public void UrlInsideAString_IsNotTreatedAsAComment()
        {
            string input = "{\"u\":\"nats://h:4222\"}";
            Assert.Equal(input, AppConfig.StripComments(input));
        }

        [Fact]
        public void BlockCommentStartInsideAString_IsNotAComment()
        {
            string input = "{\"u\":\"a/*b\"}";
            Assert.Equal(input, AppConfig.StripComments(input));
        }

        [Fact]
        public void BackslashEscapeInsideAString_IsRespected()
        {
            // The \" is an escaped quote, so the following // stays in the string.
            string input = "\"a\\\"//b\"";
            Assert.Equal(input, AppConfig.StripComments(input));

            // A JSON parse of a document using that escape still keeps the value.
            var root = JObject.Parse(AppConfig.StripComments("{\"u\":\"a\\\"//b\"}"));
            Assert.Equal("a\"//b", (string)root["u"]);
        }

        [Fact]
        public void LineCommentToEndOfLine_IsRemovedAndNewlineKept()
        {
            string stripped = AppConfig.StripComments("{\"a\":1} // note\n{\"b\":2}");
            Assert.DoesNotContain("note", stripped);
            Assert.Contains("{\"b\":2}", stripped);
            Assert.Contains("\n", stripped);
        }

        [Fact]
        public void BlockComment_IsRemoved()
        {
            string stripped = AppConfig.StripComments("{\"a\":1} /* note */ {\"b\":2}");
            Assert.DoesNotContain("note", stripped);
            Assert.DoesNotContain("/*", stripped);
            Assert.Contains("{\"a\":1}", stripped);
            Assert.Contains("{\"b\":2}", stripped);
        }

        [Fact]
        public void UnterminatedBlockComment_DropsTheRestWithoutThrowing()
        {
            // Documented/observed behavior: an unterminated /* */ swallows the
            // remainder of the text and never throws (the loader then reports a
            // JSON parse error for the truncated document).
            string stripped = null;
            Exception thrown = Record.Exception(
                () => { stripped = AppConfig.StripComments("{\"a\":1} /* never closed"); });
            Assert.Null(thrown);
            Assert.Equal("{\"a\":1} ", stripped);
        }

        [Fact]
        public void EmptyAndNullInput_DoNotThrow()
        {
            Assert.Equal(string.Empty, AppConfig.StripComments(string.Empty));
            // A null input would throw inside the StringBuilder ctor; assert the
            // actual behavior rather than assume a null-safe contract.
            Exception thrown = Record.Exception(() => AppConfig.StripComments(null));
            Assert.NotNull(thrown);
        }
    }
}

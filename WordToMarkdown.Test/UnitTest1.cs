using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace WordToMarkdown.Test
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            string tmpFileName = Path.GetTempFileName();

            string inputPath = Path.GetFullPath("HeadingTest.docx");

            WordToMarkdown.Program p = new WordToMarkdown.Program(inputPath, tmpFileName);

            string expectedFath = Path.GetFullPath("HeadingTest.md");

            bool equal = EqualTextFiles(expectedFath, tmpFileName);

            System.IO.File.Delete(tmpFileName);
        }

        private bool EqualTextFiles(string pathname1, string pathname2)
        {
            bool same = File.ReadLines(pathname1).SequenceEqual(File.ReadLines(pathname2));

            return same;
        }
    }
}

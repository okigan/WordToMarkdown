using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace WordToMarkdown.Test
{
    [TestClass]
    public class TestWordToMarkdown
    {
        [TestMethod]
        public void TestBasics()
        {
            string[] files = { "lorem.docx", "basics.docx", "headings.docx", "lists.docx" };

            foreach (string file in files)
            {
                string inputPath = Path.GetFullPath(file);
                string expectedName = Path.GetFileNameWithoutExtension(inputPath) + ".md";
                string expectedFath = Path.GetFullPath(expectedName);

                string tmpFileName = Path.GetTempFileName();

                {
                    WordToMarkdown.Program p = new WordToMarkdown.Program(inputPath, tmpFileName);

                    // give some time to word to close down
                    System.Threading.Thread.Sleep(100);
                }

                bool equal = EqualTextFiles(expectedFath, tmpFileName);

                System.IO.File.Delete(tmpFileName);

                Assert.IsTrue(equal);
            }
        }

        private bool EqualTextFiles(string pathname1, string pathname2)
        {
            bool same = File.ReadLines(pathname1).SequenceEqual(File.ReadLines(pathname2));

            return same;
        }
    }
}

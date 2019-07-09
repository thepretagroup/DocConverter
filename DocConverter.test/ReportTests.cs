using NUnit.Framework;
using System.Collections.Generic;
using System.Linq;

namespace DocConverter.Tests
{
    public class ReportTests
    {
        private IList<string> paragraphs;

        [SetUp]
        public void Setup()
        {
            paragraphs = new List<string>()
            {
                "a",
                "b",
                "Xeloda",
                "d",
                "e",
                "Xeloda 555",
                "g",
                "Xeloda 1234",
                "567 Xeloda"
            };

        }

        [TestCase(0, 4, 0)]
        [TestCase(1, 4, 0)]
        [TestCase(2, 4, 0)]
        [TestCase(3, 4, 1)]
        [TestCase(4, 4, 1)]
        [TestCase(5, 4, 1)]
        [TestCase(6, 4, 2)]
        [TestCase(7, 4, 2)]
        [TestCase(8, 4, 3)]
        [TestCase(9, 4, 4)]
        public void XelodaRemovalHackTest(int currentParagraph, int expectedBefore, int expectedAfter)
        {
            var report = new Report();

            var xelodasBefore = paragraphs.ToList().Where(p => p.Contains("Xeloda"));

            Assert.AreEqual(expectedBefore, xelodasBefore.Count());

            report.XelodaRemovalHack(paragraphs, ref currentParagraph);

            var xelodasAfter = paragraphs.ToList().Where(p => p.Contains("Xeloda"));

            Assert.AreEqual(expectedAfter, xelodasAfter.Count());
        }
    }
}
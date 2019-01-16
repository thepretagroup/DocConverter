using NUnit.Framework;


namespace DocConverter.test
{
    class DrugEffectTests
    {
        [Test]
        public void DrugEffectTest()
        {
            //var de = new DrugEffect(new string[] { "tylenol", "123", "ug/ml", "Resistant" });
            var de = new DrugEffect("tylenol", "123", "ug/ml", "Resistant");

            Assert.IsNotNull(de);
            Assert.AreEqual("tylenol", de.Drug);
            Assert.AreEqual("123", de.IC50);
            Assert.AreEqual("ug/ml", de.Units);
            Assert.AreEqual("Resistant", de.Interpretation);
        }

        [TestCase("Resistant","Lower")]
        [TestCase("Inactive","Lower")]
        [TestCase("Sensitive", "Higher")]
        [TestCase("Intermediate", "Average")]
        [TestCase("Moderately Active", "Average")]
        [TestCase("Active", "?????")]
        [TestCase("Non Existent Value", "?")] 
        [TestCase("", "?")]
        [TestCase(null, "?")]
        public void ExVivoInterpretationTest(string interprectation, string ExpectedExVivoInterpretation)
        {
            var de = new DrugEffect("tylenol", "123", "ug/ml", interprectation);

            Assert.AreEqual(ExpectedExVivoInterpretation, de.ExVivoInterpretation);
        }

        [Test]
        public void MultiDrugEffectTest()
        {
            var mde = new MultiDrugEffect("tylenol", "1:2", "123", "ug/ml", "Resistant","N/A" );

            Assert.IsNotNull(mde);
            Assert.AreEqual("tylenol", mde.Drug);
            Assert.AreEqual("1:2", mde.Ratio);
            Assert.AreEqual("123", mde.IC50);
            Assert.AreEqual("ug/ml", mde.Units);
            Assert.AreEqual("Resistant", mde.Interpretation);
            Assert.AreEqual("N/A", mde.Synergy);
        }

        [TestCase("Resistant", "Lower")]
        [TestCase("Inactive", "Lower")]
        [TestCase("Sensitive", "Higher")]
        [TestCase("Intermediate", "Average")]
        [TestCase("Moderately Active", "Average")]
        [TestCase("Non Existent Value", "?")]
        [TestCase("", "?")]
        [TestCase(null, "?")]
        public void ExVivoInterpretationMultiTest(string interprectation, string ExpectedExVivoInterpretation)
        {
            var mde = new MultiDrugEffect("tylenol", "1:2", "123", "ug/ml", interprectation, "N/A");

            Assert.AreEqual(ExpectedExVivoInterpretation, mde.ExVivoInterpretation);
        }


    }
}

// Copyright (c) Bodoconsult EDV-Dienstleistungen GmbH. All rights reserved.


using Bodoconsult.Core.Pdf.Helpers;
using MigraDoc.DocumentObjectModel;
using NUnit.Framework;

namespace Bodoconsult.Core.Pdf.Test
{
    [TestFixture]
    public class UnitTestObjectHelper
    {
        [Test]
        public void TestMapProperties()
        {
            var style1 = new Style("Test1", "Normal") {Font = {Size = 25}};


            var style2 = new Style("Test2", "Normal");

            Assert.IsTrue(style1.Font.Size!=style2.Font.Size);

            ObjectHelper.MapProperties(style1, style2);
            ObjectHelper.MapProperties(style1.Font, style2.Font);

            Assert.IsTrue(style1.Font.Size == style2.Font.Size);
        }
    }
}
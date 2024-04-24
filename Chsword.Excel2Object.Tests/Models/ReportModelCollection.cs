using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace Chsword.Excel2Object.Tests.Models;

internal class ReportModelCollection : List<ReportModel>
{
    public void AreEqual(ICollection<ReportModel> reportModels)
    {
        Assert.AreEqual(Count, reportModels.Count);
        var index = 0;
        foreach (var model1 in reportModels)
        {
            var model2 = this[index];
            Assert.AreEqual(model2.Title, model1.Title);
            Assert.AreEqual(model2.Name, model1.Name);
            Assert.AreEqual(model2.Enabled, model1.Enabled);
            index++;
        }
    }
}
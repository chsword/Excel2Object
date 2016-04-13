using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests.Models
{
    class ReportModelCollection : List<ReportModel>
    {
        public void AreEqual(ICollection<ReportModel> reportModels)
        {
            Assert.AreEqual(this.Count,reportModels.Count);
            var index = 0;
            foreach (var model1 in reportModels)
            {
                var model2 = this[index];
                Assert.AreEqual(model2.Title,model1.Title);
                Assert.AreEqual(model2.Name, model1.Name);
                Assert.AreEqual(model2.Enabled, model1.Enabled);
                index++;
            }
        }
    }
}

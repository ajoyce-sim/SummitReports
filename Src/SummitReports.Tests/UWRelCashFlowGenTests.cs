using System;
using Xunit;
using SummitReports.Objects;
using System.Collections.Generic;

namespace SummitReports.Tests
{
    public class UWRelCashFlowGenTests
    {
        [Fact]
        public void ReportGenOk()
        {
            var xlsFinancial = new NPV();
            
            /* This will allow you to set the formula that will modify the rate before it is used in the calculation,  if left out, it will just use the number passed in the first parameter of calculate */
            //xlsFinancial.RateFormula = string.Format("(({0}+1)^(1/12))-1", NPV.RATE_VAR);

            List<double> inputValues = new List<double>();
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(1020882.11);
            Assert.True(xlsFinancial.Calculate(0.22, inputValues).ToString("0.00").Equals("890152.51"), "NPV with .22 not correct 890152.51");
            Assert.True(xlsFinancial.Calculate(0.18, inputValues).ToString("0.00").Equals("932217.54"), "NPV with .18 not correct 932217.54");
            Assert.True(xlsFinancial.Calculate(0.15, inputValues).ToString("0.00").Equals("966142.30"), "NPV with .15 not correct 966142.30");

            /* this will write out the resulting excel file to let you play around with it */
            //xlsFinancial.SaveToFile(@"\\vmware-host\Shared Folders\Downloads\npv.xls");
        }
    }
}

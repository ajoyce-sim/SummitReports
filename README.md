# SummitReports - Summit Report Object Library

Object library that is used by other applcations in the company that need report generation.  If follows the model of using NPOI and excep templates to generate
excel reports

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. 

### Creating a report 
1. Right click on SummitReports.Objects\Reports\ModelReport and copy the entire folder with contents.  Copy it under the same folder under the SummitReports.Objects project (it should now be ModelReport - copy,  rename it to "SampleReport",  rename the class ModelReport.cs to 'SampleReport.cs' and ModelReportTemplate.xlsx to 'SampleReportTemplate.xlsx'.  Rename the class from 'ModelReport' to 'SampleReport'
2.  Rename : 
```public class ModelReport : ModelReportBaseObject```
to
```public class SampleReport : SummitReportBaseObject```
```
        public ModelReport() : base(@"ModelReport\ModelReportTemplate.xlsx")
```		
to
```
        public SampleReport() : base(@"SampleReport\SampleReportTemplate.xlsx")
```		

To Test this,  go to SummitReports.Tests class UWRelCashFlowGenTests.cs copy and paste function ModelReportGenTestOk (be sure the include [fact] function attribute)
It should look like this, rename it to SampleReportGenTestOk.
```
        [Fact]
        public async void SampleReportGenTestOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = new SummitReport();
            var generatedFIleName = await rpt.GenerateAsync(13);
        }
```
go To Test Explorer (Menu bar Test->Windows->Test Explorer), right click and select the test  If this runs successfully, open windows explorer and type %TEMP% and press enter (this will open up your temp directory which should have your newly geneated report).

To create the nuget package for this,  navigate to the solution root and type powershell,  type ./build.ps1 and press enter.  


These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. 

### Using this project:

Please see the xunit .Tests project included to see samples of how this is used.  

## Deployment

This project was design to be hosted and distributed with nuget.com.

## Built With

* [.net standard](https://www.microsoft.com/net/learn/get-started) - The framework used

## Contributing

Please read [CONTRIBUTING.md](https://gist.github.com/rvegajr/651875c08acb76009e563db128f33e7e) for details on our code of conduct, and the process for submitting pull requests to us.

## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/rvegajr/tags). 

## Authors

* **Ricky Vega** - *Initial work* - [Noctusoft](https://github.com/rvegajr)

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

Many thanks to the following projects that have helped in this project
* NPOI


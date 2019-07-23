#tool nuget:?package=NUnit.ConsoleRunner&version=3.4.0
#tool nuget:?package=vswhere
#addin "Cake.WebDeploy"

var IncrementMinorVersion = true;
var NuGetReleaseNotes = new [] {".netcore2.2 target nuget deploy fix", "VS019", "Added Sentence Casing", "Updated by adding StringFormat helper that lets the template chain string functions such as 'lower,snake,title,trim', etc", "Upgraded all Nuget Packages"};

DirectoryPath vsLatest  = VSWhereLatest();
FilePath msBuildPathX64 = (vsLatest==null)
                            ? null
                            : vsLatest.CombineWithFilePath("./MSBuild/15.0/Bin/amd64/MSBuild.exe");
var target = Argument("target", "Default");
//////////////////////////////////////////////////////////////////////
// ARGUMENTS
//////////////////////////////////////////////////////////////////////
var dirSep = System.IO.Path.DirectorySeparatorChar;

var configuration = Argument("configuration", "Release");

var thisDir = System.IO.Path.GetFullPath(".") + dirSep;
var srcDir = System.IO.Path.GetFullPath(Directory("./Src/SummitReports.Objects/"));
var binDir = System.IO.Path.GetFullPath(Directory("./Src/SummitReports.Objects/bin/"));
var projectFile = System.IO.Path.GetFullPath(Directory("./Src/SummitReports.Objects/SummitReports.Objects.csproj"));
var solutionFile = System.IO.Path.GetFullPath(Directory("./Src/SummitReports.sln"));
var frameworkVersion = Pluck(System.IO.File.ReadAllText(projectFile), @"<TargetFramework>", @"</TargetFramework>");
var buildDir = binDir + dirSep + configuration + dirSep + frameworkVersion + dirSep + "Publish" + dirSep;
var publishDir = binDir + dirSep + configuration + dirSep + frameworkVersion + dirSep + "Publish" + dirSep;
var tempPath = System.IO.Path.GetTempPath();
var deployPath = thisDir + "artifacts" + dirSep;
var SummitNuGetPath = @"\\SIM-SVR03\Software\NuGet\";
public int MAJOR = 0; public int MINOR = 1; public int REVISION = 2; public int BUILD = 3; //Version Segments

var VersionInfoText = System.IO.File.ReadAllText(thisDir + "Src/VersionInfo.cs");
var AssemblyFileVersionAttribute = Pluck(VersionInfoText, "AssemblyFileVersionAttribute(\"", "\")]");
var CurrentAssemblyVersionAttribute = Pluck(VersionInfoText, "System.Reflection.AssemblyVersionAttribute(\"", "\")]");

var AssemblyVersionAttribute = CurrentAssemblyVersionAttribute;
var CurrentNugetVersion = VersionStringParts(AssemblyVersionAttribute, MAJOR, MINOR, REVISION);
var NugetVersion = CurrentNugetVersion;
if (IncrementMinorVersion) {	
	AssemblyVersionAttribute = VersionStringIncrement(CurrentAssemblyVersionAttribute, REVISION);
	NugetVersion = VersionStringParts(AssemblyVersionAttribute, MAJOR, MINOR, REVISION);
	AssemblyFileVersionAttribute = NugetVersion + ".*";
}

Information("	  AssemblyVersionAttribute: {0}... Next: {1}", CurrentAssemblyVersionAttribute, AssemblyVersionAttribute);
Information("	      CoreVersionAttribute: {0}... Next: {1}", GetVersionInProjectFile(projectFile), AssemblyVersionAttribute);
Information("        		 Nuget version: {0}... Next: {1}", CurrentNugetVersion, NugetVersion);
Information("AssemblyFileVersionAttribute : {0}", AssemblyFileVersionAttribute);

//////////////////////////////////////////////////////////////////////
// PREPARATION
//////////////////////////////////////////////////////////////////////

// Define directories.
Console.WriteLine(string.Format("binDir={0}", binDir));
Console.WriteLine(string.Format("thisDir={0}", thisDir));
Console.WriteLine(string.Format("buildDir={0}", buildDir));

var buildSettings = new DotNetCoreBuildSettings
 {
	 Framework = frameworkVersion
	 , Configuration = configuration
	 //, OutputDirectory = deployPath
 };

//////////////////////////////////////////////////////////////////////
// TASKS
//////////////////////////////////////////////////////////////////////

Task("Clean")
    .Does(() =>
{
    CleanDirectory(buildDir);
});

Task("Restore-NuGet-Packages")
    .IsDependentOn("Clean")
    .Does(() =>
{

	var settings = new NuGetRestoreSettings()
	{
		// VSTS has old version of Nuget.exe and Automapper restore fails because of that
		ToolPath = "./nuget/nuget.exe",
		Verbosity = NuGetVerbosity.Detailed,
	};
	//NuGetRestore("./Src/SummitReports.sln", settings);
	DotNetCoreRestore( srcDir );
});

Task("Build")
    .IsDependentOn("Restore-NuGet-Packages")
    .Does(() =>
{
    if(IsRunningOnWindows())
    {
	   DotNetCoreBuild(srcDir, buildSettings);
    }
    else
    {
      // Use XBuild
      XBuild("./Src/SummitReports.sln", settings =>
        settings.SetConfiguration(configuration));
    }
});

Task("NuGet-Pack")
    .IsDependentOn("Build")
    .Does(() =>
{
    ClearDirectory(thisDir + "artifacts/");
    var nuGetPackSettings   = new NuGetPackSettings {
		BasePath 				= thisDir,
        Id                      = @"SummitReports",
        Version                 = NugetVersion,
        Title                   = @"SummitReports - Easy Summit Reports",
        Authors                 = new[] {"Ricardo Vega Jr."},
        Owners                  = new[] {"Ricardo Vega Jr."},
        Description             = @"A library that allows easy processing of Summit Reports.",
        Summary                 = @"A library that allows easy processing of Summit Reports.",
        ProjectUrl              = new Uri(@"https://github.com/ajoyce-sim/SummitReports"),
        //IconUrl                 = new Uri(""),
        LicenseUrl              = new Uri(@"https://github.com/ajoyce-sim/SummitReports/blob/master/LICENSE"),
        Copyright               = @"Summit Investments 2019",
        ReleaseNotes            = new [] {"This report libary can generated UW Relationship Cash Flow"},
        Tags                    = new [] {"Summit", "Reports"},
        RequireLicenseAcceptance= false,
        Symbols                 = false,
        NoPackageAnalysis       = false,
        OutputDirectory         = thisDir + "artifacts/",
		Properties = new Dictionary<string, string>
		{
			{ @"Configuration", @"Release" }
		},
		Files = new[] {
			new NuSpecContent { Source = thisDir + @"Src/SummitReports.Objects/bin/Release/netstandard2.0/SummitReports.Objects.deps.json", Target = "lib/netstandard2.0" },
			new NuSpecContent { Source = thisDir + @"Src/SummitReports.Objects/bin/Release/netstandard2.0/SummitReports.Objects.dll", Target = "lib/netstandard2.0" },
			new NuSpecContent { Source = thisDir + @"Src/SummitReports.Objects/bin/Release/netstandard2.0/SummitReports.Objects.pdb", Target = "lib/netstandard2.0" },

			new NuSpecContent { Source = thisDir + @"Src/SummitReports.Objects/bin/Release/net461/SummitReports.Objects.dll", Target = "lib/net461" },
			new NuSpecContent { Source = thisDir + @"Src/SummitReports.Objects/bin/Release/net461/SummitReports.Objects.pdb", Target = "lib/net461" },
		},
		ArgumentCustomization = args => args.Append("")		
    };
            	
    NuGetPack(thisDir + "NuGet/SummitReports.nuspec", nuGetPackSettings);
});

Task("SetVersion")
.IsDependentOn("NuGet-Pack")
.Does(() => {
	var VersionData = string.Format(@"using System.Reflection;
[assembly: System.Reflection.AssemblyFileVersionAttribute(""{0}"")]
[assembly: System.Reflection.AssemblyVersionAttribute(""{1}"")]
", AssemblyFileVersionAttribute, AssemblyVersionAttribute);
		System.IO.File.WriteAllText(thisDir + "Src/VersionInfo.cs", VersionData);
		UpdateVersionInProjectFile(projectFile, AssemblyVersionAttribute);
});

Task("Finish")
  .IsDependentOn("SetVersion")
  .Does(() =>
{
	if (!DirectoryExists(SummitNuGetPath)) {
		Information(string.Format("Nuget Directory {0} does not exist, skipping nuget copy", SummitNuGetPath ));
	} else {
		Information(string.Format("Copying {0} to {1}", publishDir, deployPath ));
		CopyDirectory(publishDir, deployPath);
		var sourceFolder = thisDir + @"artifacts/";
		foreach (string sourceFile in System.IO.Directory.GetFiles(sourceFolder, @"*", SearchOption.AllDirectories))
		{
			string destinationFile = sourceFile.Replace(sourceFolder, SummitNuGetPath);
			Information(string.Format("Copying {0} to {1}", sourceFile, destinationFile ));
			System.IO.File.Copy(sourceFile, destinationFile, true);
		}
	}
	Information("Build Script has completed");
});

Task("Default")
    .IsDependentOn("Finish");

//////////////////////////////////////////////////////////////////////
// EXECUTION
//////////////////////////////////////////////////////////////////////

RunTarget(target);

public bool ClearDirectory(string PathToClear) {
	System.IO.DirectoryInfo di = new DirectoryInfo(PathToClear);

	foreach (FileInfo file in di.GetFiles())
	{
		file.Delete(); 
	}
	foreach (DirectoryInfo dir in di.GetDirectories())
	{
		dir.Delete(true); 
	}
	Information(string.Format("Directory {0} Cleared", PathToClear ));
	return true;
}
public bool IsValidVersionString(string versionString) {
	var v = versionString.Replace(".", "");
	int test;
    return int.TryParse(v, out test);
}

//segmentToIncrement can equal Major, Minor, Revision, Build in format Major.Minor.Revision.Build
public string VersionStringIncrement(string versionString, int segmentToIncrement) {
	var vArr = versionString.Split('.');
	var valAsStr = vArr[segmentToIncrement];
	int valAsInt = 0;
    int.TryParse(valAsStr, out valAsInt);	
	vArr[segmentToIncrement] = (valAsInt + 1).ToString();
	return String.Join(".", vArr);
}

//versionSegments can equal Major, Minor, Revision, Build in format Major.Minor.Revision.Build
public string VersionStringParts(string versionString, params int[] versionSegments) {
	var vArr = versionString.Split('.');
	string newVersion = "";
	foreach ( var versionSegment in versionSegments ) {
		newVersion += (newVersion.Length>0 ? "." : "") + vArr[versionSegment].ToString();
	}
	return newVersion;
}

public string Pluck(string str, string leftString, string rightString)
{
	try
	{
//Information(@"{0}: str={1}, leftString='{2}', rightString='{3}'", "Pluck", str, leftString, rightString);
		var lpos = str.IndexOf(leftString);
//Information(@"{0}: lpos={1}", "Pluck", lpos);
		if (lpos > 0)
		{
			var rpos = str.IndexOf(rightString, lpos);
//Information(@"{0}: rpos={1}", "Pluck", rpos);
			if ((rpos > 0) && (rpos > lpos))
			{
				return str.Substring(lpos + leftString.Length, (rpos - lpos) - leftString.Length);
			}
		}
	}
	catch (Exception)
	{
		return "";
	}
	return "";
}

public string GetVersionInAssembly(string VarName, string AssemblyFileName) {
	var version = "";
	try
	{
		var manifestText = System.IO.File.ReadAllText(AssemblyFileName);
	//Information("{0}-{1}: Text={2}", AssemblyFileName, VarName, manifestText);

		var potentialVersion = Pluck(manifestText, VarName + "(\"", "\")]");
	//Information("{0}-{1}: potentialVersion={2}", AssemblyFileName, VarName, potentialVersion);
		if (IsValidVersionString(potentialVersion)) {
			version = potentialVersion;
		} else {
			throw new Exception(String.Format("{0}-{1}: Version {2} format is invalid.", AssemblyFileName, VarName, potentialVersion));   
		}
	}
	catch (Exception ex)
	{
		throw ex;
	}
	return version;
}


public bool WriteVersionInAssembly(string VarName, string AssemblyFileName, string VersionSet) {
//Information("WriteVersionInAssembly->{0}-{1}: \n\nVersionSet={2}\n\n", AssemblyFileName, VarName, VersionSet);
	try
	{
		var manifestText = System.IO.File.ReadAllText(AssemblyFileName);
		var potentialVersion = Pluck(manifestText, VarName + "(\"", "\")]");
		if (IsValidVersionString(VersionSet)) {

			var fromVersionString = VarName + "(\"" + potentialVersion + "\")]";
			var toVersionString = VarName + "(\"" + VersionSet + "\")]";
			//Information("WriteVersionInAssembly->{0}-{1}: \n\nfromVersionString={2}, toVersionString={3}\n\n", AssemblyFileName, VarName, fromVersionString, toVersionString);
			
			manifestText = manifestText.Replace(fromVersionString, toVersionString);

			//Information("WriteVersionInAssembly->{0}-{1}: Version={2}, Next Version={3}", AssemblyFileName, VarName, potentialVersion, VersionSet);

			potentialVersion = VersionStringIncrement(potentialVersion, BUILD);
			potentialVersion = VersionStringParts(potentialVersion, MAJOR, MINOR, BUILD) + ".0000";
			System.IO.File.WriteAllText(AssemblyFileName, manifestText);
			//Information("{0}: Version={1}", AssemblyFileName, manifestText);
		} else {
			throw new Exception(String.Format("{0}: Version {1} format is invalid.", AssemblyFileName, potentialVersion));   
		}
		return true;
	}
	catch (Exception ex)
	{
		throw ex;
	}
}

public List<string> DirSearch(string sDir, string FileToSearchFor) 
{
	List<string> lstFilesFound = new List<string>();
	try
	{
		foreach (string d in System.IO.Directory.GetDirectories(sDir)) 
		{
			foreach (string f in System.IO.Directory.GetFiles(d, FileToSearchFor)) 
			{
				lstFilesFound.Add(f);
			}
			lstFilesFound.AddRange(DirSearch(d, FileToSearchFor));
		}
		return lstFilesFound;
	}
	catch (System.Exception) 
	{
		throw;
	}
}

public IEnumerable<FileInfo> TraverseDirectory(string rootPath, Func<FileInfo, bool> Pattern)
{
	var directoryStack = new Stack<DirectoryInfo>();
	directoryStack.Push(new DirectoryInfo(rootPath));
	while (directoryStack.Count > 0)
	{
		var dir = directoryStack.Pop();
		try
		{
			foreach (var i in dir.GetDirectories())
				directoryStack.Push(i);
		}
		catch (UnauthorizedAccessException) {
			continue; // We don't have access to this directory, so skip it
		}
		foreach (var f in dir.GetFiles().Where(Pattern)) // "Pattern" is a function
			yield return f;
	}
}

public string GetVersionInProjectFile(string projectFileName) {
	var _VersionInfoText = System.IO.File.ReadAllText(projectFileName);
	var _AssemblyFileVersionAttribute = Pluck(_VersionInfoText, "<Version>", "</Version>");
	return _AssemblyFileVersionAttribute;
}

public bool UpdateVersionInProjectFile(string projectFileName, string NewVersion)
{
	var _VersionInfoText = System.IO.File.ReadAllText(projectFileName);
	var _AssemblyFileVersionAttribute = Pluck(_VersionInfoText, "<Version>", "</Version>");
	var VersionPattern = "<Version>{0}</Version>";
	var _AssemblyFileVersionAttributeTextOld = string.Format(VersionPattern, _AssemblyFileVersionAttribute);
	var _AssemblyFileVersionAttributeTextNew = string.Format(VersionPattern, NewVersion);
	var newText = _VersionInfoText.Replace(_AssemblyFileVersionAttributeTextOld, _AssemblyFileVersionAttributeTextNew);

	System.IO.File.WriteAllText(projectFileName, newText);	
	return true;
}
  

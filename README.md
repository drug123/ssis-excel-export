# How to copy more than 255 characters into Excel Destination from OLE DB Source in SSIS 2019?
Short answer: there is no way do to exactly that.
# However
You can use Script Task to do that.
# Solution
To save more than 255 characters to Excel file in SSIS, don’t use Excel Destination. It won't work except a few edge cases. Instead, use combination of SQL Script Task + Script Task:

General idea is to store your dataset in package variable and save this dataset using library that could save data in Excel format. I am using MiniExcel assembly that you can get using NuGet. However, there is a caveat: scripting task cannot use NuGet packages directly, instead you have to install assembly into Global Assembly Cache (GAC). To do that, your library of choice must be signed and compatible with .net Framework installed on the machine and compatible with VS 2019. MiniExcel is compatible with .net Framework 4.5 that is installed on my server, so I used exactly that. 
# Preparation
## Download MiniExcel assembly package
Open https://nuget.info/packages/MiniExcel/1.31.3 (my solution is tested with version 1.31.3) and double-click on lib\net45\MiniExcel.dll

This will download the assembly dll file.
## Install assembly into GAC
Save the file and open terminal window from Visual Studio menu: Tools - Command Line - Developer Command Prompt. In command prompt, navigate to the folder where yo have stored MiniExcel.dll:
cd c:\temp\miniexcel\
Now, install the assembly:
gacutil.exe -i miniexcel.dll
You should see "Assembly successfully added to the cache". If any error occurred, this could mean you need to start console with elevated permissions.

# Configuring the package
## Dataset variable
You have to create package variable of type Object to transfer dataset from SQL Task to Script Task. For this, right-click on empty space inside package canvas and select Variables. In Variables window, click on the first button in the toolbar, give a name to the variable, and choose Data Type - Object:

## SQL Script task configuration
Open editor of Execute SQL Task, select your OLE DB Connection and enter your select statement into SQLStatement.
ResultSet property must be set to "Full result set":

Now, go to Result Set pane (select it in the left part of the window), put "0" into Result Name field and choose your Variable created previously:

Close window by pression OK.
## Script Task configuration
Put your Script Task on the canvas, and double-click on it to open task editor, then click on "…" next to ReadOnlyVariables property. Select your dataset variable and close windows by pressing OK:

Open Visual Studio by clicking on Edit Script… button.

In Visual Studio window, go to project properties by right-clicking on project name:

In Project Properties, select .NET Framework 4.5 as target framework:

Save settings and switch to ScriptMain.cs code editor window.
Now, go to Solution Explorer and right-click on References:

In Reference Manager, click on "Browse" button:

Navigate to C:\Windows\Microsoft.NET\assembly\GAC_MSIL\MiniExcel\v4.0_1.31.3.0__e7310002a53eac39 and select MiniExcel.dll then click Add button and then OK button.

In code, expand namespaces region and add two references:
```
using System.Data.OleDb;
using MiniExcelLibs;
```

Add following code into your Main method:

```
public void Main()
{
	// You can use 
	// String outFile = Dts.Variables["User::RoutesOutputFile"].Value.ToString();
	// to pass output file name in RoutesOutputFile user variable (datatype = String)
	String outFile = "c:\temp\export.xlsx";

	// This fills DataTable ds from _ds variable
	OleDbDataAdapter da = new OleDbDataAdapter();
	DataTable ds = new DataTable();
	da.Fill(ds, Dts.Variables["User::_ds"].Value);

	// This static method saves DataTable into using provided filename, in worksheet
	// named "Export Data", overwritting file if it exists already 
	// (last parameter = true)
	MiniExcel.SaveAs(outFile, ds, true, "Export Data", ExcelType.XLSX, null, true);

	Dts.TaskResult = (int)ScriptResults.Success;
}
```

Save, close and run the package.

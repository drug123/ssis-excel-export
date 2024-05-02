# How to copy more than 255 characters into Excel Destination from OLE DB Source in SSIS 2019?
Short answer: there is no way do to exactly that (at least without a pain).
# However
You can use Script Task to do that.
# Solution
To save more than 255 characters to Excel file in SSIS, don’t use Excel Destination. It won't work except a few edge cases. Instead, use combination of SQL Script Task + Script Task:

![image](https://github.com/drug123/ssis-excel-export/assets/2583788/64c7c449-559a-45df-b751-31c2a1d309fd)

The general idea is to store your dataset in a package variable and save this dataset using a library that could save data in Excel format. I am using the MiniExcel assembly that you can get using NuGet. However, there is a caveat: scripting tasks cannot use NuGet packages directly, instead you have to install the assembly into Global Assembly Cache (GAC). To do that, your library of choice must be signed and compatible with .net Framework installed on the machine and compatible with VS 2019. MiniExcel is compatible with .net Framework 4.5 which is installed on my server, so I used exactly that. 

# Preparation
## Download MiniExcel assembly package
Open https://nuget.info/packages/MiniExcel/1.31.3 (my solution is tested with version 1.31.3) and double-click on `lib\net45\MiniExcel.dll`

![image](https://github.com/drug123/ssis-excel-export/assets/2583788/708146a8-efd3-4f55-895b-268bfa9bfe55)

This will download the assembly dll file.
## Install assembly into GAC
Save the file and open the terminal window from the Visual Studio menu: Tools - Command Line - Developer Command Prompt. In the command prompt, navigate to the folder where yo have stored `MiniExcel.dll`:
```
cd c:\temp\miniexcel\
```
Now, install the assembly:
```
gacutil.exe -i miniexcel.dll
```
You should see "Assembly successfully added to the cache". If any error occurs, this could mean you need to start the console with elevated permissions.

# Configuring the package
## Dataset variable
You have to create a package variable of type Object to transfer the dataset from SQL Task to Script Task. For this, right-click on empty space inside the package canvas and select Variables. In the Variables window, click on the first button in the toolbar, give a name to the variable, and choose Data Type - Object:

![image](https://github.com/drug123/ssis-excel-export/assets/2583788/061cc48a-e273-4d73-9791-58c54d1974f4)

## SQL Script task configuration
Open the editor of Execute SQL Task, select your OLE DB Connection and enter your select statement into SQLStatement.
ResultSet property must be set to "Full result set":

![image](https://github.com/drug123/ssis-excel-export/assets/2583788/d65e1a8e-4bd9-442d-94a6-af2782cac8c8)

Now, go to the Result Set pane (select it in the left part of the window), put "0" into the Result Name field and choose your Variable created previously:

![image](https://github.com/drug123/ssis-excel-export/assets/2583788/e7fc0873-d67c-4f67-aee6-0b3404bdb5df)

Close window by pressing OK.
## Script Task configuration
Put your Script Task on the canvas, and double-click on it to open the task editor, then click on "…" next to ReadOnlyVariables property. Select your dataset variable and close windows by pressing OK:

![image](https://github.com/drug123/ssis-excel-export/assets/2583788/8739ee03-583a-4b90-ab1b-c5b85b86ca54)

Open Visual Studio by clicking on the Edit Script… button.

![image](https://github.com/drug123/ssis-excel-export/assets/2583788/cff9991f-fa9c-4e0c-88fb-7417eeee9f66)

In the Visual Studio window, go to project properties by right-clicking on the project name:

![image](https://github.com/drug123/ssis-excel-export/assets/2583788/de3c0177-4495-494b-881d-eb37a531c4bf)

In Project Properties, select .NET Framework 4.5 as the target framework:

![image](https://github.com/drug123/ssis-excel-export/assets/2583788/91337e54-d89c-4cff-b10b-a331c75ac49f)

Save settings and switch to ScriptMain.cs code editor window.
Now, go to Solution Explorer and right-click on References:

![image](https://github.com/drug123/ssis-excel-export/assets/2583788/c85f0b71-f4a0-4647-889c-830ceb53801f)

In Reference Manager, click on the "Browse" button:

![image](https://github.com/drug123/ssis-excel-export/assets/2583788/4bb8b524-a3ce-4d16-8262-2a5fdba74713)

Navigate to `C:\Windows\Microsoft.NET\assembly\GAC_MSIL\MiniExcel\v4.0_1.31.3.0__e7310002a53eac39` and select `MiniExcel.dll` then click the Add button and then the OK button.

In the code, expand the namespaces region and add two references:
```C#
using System.Data.OleDb;
using MiniExcelLibs;
```

Add the following code to your Main method:

```C#
public void Main()
{
	// You can use 
	// String outFile = Dts.Variables["User::yourOutputFile"].Value.ToString();
	// to pass output file name in yourOutputFile user variable (datatype = String)
	String outFile = "c:\temp\export.xlsx";

	// This fills DataTable ds from _ds variable
	OleDbDataAdapter da = new OleDbDataAdapter();
	DataTable ds = new DataTable();
	da.Fill(ds, Dts.Variables["User::_ds"].Value);

	// This static method saves DataTable using the provided filename, in worksheet
	// named "Export Data", overwriting the file if it exists already 
	// (last parameter = true)
	MiniExcel.SaveAs(outFile, ds, true, "Export Data", ExcelType.XLSX, null, true);

	Dts.TaskResult = (int)ScriptResults.Success;
}
```

Save, close and run the package.

NB!: This method probably wouldn't work for a huge amount of data as it requires memory to be allocated to store the dataset in the variable. Watch your counters then.

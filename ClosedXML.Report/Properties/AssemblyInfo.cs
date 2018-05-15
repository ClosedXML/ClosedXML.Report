using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("ClosedXML.Report")]
[assembly: AssemblyDescription(@"ClosedXML.Report is a tool for report generation and data analysis in .NET applications through the use of Microsoft Excel. ClosedXML.Report is a .NET-library for report generation Microsoft Excel without requiring Excel to be installed on the machine that's running the code.")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("ClosedXML.Report")]
[assembly: AssemblyCopyright("MIT")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// Tests accessing
#if STRONGNAME
[assembly: InternalsVisibleTo("ClosedXML.Report.Tests, PublicKey=00240000048000009400000006020000002400005253413100040000010001002b89cad160592ac2770c3651edc6b07038b8d00544858ab632e767ba4052afbdfb4c7fc18deb924b3aea7ecceca8b394e2b14cb9d98e780c8e6159dcb62da75d99f1e1e583ad078ac4836275c58c0cbb786ed317a9a6ac7fa70a9355a88bc41b9f8cf5d5547020c6f02eca4ecf01e1db5d84c6e6788b103064cdd78e14a3c0b2")]
#else
[assembly: InternalsVisibleTo("ClosedXML.Report.Tests")]
#endif

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("fa0a7362-87d7-466e-a6ec-2e1ad0171880")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Build and Revision Numbers 
// by using the '*' as shown below:
// [assembly: AssemblyVersion("1.0.*")]
[assembly: AssemblyVersion("0.1.0.0")]
[assembly: AssemblyFileVersion("0.1.0.0")]

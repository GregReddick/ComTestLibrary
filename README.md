# Calling .NET 5 (.NET Core) from a COM Host Like Excel
This project is a minimal example of using .NET 5 to create a library that is 
able to be early bound and called by any Windows COM (Component Object Model) 
host, particularly from Microsoft Office (Access, Excel, Outlook, PowerPoint, 
Visio, Word). Note that .NET 5 is a shortened name for .NET Core Version 5. 
While this was relatively easy in the earlier .NET Framework projects there 
is missing functionality in .NET Core that makes this difficult.

*Recommendation: Unless you have a good reason for building the library that 
needs to be called from a COM host in .NET 5 (or any version of .NET Core), I 
recommend instead using .NET Framework 4.8 as your target environment. An 
example of a good reason is a library that will **also** be called by a .NET 
Core application. There are complexities, particularly in the .idl file and 
registration, that you have to do by hand in this technique that are solved 
automatically by the .NET Framework. I expect in future versions of .NET Core 
(.NET 5.1?, .NET 6?), it will create the type library and register it. 
When that happens, you can make the transition without what is shown here.*

The description that needs to go into the .idl file has to describe the 
interface that you define in C# in the terms that they will translate to 
under the covers and be consumed by COM. This translation is beyond the scope 
of what I can cover here. The definitive work on the subject is Adam Nathan's 
[.NET and COM: The Complete Interoperability Guide] 
(https://www.amazon.com/gp/product/B003AYZB7U), which although old (2002) and 
still in print only in a Kindle edition, covers this (and many other things) 
in excruciating detail (1579 pages). If you ever need to do communication 
between COM and .NET, it is worth having a copy of this book. Very little of 
the topics covered in the book have changed in almost 20 years, including the 
transition from .NET Framework to .NET Core.

# The Issue
There has been a change in what .NET Core does compared to what the .NET 
Framework did. In the .NET Framework, it built a DLL and embedded a type 
library into it describing the interface. It then used a single process, 
mscoree, that translated the information in the DLL into terms that COM could 
process. .NET Core uses a different technique. Instead of having a single 
process do the translation, it provides a wrapper library, so that if the 
base library is x.dll, the wrapper is x.comhost.dll. The comhost dll does the 
work of translating COM calls into .NET functionality. The main problem is 
that the comhost DLL does not embed a type library that allows early binding 
to the library. So we need to create that type library ourselves, and create 
the correct registry entries so that COM can connect to the library and make 
the proper calls. The technical details of the .NET Core wrapping of 
libraries can be found 
[here](https://github.com/dotnet/runtime/blob/1d9e50cb4735df46d3de0cee5791e97295eaf588/docs/design/features/COM-activation.md).

# Limitation
One COM host cannot start two different versions of .NET Core in the same 
process. Because the COM library runs in-process with the COM host, this 
means that if you use this technique on two libraries that use different 
versions of the .NET Core, the second library called may not work. There
currently is no way to start .NET Core COM libraries out-of-process.

# Getting Started
You will need the MIDL compiler. This is installed in the C++ packages, not 
the C# packages in the Visual Studio installer. You will also need the CL 
compiler to perform C++ preprocessing on the .IDL file. I have not determined 
exactly what is the minimal install, but installing all the stuff needed to 
build C++ programs will give you what you need.

# Bitness (32 bit versus 64 bit)
Another issue that you will run into is a problem of bitness. A 32 bit COM 
host cannot call a 64 bit .NET library. And the same is true the other way 
around--a 64 bit COM host cannot call a 32 bit library. We have three 
different pieces that must be all the same bitness: the COM host, the comhost 
file, and the .NET library (e.g. Excel.exe, x.comhost.dll, x.dll). There is a 
fourth piece that also has bitness: the version of Windows. Almost all copies 
of Windows installed now are 64 bit, but that hasn't always been the case. 64 
bit programs and libraries cannot run on 32 bit Windows, but 32 bit programs 
and libraries will run on 64 bit Windows.

The COM host will be something like Excel. By default, the Office installer 
installs 32 bit editions of Microsoft Office. However, you can also install 
64 bit editions of Office by running the 64 bit installer. The .NET library 
that we build can be compiled as 32 bit (x86), 64 bit (x64) or have it marked 
as AnyCPU. What AnyCPU means is that it can be just-in-time (JIT) compiled to 
32 bit or 64 bit as needed. There is a compiler flag that marks that if it 
can be compiled to either one, which one should it compile to. If you don't 
pass in the flag, the JIT compiler will default to 32 bit, but compile to 64 
bit if needed. The flag can also be set to default to 64 bit. However, the 
comhost wrapper file does not have that flexibility. It is static compiled to 
either 32 bit or 64 bit. We need to get this compiled to the same bitness as 
the COM host that will calling it. This is also performed by a compiler flag. 
All of these compiler flags are set in the Visual Studio project file.

The current project has two different builds, x86 (32 bit) and x64 (64 bit). 
These can be "Batch build" to build both of them at the same time. Some of 
the complexity of handling both 32 and 64 bit will soon be fixed: 
https://github.com/dotnet/runtime/issues/32493. This will allow a AnyCPU 
build to have both a 32 bit and 64 bit comhost file. I will simplify the 
build process once this enhancement is made and released.

# GUIDs
You will need to generate three GUIDs (Globally Unique Identifiers, also 
called UUID, Universally Unique Identifiers) to get your own projects to 
work. These are essentially very large random numbers. **DO NOT USE the GUIDs 
I have in the sample project in your own projects.** The purpose of a GUID is 
to be unique across the entire universe, so that no two projects will ever 
conflict. You will need one for library and one each for the class and the 
interface. You can generate GUIDs in Visual Studio by selecting **Tools > 
Create GUID** from the menu. These GUIDs will need to be placed into the 
appropriate places in the AssemblyInfo.cs and definitions.idl files. Once you 
create a GUID for a project, never change it, as each time you do you will be 
creating additional entries in the Windows registry, bloating the registry 
and making Windows just a little slower.

# The Sample Project
The sample project provides just one method that calculates the area of a 
circle with a given radius. It also accepts a string, although it really 
doesn't do anything with it. The point is to show passing stuff in and 
getting stuff back. The name of the library is ComTestLibrary, the name of 
the class is ComTest, the name of the interface is IComTest, and the name of 
the method is ComTestMethod.

# The AssemblyInfo.cs File
The AssemblyInfo file defines two important assembly-wide attributes, 
ComVisible and Guid.
```
[assembly: ComVisible(false)]
[assembly: Guid(ComTestLibrary.AssemblyInfo.LibraryGuid)]
```
The first attribute tell the compiler not to make anything visible to COM 
unless they are specifically marked with ComVisible(true). This keeps 
anything in the library that we don't explicity mark from poluting the 
registry. The second assigns the GUID for the library. Also included is a 
little helper class in the AssemblyInfo.cs file. This class defines the three 
GUIDs as constants that are then used throughout the project. It also has a 
method that makes it easy to retrieve assembly attributes. This will be used 
when registering the type library. This class was first published in my book, 
[The Reddick C# Style 
Guide](https://www.amazon.com/Reddick-Style-Guide-practices-writing/dp/06925317
 42).
```
namespace ComTestLibrary
{
	/// <summary>Gives information about the assembly. Change the GUIDs in your own project.</summary>
	internal static class AssemblyInfo
	{
		/// <summary>Unique identifier for the class.</summary>
		internal const string ClassGuid = "71AD0B2F-E5D0-4272-A4FD-18F707D5E0D6";

		/// <summary>Unique identifier for the interface.</summary>
		internal const string InterfaceGuid = "1B31B683-F0AA-4E71-8F50-F2D2E5E9E210";

		/// <summary>Unique identifier for the library.</summary>
		internal const string LibraryGuid = "1B31B683-F0AA-4E71-8F50-F2D2E5E9E210";

		/// <summary>Gets an assembly attribute.</summary>
		/// <typeparam name="T">Assembly attribute type.</typeparam>
		/// <returns>The assembly attribute of type T.</returns>
		internal static T Attribute<T>()
			where T : Attribute
		{
			return typeof(AssemblyInfo).Assembly.GetCustomAttribute<T>();
		}
	}
}
```
# The Interface File
To start, you will need an interface file. It will describe the interface
to the library. The important part of this interface is that it must be
decorated with a few attributes.
```
[ComVisible(true)]
[Guid(AssemblyInfo.InterfaceGuid)]
public interface IComTest
```
The ComVisible attribute tells the compiler that it should make this interface
visible to COM. The Guid attribute assigns the Guid for the interface. The
ComInterfaceType default is InterfaceIsDual, which means that it can be early
or late bound.

# The Class File
The class file provides the actual functionality. It implements the 
interface. It must have its own GUID.
```
[ComVisible(true)]
[Guid(AssemblyInfo.ClassGuid)]
public class ComTest : IComTest
```
The class file needs to provide two additional methods besides implementing 
the interface, DLLRegisterServer and DLLUnregisterServer. These will be 
discussed later.

# Setting Up the Project File
At the top of the project file are these settings:
```
<PropertyGroup>
	<CLVersion>14.28.29828</CLVersion>
	<DriveLetter>N:</DriveLetter>
	<KitsVersion>10.0.19041.0</KitsVersion>
	<VSVersion>2019\Preview</VSVersion>
</PropertyGroup>
```
You will need to change these settings to the values appropriate for your 
environment. The CLVersion is set to the version of the CL compiler on your 
computer. The KitsVersion is set to the version of the Windows Kits on your 
computer. The DriveLetter is set to a drive letter that is not in use on your 
computer. The VSVersion is set to the version of Visual Studio on your 
computer such as 2017\Community. These get expanded into settings later in 
the Project File.

# The Project File
The project file needs these XML settings:
```
<PropertyGroup>
	<AssemblyName>$(MSBuildProjectName)$(Platform)</AssemblyName>
	<Description>Com Test Library</Description>
	<EnableComHosting>true</EnableComHosting>
	<EnableDefaultItems>false</EnableDefaultItems>
	<GenerateDocumentationFile>True</GenerateDocumentationFile>
	<NETCoreSdkRuntimeIdentifier>win-$(Platform)</NETCoreSdkRuntimeIdentifier>
	<Platforms>x64;x86</Platforms>
	<TargetFramework>net5.0-windows</TargetFramework>
	<PathKitsBin>c:\Program Files (x86)\Windows Kits\10\bin\$(KitsVersion)</PathKitsBin>
	<PathKitsInclude>C:\Program Files (x86)\Windows Kits\10\include\$(KitsVersion)</PathKitsInclude>
	<MidlOptions>
		/cpp_cmd "$(DriveLetter)\cl.exe"
		/dlldata nul /h nul /iid nul /proxy nul
		/I "$(PathKitsInclude)\um\64"
		/I "$(PathKitsInclude)\um"
		/I "$(PathKitsInclude)\shared"
		/out "bin\$(Platform)\$(Configuration)\$(TargetFramework)"
		/tlb "$(MSBuildProjectName)$(Platform).comhost.tlb"
		definitions.idl
	</MidlOptions>
</PropertyGroup>
```
# Build and Test
The AssemblyName setting will change the name of the final assembly to have 
the bitness appended to the name. The Description setting is used as the 
description of the type library inside the registry. The EnableComHosting 
setting tells the compiler that it needs to build the 
ComTestLibraryx86.comhost.dll (or ComTestLibraryx64.comhost.dll). The 
MidlOptionsThe are passed to the midl compiler. NETCoreSdkRuntimeIdentifier 
win-$(Platform) setting informs the compiler that when it builds the 
comhost.dll, it should match the bitness of the library. The Platforms 
setting tells Visual Studio it should support building both bitness values. 
The TargetFramework setting tells the compiler that it should use .NET 5 
built for windows. Since none of this runs on any operating system other than 
Windows, it won't complain about features that don't work elsewhere.

The MidlOptions are passed to the Midl compiler. A somewhat big catch 22 is 
that the /cpp_cmd path to the cl.exe file is installed in a directory that 
has spaces in the path, but the midl compiler cannot find it if it has 
spaces. This problem has existed forever, but Microsoft seems to refuse to 
fix it. The workaround used here is to use the ancient subst to provide a 
drive letter for this directory, then use the drive letter in the midl 
options. The drive letter must not be in use.

# The IDL File
The next task is to generate a type library. Because .NET Core doesn't do 
this task, we need to do it ourselves. This is done by creating an IDL file. 
The structure of an IDL file can be found 
[here](https://docs.microsoft.com/en-us/windows/win32/midl/midl-start-page).
```
[
	uuid(49d618d5-a2d1-4fb6-ad3a-f8dc5ca25e02),
	version(1.0),
	helpstring("ComTestLibrary")
]
library ComTestLibrary
{
	importlib("STDOLE2.TLB");

	[
		odl,
		uuid(1B31B683-F0AA-4E71-8F50-F2D2E5E9E210),
		dual,
		oleautomation,
		nonextensible,
		helpstring("ComTestLibrary"),
		object,
	]
	interface IComTest : IDispatch
	{
		[
			id(1),
			helpstring("ComTestMethod")
		]
		HRESULT ComTestMethod(
			[in] double radius,
			[in] BSTR comment,
			[out, retval] double* ReturnVal);
	};

	[
		uuid(71AD0B2F-E5D0-4272-A4FD-18F707D5E0D6),
		helpstring("ComTest")
	]
	coclass ComTest
	{
		[default] interface IComTest;
	};
}
```
The first uuid must match the GUID provided for library in the AssemblyInfo 
file. The second uuid must match the one for the interface. The third 
must match the one for the class.

The critical part of the file is the description of the method. This must be 
the COM equivalent for the method in the class. For example, strings in C# 
must be listed as the BSTR data type for COM. The method must return a 
HRESULT. A parameter is the actual return value. Each method must have a 
unique id number.

This file is compiled by the MIDL compiler. The MIDL compiler calls the CL 
compiler to preprocess the file. There are a number of libraries and header 
files that it needs to access to get the compiling done.

To process the file, rather than attempting to pass everything to the MIDL 
command line, we pass in a response file midloptions.txt. The response file 
is created by the project file and saved to the target directory.

There is a new version of the MIDL language (version 3) that is much 
simplified, but it seems that it will not produce .tlb files, so is useless 
for the task needed here.

# More Project File
Going back to the project file, here is another section of the file:
```
<Target AfterTargets="PostBuildEvent" Name="PostBuild">
	<WriteLinesToFile File="$(TargetDir)midloptions.txt" Overwrite="true" Lines="$(MidlOptions)" />
	<Exec command="&quot;$(PathKitsBin)\x64\midl.exe&quot; @$(TargetDir)midloptions.txt&quot;" />
	<Exec command="regsvr32 /s &quot;$(TargetDir)$(TargetName).comhost.dll&quot;" />
</Target>
<Target AfterTargets="BeforeClean" BeforeTargets="CoreClean" Name="RegClean">
	<Exec IgnoreExitCode="true" Command="regsvr32 /s /u &quot;$(TargetDir)$(TargetName).comhost.dll&quot;" />
</Target>
```
After a successful build of the project, this executes the MIDL compiler and 
compiles the type library. Then it registers the comhost file. The clean part 
unregisters the comhost file before cleaning. **For the registration to work, 
Visual Studio must be executed as an administrator, otherwise it will not 
have the privilege necessary to create the registry entries.**

# DLLRegisterServer and DLLUnregisterServer
The comhost file has a minimal amount of stuff to register the file. However 
it does not create all of the registry entries needed to use the library from 
a COM host. There needs to be a number of other entries to register the type 
library. To create those entries, create two additional methods in the class 
file. When the comhost file is registered, if the DLLRegisterServer method 
exists, it will get called. When it is unregistered, it will call the 
DLLUnregisterServer method. These methods need to be decorated with attributes
 to identify them. The methods look like this:
```
[ComRegisterFunction]
public static void DllRegisterServer(Type t)
{
	using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(@"TypeLib\{" + AssemblyInfo.ClassGuid + @"}"))
	{
		Version version = typeof(AssemblyInfo).Assembly.GetName().Version;
		using (RegistryKey keyVersion = key.CreateSubKey(string.Format("{0}.{1}", version.Major, version.Minor)))
		{
			keyVersion.SetValue(string.Empty, AssemblyInfo.Attribute<AssemblyDescriptionAttribute>().Description, RegistryValueKind.String);
			using (RegistryKey keyWin32 = keyVersion.CreateSubKey(@"0\win32"))
			{
				keyWin32.SetValue(string.Empty, Path.ChangeExtension(Assembly.GetExecutingAssembly().Location, ".comhost.tlb"), RegistryValueKind.String);
			}

			using (RegistryKey keyFlags = keyVersion.CreateSubKey(@"FLAGS"))
			{
				keyFlags.SetValue(string.Empty, "0", RegistryValueKind.String);
			}
		}
	}
}

[ComUnregisterFunction]
public static void DllUnregisterServer(Type t)
{
	Registry.ClassesRoot.DeleteSubKeyTree(@"TypeLib\{" + AssemblyInfo.ClassGuid + @"}", false);
}
```
The AssemblyInfo class is found in the AssemblyInfo.cs file.

# Calling the DLL From Excel
To try calling the DLL from Excel VBA, start Excel and press Alt+F11 to enter 
the Visual Basic Editor. Create a reference by selecting **Tools > References** 
from the menu. Check the checkbox next to ComTestLibrary and press OK. This 
creates an early binding reference to the library. Select **Insert > Module** 
from the menu. Insert the following VBA code:
```
Public Sub EarlyBinding
	'Requires reference to ComTestLibrary being added
	Dim comtest As ComTestLibrary.ComTest
	Dim area As Double
	Set comtest = New ComTestLibrary.ComTest
	area = comtest.ComTestMethod(3, "abcdefghi")
	MsgBox area
End Sub
```
Click in the middle of the sub and press the F5 key to execute it. If all 
went well, it will show you the area of a circle with the radius of 3.

Try it again with late binding, where a reference to the library is not
required. Late binding will not provide intellisense when calling the method.
```
Public Sub LateBinding
	Dim comtest As Object
	Dim area As Double
	Set comtest = CreateObject("ComTestLibrary.ComTest")
	area = comtest.ComTestMethod(3, "abcdefghi")
	MsgBox area
End Sub
```

# Using COM Library in Production

To use this in production, six files should be copied to the install 
directory: ComTestLibraryx32.comhost.dll, ComTestLibraryx32.comhost.tlb, 
ComTestLibraryx32.dll, ComTestLibraryx64.comhost.dll, 
ComTestLibraryx64.comhost.tlb, ComTestLibraryx64.dll After they are copied, 
then regsvr32 needs to be run on the two comhost files to get the directories
in the registry to point to the files.
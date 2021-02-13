# Calling .NET 5 (.NET Core) from COM Host Like Excel
This project is a minimal example of using .NET 5 (or really any recent 
version of .NET Core, .NET 5 is just .NET Core version 5) to create a library 
that is able to be early bound and called by any Windows COM host, 
particularly from Microsoft Office (Access, Excel, Outlook, PowerPoint, 
Visio, Word). While this was relatively easy in the earlier .NET Framework 
projects there is missing functionality in .NET Core that makes this difficult.

Recommendation: Unless you have a good reason for building the library in 
.NET 5 (or any version of .NET Core), I recommend using .NET Framework 4.8 as 
your target environment. An example of a good reason is a library that will 
*also* be called by a .NET core application. There are complexities, 
particularly the .idl file and registration, that you have to do by hand in 
this technique that are solved automatically by the .NET Framework. I expect 
in future versions of .NET Core (.NET 5+, .NET 6?), that they will do 
automatically some of the things I am doing by hand here. When that happens, 
you can make the transition.

The description that needs to go into the .idl file has to describe the 
interface that you define in C# in the terms that they will translate to 
under the covers and be consumed by COM. This translation is beyond the scope 
of what I can cover here. The definitive work on the subject is Adam Nathan's 
[.NET and COM: The Complete Interoperability Guide] 
(https://www.amazon.com/gp/product/B003AYZB7U), which although old (2002) and 
only still in print in a Kindle edition, covers all of these details in 
excruciating detail (1579 pages). If you ever need to do communication 
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
the proper calls.

# Getting Started
You will need the MIDL compiler. This is installed in the C++ packages, not 
the C# packages in the Visual Studio installer. You will also need the CL 
compiler to perform C++ preprocessing on the .IDL file. I have not determined 
exactly what is the minimal install, but installing all the stuff needed to 
build C++ programs will give you what you need.

# Build and Test
References to tools and projects in this example are fragile. I have 
hard-coded strings that reference the tools and project files that you will 
need to update to match your environment. I had this sample project break 
just by updating to the latest version of Visual Studio, which changed the 
path to one of the tools. Inside the response.txt file that is fed to the 
MIDL compiler, there is a hack to get around the fact that the MIDL compiler 
doesn't like spaces in some file or directory names. This hack is to use the 
short names (8.3 FAT names) to access the files. This is also particularly 
fragile as short names are assigned by the operating system and can be 
different on different computers.

Making this work in a team environment will be difficult as a build that 
works on one computer may not work on the next one. I will leave this as a 
task for the reader to get this working where it can just be downloaded to an 
arbitrary computer and have it build correctly and also not break when there 
are updated tools. Adding things to the computer's path or creating 
enivornment variables may be a solution. Another technique that may work is 
copying tools to a shared directory that has no spaces in the name.

# Bitness (32 bit versus 64 bit)
Another issue that you will run into is a problem of bitness. A 32 bit COM 
host cannot call a 64 bit .NET library. And the same is true the other way 
around--a 64 bit COM host cannot call a 32 bit library. We have three 
different pieces that must be all the same bitness: the COM host, the comhost 
file, and the .NET library. There is a fourth piece that also has bitness: 
the version of Windows. Almost all copies of Windows installed now are 64 
bit, but that hasn't always been the case. 64 bit programs and libraries 
cannot run on 32 bit Windows, but 32 bit programs and libraries will run on 
64 bit Windows.

The COM host will be something like Excel. By default, the Office installer 
installs 32 bit editions of Microsoft Office. However, you can also install 
64 bit editions of Office by running the 64 bit installer, which happens in 
some organizations. The .NET library that we build can be compiled as 32 bit 
(x86), 64 bit (x64) or have it marked as AnyCPU. What AnyCPU means is that it 
can be just-in-time (JIT) compiled to 32 bit or 64 bit as needed. There is a 
compiler flag that marks that if it can be compiled to either one, which one 
should it compile to. If you don't pass in the flag, the JIT compiler will 
default to 32 bit, but compile to 64 bit if that is the only version that 
will run (which is true on some Windows servers). The flag can also be set to 
default to 64 bit. However, the comhost wrapper file does not have that 
flexibility. It is static compiled to either 32 bit or 64 bit. We need to get 
this compiled to the same bitness as the COM host that will calling it. This 
is also performed by a compiler flag. All of these compiler flags are set in 
the Visual Studio project file.

# GUIDs
You will need to generate three GUIDs (Globally Unique Identifiers, also 
called UUID, Universally Unique Identifiers) to get your own projects to 
work. These are essentially very large random numbers. **DO NOT USE** the 
GUIDs I have in the sample project in your own projects. The purpose of a 
GUID is to be unique across the entire universe, so that no two projects will 
ever conflict. You will need one for library and one each for the class and 
the interface. You can generate GUIDs in Visual Studio by selecting "Tools > 
Create GUID" from the menu. These GUIDs will need to be placed into the 
appropriate places in the AssemblyInfo.cs, IComTest.cs, ComTest.cs, and 
definitions.idl files. Once you create a GUID for a project, never change it, 
as each time you do you will be creating additional entries in the Windows 
registry, bloating the registry and making Windows just a little slower.

# The Sample Project
The sample project provides just one method, that calculates the area of a 
circle with a given radius. It also accepts a string, although it really 
doesn't do anything with it. The point is to show passing stuff in and 
getting stuff back. The name of the library is ComTestLibrary, the name of 
the class is ComTest, the name of the interface is IComTest, and the name of 
the method is ComTestMethod.

# The AssemblyInfo.cs File
The AssemblyInfo file defines two important attributes.

```
[assembly: ComVisible(false)]
[assembly: Guid("1B31B683-F0AA-4E71-8F50-F2D2E5E9E210")]
```

The first attribute tell the compiler not to make anything visible to COM
unless they are specifically marked with ComVisible(true). This keeps anything
in the library that we don't explicity mark from poluting the registry.
The second assigns the GUID for the
library.

# The Interface File
To start, you will need an interface file. It will describe the interface
to the library. The important part of this interface is that it must be
decorated with a few attributes.

```
[ComVisible(true)]
[Guid("1B31B683-F0AA-4E71-8F50-F2D2E5E9E210")]
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
[Guid("71AD0B2F-E5D0-4272-A4FD-18F707D5E0D6")]
public class ComTest : IComTest
```

The class file needs to provide two additional methods besides implementing 
the interface, DLLRegisterServer and DLLUnregisterServer. These will be 
discussed later.

# The Project File

The project file needs these XML settings:

```
<PropertyGroup>
	<EnableComHosting>true</EnableComHosting>
	<NETCoreSdkRuntimeIdentifier>win-x86</NETCoreSdkRuntimeIdentifier>
	<PlatformTarget>AnyCPU</PlatformTarget>
	<TargetFramework>net5.0-windows7.0</TargetFramework>
</PropertyGroup>
```
The EnableComHosting setting tells the compiler that it needs to build the 
ComTestLibrary.comhost.dll. The NETCoreSdkRuntimeIdentifier win-x86 setting 
informs the compiler that when it builds the comhost.dll, it should be a 32 
bit file (use win-x64 if you wanted to build 64 bit). The PlatformTarget 
setting tells the compiler that it should make the actual DLL AnyCPU (use x64 
for 64 bit). The TargetFramework setting tells the compiler that it should 
use .NET 5 built for windows 7.0 or later. Since none of this runs on any 
operating system other than Windows, it won't complain about features that 
don't work elsewhere.

# The IDL File
The next task is to generate a type library. Because .NET Core doesn't do 
this task, we need to do it ourselves. This is done by creating an IDL file. 
The structure of an IDL file can be found [here] 
(https://docs.microsoft.com/en-us/windows/win32/midl/midl-start-page).

```
import "unknwn.idl";

[
	object,
	uuid(1B31B683-F0AA-4E71-8F50-F2D2E5E9E210),
	dual,
	nonextensible,
	helpstring("ComTestLibrary"),
	pointer_default(unique),
	oleautomation
]
interface IComTest : IDispatch
{
	[id(1), helpstring("ComTestMetho")]
	HRESULT ComTestMethod(
		[in] double radius,
		[in] BSTR comment,
		[out, retval] double* ReturnVal);
};

[
	uuid(49d618d5-a2d1-4fb6-ad3a-f8dc5ca25e02),
	helpstring("ComTestLibrary")
]
library ComTestLibrary
{
	importlib("STDOLE2.TLB");

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
file. The second uuid must match the one in the interface file. The third 
must match the one in the class file.

This file is compiled by the MIDL compiler. The MIDL compiler calls the CL 
compiler to preprocess the file. There are a number of libraries and header 
files that it needs to access to get the compiling done.

To process the file, rather than attempting to pass everything to the MIDL 
command line, we pass in a response file. The response file looks like this:

```
/I "C:\Program Files (x86)\Windows Kits\10\include\10.0.19041.0\um\64;C:\Program Files (x86)\Windows Kits\10\include\10.0.19041.0\um;c:\Program Files (x86)\Windows Kits\10\Include\10.0.19041.0\shared"
/cpp_cmd C:\Progra~2\Micros~2\2019\Preview\VC\Tools\MSVC\14.28.29828\bin\Hostx64\x64\cl.exe
/cpp_opt "/E /I C:\Progra~2\wi3cf2~1\10\include\10.0.19041.0\um"
/win32
"D:\src.cs\ComTestLibrary\definitions.idl"
/tlb "D:\src.cs\ComTestLibrary\bin\Debug\net5.0-windows7.0\ComTestLibrary.comhost.tlb"
```

This file will need to be tweaked to match your file system. As mentioned 
above the MIDL compiler gets picky about having spaces in file names, so the 
Progra~2\Micros~2 needs to match the short names for your C:\Program Files 
(x86)\Microsoft Visual Studio directory. the C:\Progra~2\wi3cf2~1 needs to 
match the short names of the C:\Program Files (x86)\Windows Kits directory.
The D:\src.cs will need to match where the sources for the project are
located.

# More Project File
Going back to the project file, here is another section of the file:
```
<Target AfterTargets="PostBuildEvent" Name="PostBuild">
	<Exec command="cmd /S /c &quot;&quot;C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64\midl.exe&quot; @$(ProjectDir)response.txt&quot;" />
	<Exec command="regsvr32 /s &quot;$(TargetDir)$(TargetName).comhost.dll&quot;" />
</Target>
<Target BeforeTargets="Clean" Name="RegClean">
	<Exec IgnoreExitCode="true" Command="regsvr32 /s /u &quot;$(TargetDir)$(TargetName).comhost.dll&quot;" />
</Target>
```
After a successful build of the project, this executes the MIDL compiler and compiles the type library.
Then it registers the comhost file. The clean part unregisters the comhost file before cleaning.
For the registration to work, Visual Studio must be executed as an administrator, otherwise it will
not have the privilege necessary to create the registry entries.

# DLLRegisterServer and DLLUnregisterServer
The comhost file has a minimal amount of stuff to register the file. However it does not create
all of the registry entries needed to use the library from a COM host. There needs to be a number
of other entries. To create those entries, create two additional methods in the class file. When
the comhost file is registered, if the DLLRegisterServer method exists, it will get called.
When it is unregistered, it will call the DLLUnregisterServer method. These methods need
to be decorated with attributes to identify them. The methods look like this:
```
	[ComRegisterFunction]
	public static void DllRegisterServer(Type t)
	{
		Registry.CurrentUser.CreateSubKey(@"SOFTWARE\ComTestLibrary");
	}

	[ComUnregisterFunction]
	public static void DllUnregisterServer(Type t)
	{
		Registry.CurrentUser.DeleteSubKeyTree(@"SOFTWARE\ComTestLibrary");
	}
```

# Calling the DLL From Excel
To try calling the DLL from Excel VBA, start Excel and press Alt+F11 to enter
the Visual Basic Editor. Create a reference by selecting Tools > References from the
menu. Check the checkbox next to ComTestLibrary and press OK. This creates an early
binding reference to the library. Select Insert > Module from the menu. Insert the following
VBA code:
```
Public Sub TryIt
	Dim comtest As ComTestLibrary.ComTest
	Dim area As Double
	Set comtest = New ComTestLibrary.ComTest
	area = comtest.ComTestMethod(3, "abcdefghi")
	MsgBox area
End Sub
```
Click in the middle of the sub and press the F5 key to execute it. If all went well, it will show you the area
of a circle with the radius of 3.
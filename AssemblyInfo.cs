//--------------------------------------------------------------------------------------------------
// <copyright file="AssemblyInfo.cs" company="Xoc Software">
// Copyright © 2021 Xoc Software
// </copyright>
// <summary>Implements the assembly information class</summary>
//--------------------------------------------------------------------------------------------------
using System;
using System.Reflection;
using System.Runtime.InteropServices;

[assembly: ComVisible(false)]
[assembly: Guid(ComTestLibrary.AssemblyInfo.LibraryGuid)]

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
		internal const string LibraryGuid = "47A20781-26AD-465F-BDA9-AC59CEA74B69";

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
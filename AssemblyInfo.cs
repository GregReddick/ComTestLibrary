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
[assembly: Guid("1B31B683-F0AA-4E71-8F50-F2D2E5E9E210")]

namespace ComTestLibrary
{
	/// <summary>Gives information about the assembly.</summary>
	internal static class AssemblyInfo
	{
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
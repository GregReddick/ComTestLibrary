// --------------------------------------------------------------------------------------------------
// <copyright file="ComTest.cs" company="Xoc Software">
//     Copyright © 2021 Xoc Software
// </copyright>
// --------------------------------------------------------------------------------------------------

namespace ComTestLibrary
{
	using System;
	using System.Runtime.InteropServices;

	using Microsoft.Win32;

	/// <summary>(COM visible) a com test class.</summary>
	/// <seealso cref="T:ComTestLibrary.IComTest"/>
	[ComVisible(true)]
	[Guid("71AD0B2F-E5D0-4272-A4FD-18F707D5E0D6")]
	public class ComTest : IComTest
	{
		/// <summary>DLL register server.</summary>
		/// <param name="t">A Type to process.</param>
		[ComRegisterFunction]
		public static void DllRegisterServer(Type t)
		{
			Registry.CurrentUser.CreateSubKey(@"SOFTWARE\ComTestLibrary");
		}

		/// <summary>DLL unregister server.</summary>
		/// <param name="t">A Type to process.</param>
		[ComUnregisterFunction]
		public static void DllUnregisterServer(Type t)
		{
			Registry.CurrentUser.DeleteSubKeyTree(@"SOFTWARE\ComTestLibrary");
		}

		/// <summary>Com test method.</summary>
		/// <param name="radius">The radius of a circle.</param>
		/// <param name="comment">The a random pointless comment.</param>
		/// <returns>The area of a circle for the given radius.</returns>
		/// <seealso cref="M:ComTestLibrary.IComTest.ComTestMethod(double,string)"/>
		public double ComTestMethod(double radius, string comment)
		{
			// Do some pointless work. This just shows that you can pass in a VBA string.
			comment = comment.Replace("abc", "def");
			return Math.PI * (radius * radius);
		}
	}
}
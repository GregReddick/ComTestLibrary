// --------------------------------------------------------------------------------------------------
// <copyright file="IComTest.cs" company="Xoc Software">
// Copyright © 2021 Xoc Software
// </copyright>
// --------------------------------------------------------------------------------------------------

namespace ComTestLibrary
{
	using System.Runtime.InteropServices;

	/// <summary>Interface for com test.</summary>
	[ComVisible(true)]
	[Guid(AssemblyInfo.InterfaceGuid)]
	public interface IComTest
	{
		/// <summary>Com test method.</summary>
		/// <param name="radius">The radius of a circle.</param>
		/// <param name="comment">The a random pointless comment.</param>
		/// <returns>The area of a circle for the given radius.</returns>
		double ComTestMethod(double radius, string comment);
	}
}
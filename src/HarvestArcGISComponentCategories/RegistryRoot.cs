namespace HarvestArcGISComponentCategories
{
	/// <summary>
	/// Defines values for the Root column of the Registry, RemoveRegistry, and RegLocator tables.
	/// </summary>
	//-------------------------------------------------------------------------------------------------
	// <copyright file="RegistryRoot.cs" company="Microsoft">
	//    Copyright (c) Microsoft Corporation.  All rights reserved.
	//    
	//    The use and distribution terms for this software are covered by the
	//    Common Public License 1.0 (http://opensource.org/licenses/cpl1.0.php)
	//    which can be found in the file CPL.TXT at the root of this distribution.
	//    By using this software in any fashion, you are agreeing to be bound by
	//    the terms of this license.
	//    
	//    You must not remove this notice, or any other, from this software.
	// </copyright>
	// 
	// <summary>
	// Harvest WiX authoring from the registry.
	// </summary>
	//-------------------------------------------------------------------------------------------------
	public enum RegistryRoot
	{
		/// <summary>HKEY_CURRENT_USER for a per-user installation,
		/// or HKEY_LOCAL_MACHINE for a per-machine installation.</summary>
		UserOrMachine = -1,

		/// <summary>HKEY_CLASSES_ROOT</summary>
		ClassesRoot = 0,

		/// <summary>HKEY_CURRENT_USER</summary>
		CurrentUser = 1,

		/// <summary>HKEY_LOCAL_MACHINE</summary>
		LocalMachine = 2,

		/// <summary>HKEY_USERS</summary>
		Users = 3,
	}
}
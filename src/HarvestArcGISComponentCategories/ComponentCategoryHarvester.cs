using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace HarvestArcGISComponentCategories
{
	/// <summary>
	/// Harvest Component Categories that get written to registry. 
	/// Originally adapted from RegistryHarvester in WIX. Original copyright notice:
	//-------------------------------------------------------------------------------------------------
	// <copyright file="RegistryHarvester.cs" company="Microsoft">
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
	/// </summary>
	public sealed class ComponentCategoryHarvester : IDisposable
	{
		private const string HKCRPathInHKLM = @"Software\Classes";
		private readonly string _remappedPath;
		private static readonly int _majorOsVersion = Environment.OSVersion.Version.Major;
		private readonly RegistryKey _regKeyToOverride = Registry.LocalMachine;
		private readonly UIntPtr _regRootToOverride = NativeMethods.HkeyLocalMachine;

		private static readonly IList<Guid> _excludeCategories =
			new List<Guid>
				{
					new Guid(
						"{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}")
				};

		/// <summary>
		/// Instantiate a new ComponentCategoryHarvester.
		/// </summary>
		/// <param name="remap">Set to true to remap the entire registry to a private location for this process.</param>
		public ComponentCategoryHarvester(bool remap)
		{
			// Detect OS major version and set the hive to use when
			// redirecting registry writes. We want to redirect registry
			// writes to HKCU on Windows Vista and higher to avoid UAC
			// problems, and to HKLM on downlevel OS's.
			if (_majorOsVersion >= 6)
			{
				_regKeyToOverride = Registry.CurrentUser;
				_regRootToOverride = NativeMethods.HkeyCurrentUser;
			}

			// create a path in the registry for redirected keys which is process-specific
			if (remap)
			{
				_remappedPath = string.Concat(@"SOFTWARE\WiX\heat\",
				                              Process.GetCurrentProcess().Id.ToString(
				                              	CultureInfo.InvariantCulture));

				// remove the previous remapped key if it exists
				RemoveRemappedKey();

				// remap the registry roots supported by MSI
				// note - order is important here - the hive being used to redirect
				// to must be overridden last to avoid creating the other override
				// hives in the wrong location in the registry. For example, if HKLM is
				// the redirect destination, overriding it first will cause other hives
				// to be overridden under HKLM\Software\WiX\heat\HKLM\Software\WiX\HKCR
				// instead of under HKLM\Software\WiX\heat\HKCR
				if (_majorOsVersion < 6)
				{
					RemapRegistryKey(NativeMethods.HkeyClassesRoot,
					                 string.Concat(_remappedPath,
					                               @"\\HKEY_CLASSES_ROOT"));
					RemapRegistryKey(NativeMethods.HkeyCurrentUser,
					                 string.Concat(_remappedPath,
					                               @"\\HKEY_CURRENT_USER"));
					RemapRegistryKey(NativeMethods.HkeyUsers,
					                 string.Concat(_remappedPath, @"\\HKEY_USERS"));
					RemapRegistryKey(NativeMethods.HkeyLocalMachine,
					                 string.Concat(_remappedPath,
					                               @"\\HKEY_LOCAL_MACHINE"));
				}
				else
				{
					RemapRegistryKey(NativeMethods.HkeyClassesRoot,
					                 string.Concat(_remappedPath,
					                               @"\\HKEY_CLASSES_ROOT"));
					RemapRegistryKey(NativeMethods.HkeyLocalMachine,
					                 string.Concat(_remappedPath,
					                               @"\\HKEY_LOCAL_MACHINE"));
					RemapRegistryKey(NativeMethods.HkeyUsers,
					                 string.Concat(_remappedPath, @"\\HKEY_USERS"));
					RemapRegistryKey(NativeMethods.HkeyCurrentUser,
					                 string.Concat(_remappedPath,
					                               @"\\HKEY_CURRENT_USER"));

					// Typelib registration on Windows Vista requires that the key 
					// HKLM\Software\Classes exist, so add it to the remapped root
					Registry.LocalMachine.CreateSubKey(HKCRPathInHKLM);
				}
			}
		}

		/// <summary>
		/// Close the ComponentCategoryHarvester and remove any remapped registry keys.
		/// </summary>
		public void Close()
		{
			// note - order is important here - we must quit overriding the hive
			// being used to redirect first
			if (_majorOsVersion < 6)
			{
				NativeMethods.OverrideRegistryKey(NativeMethods.HkeyLocalMachine,
				                                  IntPtr.Zero);
				NativeMethods.OverrideRegistryKey(NativeMethods.HkeyClassesRoot,
				                                  IntPtr.Zero);
				NativeMethods.OverrideRegistryKey(NativeMethods.HkeyCurrentUser,
				                                  IntPtr.Zero);
				NativeMethods.OverrideRegistryKey(NativeMethods.HkeyUsers, IntPtr.Zero);
			}
			else
			{
				NativeMethods.OverrideRegistryKey(NativeMethods.HkeyCurrentUser,
				                                  IntPtr.Zero);
				NativeMethods.OverrideRegistryKey(NativeMethods.HkeyClassesRoot,
				                                  IntPtr.Zero);
				NativeMethods.OverrideRegistryKey(NativeMethods.HkeyLocalMachine,
				                                  IntPtr.Zero);
				NativeMethods.OverrideRegistryKey(NativeMethods.HkeyUsers, IntPtr.Zero);
			}

			RemoveRemappedKey();
		}

		/// <summary>
		/// Dispose the ComponentCategoryHarvester.
		/// </summary>
		public void Dispose()
		{
			Close();
		}

		/// <summary>
		/// Harvest all registry roots needed (only ClassesRoot for 
		/// component categories).
		/// </summary>
		/// <returns>The registry keys and values in the registry.</returns>
		public IDictionary<Guid, IList<Guid>> HarvestRegistry()
		{
			var registryValues = new ArrayList();

			HarvestRegistryKey(Registry.ClassesRoot, registryValues);
			//HarvestRegistryKey(Registry.CurrentUser, registryValues);
			//HarvestRegistryKey(Registry.LocalMachine, registryValues);
			//HarvestRegistryKey(Registry.Users, registryValues);

			IDictionary<Guid, IList<Guid>> componentCategories =
				new Dictionary<Guid, IList<Guid>>();

			foreach (RegistryValue registryValue in registryValues)
			{
				if (!registryValue.KeyName.Contains("Implemented Categories"))
				{
					continue;
				}

				Guid category = GetComponentCategory(registryValue.KeyName);

				if (_excludeCategories.Contains(category))
				{
					continue;
				}

				Guid classId = GetClassId(registryValue.KeyName);

				if (!componentCategories.ContainsKey(category))
				{
					componentCategories.Add(category, new List<Guid>());
				}

				componentCategories[category].Add(classId);
			}

			return componentCategories;
			//return (object[]) registryValues.ToArray(typeof (object));
		}

		private static Guid GetComponentCategory(string keyName)
		{
			string[] parts = GetAllPathParts(keyName);

			string guidString = parts[parts.Length - 1];

			return new Guid(guidString);
		}

		private static Guid GetClassId(string keyName)
		{
			string[] parts = GetAllPathParts(keyName);

			string guidString = parts[parts.Length - 3];

			return new Guid(guidString);
		}

		/// <summary>
		/// Gets the parts of a registry key's path.
		/// </summary>
		/// <param name="path">The registry key path.</param>
		/// <returns>The root and key parts of the registry key path.</returns>
		private static string[] GetPathParts(string path)
		{
			return path.Split(@"\".ToCharArray(), 2);
		}

		private static string[] GetAllPathParts(string path)
		{
			return path.Split(@"\".ToCharArray());
		}

		/// <summary>
		/// Harvest a registry key.
		/// </summary>
		/// <param name="registryKey">The registry key to harvest.</param>
		/// <param name="registryValues">The collected registry values.</param>
		private static void HarvestRegistryKey(RegistryKey registryKey, ArrayList registryValues)
		{
			// harvest the sub-keys
			foreach (string subKeyName in registryKey.GetSubKeyNames())
			{
				using (RegistryKey subKey = registryKey.OpenSubKey(subKeyName))
				{
					HarvestRegistryKey(subKey, registryValues);
				}
			}

			string[] parts = GetPathParts(registryKey.Name);

			RegistryRoot root;
			switch (parts[0])
			{
				case "HKEY_CLASSES_ROOT":
					root = RegistryRoot.ClassesRoot;
					break;
				case "HKEY_CURRENT_USER":
					root = RegistryRoot.CurrentUser;
					break;
				case "HKEY_LOCAL_MACHINE":
					// HKLM\Software\Classes is equivalent to HKCR
					if (1 < parts.Length &&
					    parts[1].StartsWith(HKCRPathInHKLM,
					                        StringComparison.OrdinalIgnoreCase))
					{
						root = RegistryRoot.ClassesRoot;
						parts[1] = parts[1].Remove(0, HKCRPathInHKLM.Length);

						if (0 < parts[1].Length)
						{
							parts[1] = parts[1].TrimStart('\\');
						}

						if (string.IsNullOrEmpty(parts[1]))
						{
							parts = new[] {parts[0]};
						}
					}
					else
					{
						root = RegistryRoot.LocalMachine;
					}
					break;
				case "HKEY_USERS":
					root = RegistryRoot.Users;
					break;
				default:
					// TODO: put a better exception here
					throw new Exception();
			}

			// harvest the values
			foreach (string valueName in registryKey.GetValueNames())
			{
				// this is where the actual information is havested:
				var registryValue =
					new RegistryValue(root, valueName, registryKey.Name,
					                  registryKey.GetValue(valueName));

				registryValues.Add(registryValue);
			}

			// If there were no subkeys and no values, we still need an element for this empty registry key.
			// But specifically avoid SOFTWARE\Classes because it shouldn't be harvested as an empty key.
			if (parts.Length > 1 && registryKey.SubKeyCount == 0 &&
			    registryKey.ValueCount == 0 &&
			    !string.Equals(parts[1], HKCRPathInHKLM,
			                   StringComparison.OrdinalIgnoreCase))
			{
				var emptyRegistryValue = new RegistryValue(root, null,
				                                           registryKey.Name, null);

				registryValues.Add(emptyRegistryValue);
			}
		}

		/// <summary>
		/// Remap a registry key to an alternative location.
		/// </summary>
		/// <param name="registryKey">The registry key to remap.</param>
		/// <param name="remappedPath">The path to remap the registry key to under HKLM.</param>
		private void RemapRegistryKey(UIntPtr registryKey, string remappedPath)
		{
			IntPtr remappedKey = IntPtr.Zero;

			try
			{
				remappedKey = NativeMethods.OpenRegistryKey(_regRootToOverride,
				                                            remappedPath);

				NativeMethods.OverrideRegistryKey(registryKey, remappedKey);
			}
			finally
			{
				if (IntPtr.Zero != remappedKey)
				{
					NativeMethods.CloseRegistryKey(remappedKey);
				}
			}
		}

		/// <summary>
		/// Remove the remapped registry key.
		/// </summary>
		private void RemoveRemappedKey()
		{
			try
			{
				_regKeyToOverride.DeleteSubKeyTree(_remappedPath);
			}
			catch (ArgumentException)
			{
				// ignore the error where the key does not exist
			}
		}

		/// <summary>
		/// The native methods for re-mapping registry keys.
		/// </summary>
		private static class NativeMethods
		{
			internal static readonly UIntPtr HkeyClassesRoot = (UIntPtr) 0x80000000;
			internal static readonly UIntPtr HkeyCurrentUser = (UIntPtr) 0x80000001;
			internal static readonly UIntPtr HkeyLocalMachine = (UIntPtr) 0x80000002;
			internal static readonly UIntPtr HkeyUsers = (UIntPtr) 0x80000003;

			private const uint GenericRead = 0x80000000;
			private const uint GenericWrite = 0x40000000;
			private const uint GenericExecute = 0x20000000;
			private const uint GenericAll = 0x10000000;
			private const uint StandardRightsAll = 0x001F0000;

			/// <summary>
			/// Opens a registry key.
			/// </summary>
			/// <param name="key">Base key to open.</param>
			/// <param name="path">Path to subkey to open.</param>
			/// <returns>Handle to new key.</returns>
			internal static IntPtr OpenRegistryKey(UIntPtr key, string path)
			{
				IntPtr newKey;
				uint disposition;

				const uint sam = StandardRightsAll | GenericRead | GenericWrite | GenericExecute |
				                 GenericAll;

				if (0 !=
				    RegCreateKeyEx(key, path, 0, null, 0, sam, 0, out newKey,
				                   out disposition))
				{
					throw new Exception();
				}

				return newKey;
			}

			/// <summary>
			/// Closes a previously open registry key.
			/// </summary>
			/// <param name="key">Handle to key to close.</param>
			internal static void CloseRegistryKey(IntPtr key)
			{
				if (0 != RegCloseKey(key))
				{
					throw new Exception();
				}
			}

			/// <summary>
			/// Override a registry key.
			/// </summary>
			/// <param name="key">Handle of the key to override.</param>
			/// <param name="newKey">Handle to override key.</param>
			internal static void OverrideRegistryKey(UIntPtr key, IntPtr newKey)
			{
				if (0 != RegOverridePredefKey(key, newKey))
				{
					throw new Exception();
				}
			}

			/// <summary>
			/// Interop to RegCreateKeyW.
			/// </summary>
			/// <param name="key">Handle to base key.</param>
			/// <param name="subkey">Subkey to create.</param>
			/// <param name="reserved">Always 0</param>
			/// <param name="className">Just pass null.</param>
			/// <param name="options">Just pass 0.</param>
			/// <param name="desiredSam">Rights to registry key.</param>
			/// <param name="securityAttributes">Just pass null.</param>
			/// <param name="openedKey">Opened key.</param>
			/// <param name="disposition">Whether key was opened or created.</param>
			/// <returns>Handle to registry key.</returns>
			[DllImport("advapi32.dll", EntryPoint = "RegCreateKeyExW",
				CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
			private static extern int RegCreateKeyEx(UIntPtr key, string subkey,
			                                         uint reserved, string className,
			                                         uint options, uint desiredSam,
			                                         uint securityAttributes,
			                                         out IntPtr openedKey,
			                                         out uint disposition);

			/// <summary>
			/// Interop to RegCloseKey.
			/// </summary>
			/// <param name="key">Handle to key to close.</param>
			/// <returns>0 if success.</returns>
			[DllImport("advapi32.dll", EntryPoint = "RegCloseKey",
				CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
			private static extern int RegCloseKey(IntPtr key);

			/// <summary>
			/// Interop to RegOverridePredefKey.
			/// </summary>
			/// <param name="key">Handle to key to override.</param>
			/// <param name="newKey">Handle to override key.</param>
			/// <returns>0 if success.</returns>
			[DllImport("advapi32.dll", EntryPoint = "RegOverridePredefKey",
				CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
			private static extern int RegOverridePredefKey(UIntPtr key, IntPtr newKey);
		}
	}
}
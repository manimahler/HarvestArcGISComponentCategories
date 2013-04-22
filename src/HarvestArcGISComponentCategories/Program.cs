using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml;
using Ionic.Zip;

namespace HarvestArcGISComponentCategories
{
	public class Program
	{
		private const string _zipFileName = "Config.xml";

		private static void Main(string[] args)
		{
			string inputAssemblyFileName = null;

			try
			{
				string outputFolder;

				if (args.Length == 0)
				{
					Console.WriteLine("ERROR: Incorrect number of arguments.");
					Console.WriteLine();
					Console.WriteLine(
						"Usage: HarvestArcGISCategories.exe <input assembly> {output folder}");

					return;
				}
				if (args.Length == 1)
				{
					inputAssemblyFileName = args[0];
					outputFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
				}
				else
				{
					inputAssemblyFileName = args[0];
					outputFolder = args[1];
				}

				Console.WriteLine(string.Format("Harvesting Categories for {0}. Output folder: {1}",
				                                inputAssemblyFileName, outputFolder));

				if (!PathsValid(inputAssemblyFileName, outputFolder))
				{
					return;
				}

				Console.WriteLine("Execution Path: " + outputFolder);

				HarvestResults componentCategories = HarvestRegistryValues(inputAssemblyFileName);

				if (componentCategories == null)
				{
					return;
				}

				string folderName = GetFolderName(componentCategories.AssemblyName,
				                                  componentCategories.AssemblyGuid);

				string dirPath = string.Format("{0}/{1}", outputFolder, folderName);

				if (File.Exists(dirPath))
				{
					File.Delete(dirPath);
				}

				if (!Directory.Exists(dirPath))
				{
					Directory.CreateDirectory(dirPath);
				}

				string zipPath = string.Format("{0}/{1}/{2}", outputFolder, folderName, _zipFileName);

				CreateXml(componentCategories, zipPath);

				string zipFolder = string.Format("{0}/{1}", outputFolder, folderName);
				string zipSaveName = string.Format("{0}.zip", zipFolder);

				using (var zip = new ZipFile())
				{
					zip.AddDirectory(zipFolder);
					zip.Save(zipSaveName);
				}

				File.Delete(zipPath);

				Directory.Delete(zipFolder);

				File.Move(zipSaveName, zipFolder);
			}
			catch (Exception ex)
			{
				// TODO: find how to communicate with msbuild that it failed and to print the error in red
				Console.WriteLine(string.Format("Error Harvesting Categories in {0}: {1}",
				                                inputAssemblyFileName, ex.Message));
				Console.WriteLine(ex.ToString());

				// NOTE: throw results in the crash-dialog to come up... but at least msbuild realises that there was an error
				throw;
			}
		}

		private static bool PathsValid(string inputAssemblyFileName, string outputFolder)
		{
			if (! File.Exists(inputAssemblyFileName))
			{
				//TODO Error handling

				Console.WriteLine(string.Format("File {0} does not exist", inputAssemblyFileName));

				return false;
			}

			if (outputFolder != null && ! Directory.Exists(outputFolder))
			{
				//TODO Error handling

				Console.WriteLine(string.Format("Folder {0} does not exist", outputFolder));

				return false;
			}
			return true;
		}

		private static void CreateXml(HarvestResults componentCategories, string zipPath)
		{
			var xmlWriter = new XmlTextWriter(zipPath, null);
			xmlWriter.Formatting = Formatting.Indented;

			xmlWriter.WriteStartDocument();

			xmlWriter.WriteStartElement("ESRI.Configuration");
			xmlWriter.WriteAttributeString("ver", "1");

			xmlWriter.WriteStartElement("Categories");

			foreach (
				KeyValuePair<Guid, IList<Guid>> componentCategory in
					componentCategories.HarvestedRegistryValues)
			{
				xmlWriter.WriteStartElement("Category");
				xmlWriter.WriteAttributeString("CATID",
				                               "{" + componentCategory.Key.ToString().ToUpper() + "}");

				Console.WriteLine("Component Category: {0}", componentCategory.Key);

				foreach (Guid classId in componentCategory.Value)
				{
					xmlWriter.WriteStartElement("Class");
					xmlWriter.WriteAttributeString("CLSID", classId.ToString("B").ToUpper());
					xmlWriter.WriteEndElement();

					Console.WriteLine("   <Class CLSID=\"" + classId.ToString("B").ToUpper());
				}
				xmlWriter.WriteEndElement();
			}
			xmlWriter.WriteEndElement();
			xmlWriter.WriteEndElement();
			xmlWriter.WriteEndDocument();

			xmlWriter.Close();
		}

		private static HarvestResults HarvestRegistryValues(string path)
		{
			var regSvcs = new RegistrationServices();
			Assembly assembly = Assembly.LoadFrom(path);

			object[] customAttributes = assembly.GetCustomAttributes(typeof (GuidAttribute), false);

			if (customAttributes.Length <= 0)
			{
				Console.WriteLine(string.Format("Assembly {0} does not have a GUID", assembly.FullName));

				return null;
			}
			
			var guidAttribute = (GuidAttribute) customAttributes[0];
			string assemblyGuidString = guidAttribute.Value;

			var assemblyGuid = new Guid(assemblyGuidString);

			Console.WriteLine("Assembly {0}_{1}", assemblyGuid.ToString("B"),
			                  GetAssemblyShortName(assembly));

			try
			{
				// must call this before overriding registry hives to prevent binding failures
				// on exported types during RegisterAssembly
				assembly.GetExportedTypes();
			}
			catch (Exception)
			{
				Console.WriteLine(
					"Error getting types from assembly. Make sure the referenced assemblies exist in the output folder.");
				throw;
			}

			const bool remapRegistration = true;

			using (var componentCategoryHarvester = new ComponentCategoryHarvester(remapRegistration))
			{
				regSvcs.RegisterAssembly(assembly, AssemblyRegistrationFlags.SetCodeBase);

				return new HarvestResults(assemblyGuid,
				                          GetAssemblyShortName(assembly),
				                          componentCategoryHarvester.HarvestRegistry());
			}
		}

		internal static string GetFolderName(string assemblyName, Guid assemblyGuid)
		{
			return string.Format("{0}_{1}.ecfg",
			                     assemblyGuid.ToString("B"),
			                     assemblyName);
		}

		internal static string GetAssemblyShortName(Assembly assembly)
		{
			string location = assembly.Location;
			string assemblyShortName = Path.GetFileNameWithoutExtension(location);
			return assemblyShortName;
		}
	}
}
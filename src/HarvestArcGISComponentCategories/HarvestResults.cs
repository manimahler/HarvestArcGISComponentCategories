using System;
using System.Collections.Generic;

namespace HarvestArcGISComponentCategories
{
	internal class HarvestResults
	{
		private readonly IEnumerable<KeyValuePair<Guid, IList<Guid>>> _harvestedRegistryValues;
		private readonly string _assemblyName;
		private readonly Guid _assemblyGuid;

		public HarvestResults(Guid assemblyGuid, string assemblyName,
		                      IEnumerable<KeyValuePair<Guid, IList<Guid>>>
		                      	harvestedRegistryValues)
		{
			_assemblyGuid = assemblyGuid;
			_assemblyName = assemblyName;
			_harvestedRegistryValues = harvestedRegistryValues;
		}

		public IEnumerable<KeyValuePair<Guid, IList<Guid>>> HarvestedRegistryValues
		{
			get { return _harvestedRegistryValues; }
		}

		public string AssemblyName
		{
			get { return _assemblyName; }
		}

		public Guid AssemblyGuid
		{
			get { return _assemblyGuid; }
		}
	}
}
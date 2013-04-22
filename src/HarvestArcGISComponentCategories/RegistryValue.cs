using System;
using System.Globalization;
using System.Text;

namespace HarvestArcGISComponentCategories
{
	public class RegistryValue
	{
		private string _keyName;
		private string _valueName;
		private RegistryRoot _root;
		private object _value;
		private string _valueAsString;

		public RegistryValue(RegistryRoot root, string valueName, string keyName, object value)
		{
			_root = root;
			_value = value;
			_valueName = valueName;
			_keyName = keyName;
		}

		public string ValueName
		{
			get { return _valueName; }
			set { _valueName = value; }
		}

		public string KeyName
		{
			get { return _keyName; }
			set { _keyName = value; }
		}

		public RegistryRoot Root
		{
			get { return _root; }
			set { _root = value; }
		}

		public object Value
		{
			get { return _value; }
			set { _value = value; }
		}

		public string ValueAsString
		{
			get
			{
				if (string.IsNullOrEmpty(_valueAsString))
				{
					_valueAsString = ConvertToString(_value);
				}
				return _valueAsString;
			}
		}

		public override string ToString()
		{
			return string.Format("Root: {0}, KeyName: {1}, Value Name: {2}, Value: {3}", Root,
			                     KeyName, ValueName, ValueAsString);
		}

		private static string ConvertToString(object value)
		{
			if (value == null)
			{
				return string.Empty;
			}

			string stringValue;
			if (value is byte[]) // binary
			{
				var hexadecimalValue = new StringBuilder();

				// convert the byte array to hexadecimal
				foreach (byte byteValue in (byte[]) value)
				{
					hexadecimalValue.Append(byteValue.ToString("X2",
					                                           CultureInfo.InvariantCulture.NumberFormat));
				}

				stringValue = hexadecimalValue.ToString();
			}
			else if (value is int) // integer
			{
				stringValue = ((int) value).ToString(CultureInfo.InvariantCulture);
			}
			else if (value is string[]) // multi-string
			{
				if (0 == ((string[]) value).Length)
				{
					stringValue = string.Empty;
				}
				else
				{
					stringValue = string.Empty;
					foreach (string multiStringValueContent in (string[]) value)
					{
						// separator?
						stringValue += multiStringValueContent;

					}
				}
			}
			else if (value is string)
				// string, expandable (there is no way to differentiate a string and expandable value in .NET 1.1)
			{
				stringValue = (string) value;
			}
			else
			{
				// TODO: put a better exception here
				throw new ArgumentOutOfRangeException("value", "Unknown registry value type.");
			}
			return stringValue;
		}
	}
}
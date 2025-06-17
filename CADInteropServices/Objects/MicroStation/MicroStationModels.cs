using Bentley.Interop.MicroStationDGN;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CADInteropServices.Objects.MicroStation
{
	public class MicroStationModels : IDisposable
	{
		private ModelReference modelRef;
		private Application microStationApplication;
		public string Name { get; private set; }

		public MicroStationModels(ModelReference modelRef, Application microStationApplication)
		{
			this.modelRef = modelRef;
			this.microStationApplication = microStationApplication;
			this.Name = modelRef.Name;
		}

		public List<MicroStationElements> GetAllElements()
		{
			List<MicroStationElements> elements = new List<MicroStationElements>();

			ElementScanCriteria scanCriteria = new ElementScanCriteriaClass();
			
			

			ElementEnumerator elementEnumerator = modelRef.Scan(scanCriteria);

			while (elementEnumerator.MoveNext())
			{
				Element element = elementEnumerator.Current;
				MicroStationElements msElement = new MicroStationElements(element);
				elements.Add(msElement);
			}

			return elements;
		}

		public List<MicroStationLevels> GetAllLevels()
		{
			List<MicroStationLevels> levels = new List<MicroStationLevels>();

			Levels levelsCollection = modelRef.Levels;

			foreach (Level level in levelsCollection)
			{
				MicroStationLevels msLevel = new MicroStationLevels(level);
				levels.Add(msLevel);
			}

			return levels;
		}

		public void Dispose()
		{
			if (modelRef != null)
			{
				Marshal.ReleaseComObject(modelRef);
				modelRef = null;
			}
		}
	}

}

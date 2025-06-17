using Bentley.Interop.MicroStationDGN;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CADInteropServices.Objects.MicroStation
{
	public class MicroStationLevels
	{
		private Level level;

		public string Name { get; private set; }
		public bool IsOn { get; private set; }
		public int Color { get; private set; }
		public bool IsLocked { get; private set; }

		public MicroStationLevels(Level level)
		{
			this.level = level;
			this.Name = level.Name;
			this.IsOn = level.IsDisplayed;
			this.Color = level.ElementColor;
			this.IsLocked = level.IsLocked;
		}
	}
}

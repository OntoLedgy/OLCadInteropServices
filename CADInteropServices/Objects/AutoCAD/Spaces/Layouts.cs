using Autodesk.AutoCAD.Interop.Common;
using CADInteropServices.Factories;
using CADInteropServices.Objects.AutoCAD;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CADInteropServices.Objects.AutoCAD.Spaces
{
    public class Layouts : IDisposable
    {
        private AcadLayout layout;
        private AcadBlock layoutBlock;
        private AcadBlocks autoCADBlocks;

        public string Name { get; set; }
        public string ConfigurationName { get; set; }
        PaperSizes PaperSize { get; set; }
        public List<AutoCADEntities> Entities { get; private set; }

        public Layouts(
            AcadLayout autoCADLayout,
            AcadBlocks blocks)
        {

            layout = autoCADLayout;

            Name = layout.Name;

            ConfigurationName = layout.ConfigName;

            layoutBlock = layout.Block;

            autoCADBlocks = blocks;

            getPaperSize();

            // Initialize entities list
            Entities = new List<AutoCADEntities>();

            // Process entities in the layout
            GetEntities();

        }

        private void getPaperSize()
        {
            double hieght;
            double width;

            layout.GetPaperSize(
                out hieght,
                out width);

            PaperSize = new PaperSizes(
                hieght,
                width);
        }


        private void GetEntities()
        {
            foreach (AcadEntity entity in layoutBlock)
            {
                AutoCADEntities entityWrapper = EntityFactories.Create(
                    entity,
                    autoCADBlocks);

                if (entityWrapper != null)
                {
                    Entities.Add(entityWrapper);
                }

                // Release the COM entity
                Marshal.ReleaseComObject(entity);
            }
        }



        public void Dispose()
        {
            // Release COM objects
            if (layoutBlock != null)
            {
                Marshal.ReleaseComObject(layoutBlock);
                layoutBlock = null;
            }

            if (layout != null)
            {
                Marshal.ReleaseComObject(layout);
                layout = null;
            }
        }
    }
}

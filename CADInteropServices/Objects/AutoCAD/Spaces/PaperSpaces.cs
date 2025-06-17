using Autodesk.AutoCAD.Interop.Common;
using Autodesk.AutoCAD.Interop;

namespace CADInteropServices.Objects.AutoCAD.Spaces
{
    public class PaperSpaces
    {
        private AcadPaperSpace? paperSpace;
        private List<AcadPaperSpace> paperSpaces;

        public AcadLayout PaperLayout { get; set; }

        public PaperSpaces(AcadPaperSpace autoCADPaperSpace)
        {
            paperSpace = autoCADPaperSpace;

            PaperLayout = paperSpace.Layout;

        }

    }
}

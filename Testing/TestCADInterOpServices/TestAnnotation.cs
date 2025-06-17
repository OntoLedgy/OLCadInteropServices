using CADInteropServices.Orchestrators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestCADInterOpServices
{
    [Apartment(ApartmentState.STA)]
    public class TestAnnotation
    {
        OrchestrateAnnotations cadOrchestrator;

        [SetUp]
        public void Setup()
        {
            cadOrchestrator = new OrchestrateAnnotations(
                "..\\..\\..\\samples\\sampleG.DWG");

        }
        [Test]
        public void TestAddLayers()
        {

            cadOrchestrator.AddLayers();


            Assert.Pass();
        }

        [Test]
        public void TestAnnotatetEntities()
        {

            cadOrchestrator.AnnotateAllEntities();


            Assert.Pass();
        }
    }
}

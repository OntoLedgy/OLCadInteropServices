
using Autodesk.AutoCAD.Interop.Common;
using CADInteropServices.Objects;
using CADInteropServices.Transformers;
using CADInteropServices.Objects.AutoCAD.Shapes;
using CADInteropServices.Objects.AutoCAD.Annotations;
using CADInteropServices.Objects.AutoCAD.Spaces;
using CADInteropServices.Objects.AutoCAD;


namespace CADInteropServices.Factories
{
    public static class EntityFactories
    {
        public static AutoCADEntities Create(
            AcadEntity entity, 
            AcadBlocks blocks = null,
            TransformationMatrix parentTransformationMatrix = null)
        {
            switch (entity)
            {
                case AcadLine acadLine:
                    return new Lines(acadLine);

                case AcadLWPolyline lwPolyline:
                    return new PolyLines(lwPolyline);

                case AcadPolyline polyline:
                    return new PolyLines(polyline);

                case AcadCircle acadCircle:
                    return new Circles(acadCircle);

                case AcadEllipse acadEllipse:                
                    return new Ellipses(acadEllipse);

                case AcadArc acadArc:
                    return new Arcs(acadArc);

                case AcadSolid acadSolid:
                    return new Solids(acadSolid);

                case AcadHatch acadHatch:
                    return new Hatches(acadHatch);
            
                case Acad3DFace face:
                    return new Faces(face);

                case AcadLeader leader:
                    return new Leaders(leader);

                case AcadText text:
                    return new Texts((AcadEntity)text);

                case AcadMText mtext :
                    return new Texts((AcadEntity)mtext);

                case AcadTextStyle textSytle:
                    return new Texts((AcadEntity)textSytle);

                case AcadBlockReference acadBlockReference:
                    return new BlockReferences(acadBlockReference, blocks, parentTransformationMatrix);

                case AcadAttribute attribute:
                     return new AttributeDefinitions((AcadEntity)attribute);

                case AcadAttributeReference attributereference: 
                    return new AttributeDefinitions((AcadEntity)attributereference);

                case AcadPViewport viewport:
                    return new Viewports(viewport, blocks);
                // Add cases for other entity types

                default:
                    // Handle unrecognized entities
                    Console.WriteLine($"Unrecognized entity type: {entity.EntityName} {entity.GetType()}");
                    return new UnrecognizedEntities(entity);
            }
        }
    }
}

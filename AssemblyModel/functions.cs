using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace AssemblyModel
{
    class functions
    {
        // Constraints on WorkAxes x,y,z
        public void AxisContraints(AssemblyDocument oAssyDoc, ComponentOccurrence oC1, ComponentOccurrence oC2, WorkAxis pPartAxis1, WorkAxis pPartAxis2)
        {

            //Inventor.PartComponentDefinition pPartdef1 = null;
            //pPartdef1 = oC1.Definition as PartComponentDefinition;

            //Inventor.PartComponentDefinition pPartdef2 = null;
            //pPartdef2 = oC2.Definition as PartComponentDefinition;

            //Inventor.WorkAxis pPartAxis1 = pPartdef1.WorkAxes[i];
            //Inventor.WorkAxis pPartAxis2 = pPartdef2.WorkAxes[j];
           

            object axisProxy1 = null;
            Inventor.WorkAxisProxy pWorkAxisProxy1 = null;
            oC1.CreateGeometryProxy(pPartAxis1, out axisProxy1);
            pWorkAxisProxy1 = axisProxy1 as WorkAxisProxy;

            object axisProxy2 = null;
            Inventor.WorkAxisProxy pWorkAxisProxy2 = null;
            oC2.CreateGeometryProxy(pPartAxis2, out axisProxy2);
            pWorkAxisProxy2 = axisProxy2 as WorkAxisProxy;

            AssemblyConstraint oConstr;
            AssemblyComponentDefinition oAxisDef;
            oAxisDef = oAssyDoc.ComponentDefinition;
            oConstr = (AssemblyConstraint)oAxisDef.Constraints.AddMateConstraint(pWorkAxisProxy1, pWorkAxisProxy2, 0.0);
          //  Console.WriteLine("it is paasing from here");
        }



        // Constraints on faces of Solid
        public void FaceContraints(AssemblyDocument oAssyDoc, ComponentOccurrence oC1, ComponentOccurrence oC2, int i, int j)
        {

            int k = 1;
            int faceC;
            faceC = oC1.SurfaceBodies[1].Faces.Count;
            //Face[] oCA = new Face[faceC + 1];
            Face oface1, oface2;

            oface1 = oC1.SurfaceBodies[1].Faces[i];
            oface2 = oC2.SurfaceBodies[1].Faces[j];

            AssemblyConstraint oConstr;
            AssemblyComponentDefinition oAxisDef;
            oAxisDef = oAssyDoc.ComponentDefinition;
            oConstr = (AssemblyConstraint)oAxisDef.Constraints.AddMateConstraint(oface1, oface2, 0, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference, null, null);
        }


        public void WorkPlaneC(AssemblyDocument oAssyDoc, ComponentOccurrence oC1,ComponentOccurrence oC2, WorkPlane oworkplane1, WorkPlane oworkplane2)
        {

            int k = 1;
            int faceC;
            //faceC = oC1.SurfaceBodies[1].Faces.Count;
            //Face[] oCA = new Face[faceC + 1];
            //Face oface1, oface2;

            //oface1 = oC1.SurfaceBodies[1].Faces[i];
            //oface2 = oC2.SurfaceBodies[1].Faces[j];

            object workProxy1 = null;
            Inventor.WorkPlaneProxy pWorkPlaneProxy1 = null;
            oC1.CreateGeometryProxy(oworkplane1,out workProxy1);
            //oC1.CreateGeometryProxy()
            pWorkPlaneProxy1 = workProxy1 as WorkPlaneProxy;

            object workProxy2 = null;
            Inventor.WorkPlaneProxy pWorkPlaneProxy2 = null;
            oC2.CreateGeometryProxy(oworkplane2, out workProxy2);
            //oC1.CreateGeometryProxy()
            pWorkPlaneProxy2 = workProxy2 as WorkPlaneProxy;


            AssemblyConstraint oConstr;
            AssemblyComponentDefinition oAxisDef;
            oAxisDef = oAssyDoc.ComponentDefinition;
            //oConstr = (AssemblyConstraint)oAxisDef.Constraints.AddMateConstraint(oworkplane1, oworkplane2, 0, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference, null, null);
            //oConstr = (AssemblyConstraint)oAxisDef.Constraints.AddMateConstraint(oworkplane1, oworkplane2, 0);
            oConstr = (AssemblyConstraint)oAxisDef.Constraints.AddMateConstraint(pWorkPlaneProxy1, pWorkPlaneProxy2, 0.0);
        }

        // Constraints on center point of the solids
        public void CenterContstraints(AssemblyDocument oAssyDoc,ComponentOccurrence oC1, ComponentOccurrence oC2,double offset)
        {
            Inventor.PartComponentDefinition pPartdef1 = null;
            pPartdef1 = oC1.Definition as PartComponentDefinition;
             Inventor.WorkPoint op1;
            op1 = (Inventor.WorkPoint)pPartdef1.WorkPoints[1];
            // object axisProxy1 = null;
            object pop1 = null;
            Inventor.WorkPointProxy pWorkpointProxy1 = null;
            oC1.CreateGeometryProxy(op1, out pop1);
            pWorkpointProxy1 = pop1 as WorkPointProxy;

            Inventor.PartComponentDefinition pPartdef2 = null;
            pPartdef2 = oC2.Definition as PartComponentDefinition;
            Inventor.WorkPoint op2;

            op2 = (Inventor.WorkPoint)pPartdef2.WorkPoints[1];
           
            object pop2 = null;
            Inventor.WorkPointProxy pWorkpointProxy2 = null;
            oC2.CreateGeometryProxy(op2, out pop2);
            pWorkpointProxy2 = pop2 as WorkPointProxy;
        

            AssemblyConstraint oConstr;
            AssemblyComponentDefinition oPointDef;
            oPointDef = oAssyDoc.ComponentDefinition;
            oConstr = (AssemblyConstraint)oPointDef.Constraints.AddMateConstraint(pWorkpointProxy1, pWorkpointProxy2, offset);

            

        }

        public void CenterContstraints2(AssemblyDocument oAssyDoc, ComponentOccurrence oC1, ComponentOccurrence oC2, double offset)
        {
            Inventor.AssemblyComponentDefinition pPartdef1 = null;
            pPartdef1 = oC1.Definition as AssemblyComponentDefinition;
            Inventor.WorkPoint op1;
            op1 = (Inventor.WorkPoint)pPartdef1.WorkPoints[1];
            // object axisProxy1 = null;
            object pop1 = null;
            Inventor.WorkPointProxy pWorkpointProxy1 = null;
            oC1.CreateGeometryProxy(op1, out pop1);
            pWorkpointProxy1 = pop1 as WorkPointProxy;

            Inventor.PartComponentDefinition pPartdef2 = null;
            pPartdef2 = oC2.Definition as PartComponentDefinition;
            Inventor.WorkPoint op2;

            op2 = (Inventor.WorkPoint)pPartdef2.WorkPoints[1];

            object pop2 = null;
            Inventor.WorkPointProxy pWorkpointProxy2 = null;
            oC2.CreateGeometryProxy(op2, out pop2);
            pWorkpointProxy2 = pop2 as WorkPointProxy;


            AssemblyConstraint oConstr;
            AssemblyComponentDefinition oPointDef;
            oPointDef = oAssyDoc.ComponentDefinition;
            oConstr = (AssemblyConstraint)oPointDef.Constraints.AddMateConstraint(pWorkpointProxy1, pWorkpointProxy2, offset);



        }

        public void WriteCustomData(string propName, object propValue, Document invDoc)
        {
            PropertySet propSet = invDoc.PropertySets["Inventor User Defined Properties"];
            Inventor.Property prop;
            try
            {
                prop = propSet[propName];
            }
            catch (Exception e)
            {
                prop = propSet.Add("", propName);
            }
            prop.Value = propValue;


        }

        public object ReadCustomData(string propName, string  f,Inventor.Application InventorApplication)
        {
           // Inventor.Application InventorApplication=default(Inventor.Application);
            PartDocument invDoc;
            //try { 
                invDoc = (PartDocument)InventorApplication.Documents.Open(f, false); 
            //}
            //catch (Exception e)
            // { 
            //    invDoc=(PartDocument)InventorApplication.Documents.
            
            //}
            

            PropertySet propSet = invDoc.PropertySets["Inventor User Defined Properties"];
            Inventor.Property prop;

            //prop = propSet.ItemByPropId("Part Number");

            prop = propSet[propName];

            //invDoc.Close();

            return prop.Value;

        }

        public void extrude( PartComponentDefinition oPartCompDef, Profile oProfile, double length, int j, int i) //Inventor.ObjectCollection Bodies)
        {
            ExtrudeDefinition oextrudedef = default(ExtrudeDefinition);
            ExtrudeFeature oExtrude = default(ExtrudeFeature);
            
            //oextrudedef.AffectedBodies.Clear();
            //= default(ObjectCollection);
            //odies = InventorApplication.TransientObjects.CreateObjectCollection();
            //Bodies.Add(oBody);
            if (j == 1)
            {
                
                oextrudedef = oPartCompDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile, PartFeatureOperationEnum.kJoinOperation);
            }
            else if (j == 2)
            {
                oextrudedef = oPartCompDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile, PartFeatureOperationEnum.kNewBodyOperation);
            }
            else
            {
               // oextrudedef = oPartCompDef.Features.
               
                oextrudedef = oPartCompDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile, PartFeatureOperationEnum.kCutOperation);
               // oextrudedef.AffectedBodies.Add(Bodies);
                
               // oextrudedef.AffectedBodies.Clear();
                // oPartCompDef.Features.ExtrudeFeatures.CreateExtrudeDefinition()
            }


            if (i == 1)
            {
                oextrudedef.SetDistanceExtent(length, PartFeatureExtentDirectionEnum.kPositiveExtentDirection);
            }
            else if (i == 2)
            {
                oextrudedef.SetDistanceExtent(length, PartFeatureExtentDirectionEnum.kNegativeExtentDirection);
            }
            else
            {
                oextrudedef.SetDistanceExtent(length, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);
            }

            //oextrudedef.AffectedBodies.Add(oBody);


            //Console.WriteLine();
            //oExtrude.;
            //oExtrude = oPartCompDef.Features.ExtrudeFeatures;
           // oExtrude.SetAffectedBodies(Bodies);
            oExtrude = oPartCompDef.Features.ExtrudeFeatures.Add(oextrudedef);
            

            // oExtrude = oPartCompDef.Features.ExtrudeFeatures[1];


            //Bodies.Remove(0);





        }


        
    }
}

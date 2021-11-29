using System;
using Grasshopper.Kernel;

namespace Rhinventor2021AssemblyGroupBuilder
{
    public class Rhinventor2021AssemblyGroupBuilderComponent : GH_Component
    {
        private static System.Reflection.Assembly ass = System.Reflection.Assembly.GetExecutingAssembly();

        public Rhinventor2021AssemblyGroupBuilderComponent()
          : base("RhinventorAssemblyGroupBuilder", "rhigbuilder",
              $"RhinventorAssemblyGroupBuilder {ass.GetName().Version.ToString()} (Created by Dimitrij Minich). Build nested assemblies",
              "Rhinventor", "Rhinventor 2021")
        {
        }

        protected override void RegisterInputParams(
            GH_Component.GH_InputParamManager pManager
            )
        {
            pManager.AddTextParameter(
                "rootPath", "CP", "Set the root path", GH_ParamAccess.item);
            pManager.AddTextParameter(
                "parentLabel", "PL", "Set parent label", GH_ParamAccess.tree);
            pManager.AddPlaneParameter(
                 "parentPlane", "PPL", "Set parent plane", GH_ParamAccess.tree);
            pManager.AddGenericParameter(
                "parentiPropertyData", "PPD", "Set parent iProperty data", GH_ParamAccess.tree);

            pManager.AddTextParameter(
                "childAssemblyFile", "cAF", "Set child assembly file", GH_ParamAccess.tree);
            pManager.AddTextParameter(
                "childLabel", "cL", "Set child label", GH_ParamAccess.tree);
            pManager.AddGenericParameter(
                "childParameterData", "cMD", "Set child parameter data", GH_ParamAccess.tree);
            pManager.AddGenericParameter(
                "childiPropertyData", "cPD", "Set child iProperty data", GH_ParamAccess.tree);
            pManager.AddPlaneParameter(
                 "childPlane", "cPL", "Set child plane", GH_ParamAccess.tree);
            pManager.AddGroupParameter(
                 "childPoints", "cRP", "Set child reference points", GH_ParamAccess.tree);

            pManager.AddBooleanParameter(
                "StartProcess", "B", "Start process", GH_ParamAccess.item);

            pManager.AddTextParameter(
                "OptionSeparator", "O1", "Set the separator for Inventor files, default '_'", GH_ParamAccess.item, "_");
            pManager.AddTextParameter(
                "OptionIdentifier3DSketch", "O2", "Set the 3d sketch identifier for the driver component, default 'SKELETT:1'", GH_ParamAccess.item, "SKELETT:1");
            pManager.AddTextParameter(
                "OptionIdentifierView", "O3", "Set the view identifier, default 'View1'", GH_ParamAccess.item, "View1");
            pManager.AddTextParameter(
                "OptionIdentifierLOD", "O4", "Set the LOD identifier, default 'Master'", GH_ParamAccess.item, "Master");
            pManager.AddBooleanParameter(
                "OptionExportToRFA", "O5", "Set the trigger for exporting the Assembly to Revit Family", GH_ParamAccess.item, false);
        }

        protected override void RegisterOutputParams(
            GH_Component.GH_OutputParamManager pManager
            )
        {
            pManager.AddTextParameter(
                "Status", "S", "Present current status", GH_ParamAccess.list);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string rootPath = "";
            if (!DA.GetData(0, ref rootPath)) return;

            Grasshopper.Kernel.Data.GH_Structure<Grasshopper.Kernel.Types.GH_String> parentLabel;
            if (!DA.GetDataTree(1, out parentLabel)) return;

            Grasshopper.Kernel.Data.GH_Structure<Grasshopper.Kernel.Types.GH_Plane> parentPlane;
            if (!DA.GetDataTree(2, out parentPlane)) return;

            Grasshopper.Kernel.Data.GH_Structure<Grasshopper.Kernel.Types.IGH_Goo> parentiPropertyData;
            if (!DA.GetDataTree(3, out parentiPropertyData)) return;


            Grasshopper.Kernel.Data.GH_Structure<Grasshopper.Kernel.Types.GH_String> childAssemblyFile;
            if (!DA.GetDataTree(4, out childAssemblyFile)) return;

            Grasshopper.Kernel.Data.GH_Structure<Grasshopper.Kernel.Types.GH_String> childLabel;
            if (!DA.GetDataTree(5, out childLabel)) return;

            Grasshopper.Kernel.Data.GH_Structure<Grasshopper.Kernel.Types.IGH_Goo> childParameterData;
            if (!DA.GetDataTree(6, out childParameterData)) return;

            Grasshopper.Kernel.Data.GH_Structure<Grasshopper.Kernel.Types.IGH_Goo> childiPropertyData;
            if (!DA.GetDataTree(7, out childiPropertyData)) return;

            Grasshopper.Kernel.Data.GH_Structure<Grasshopper.Kernel.Types.GH_Plane> childPlane;
            if (!DA.GetDataTree(8, out childPlane)) return;

            Grasshopper.Kernel.Data.GH_Structure<Grasshopper.Kernel.Types.GH_GeometryGroup> childPoints;
            if (!DA.GetDataTree(9, out childPoints)) return;


            bool startProcess = false;
            DA.GetData(10, ref startProcess);
            if (!startProcess) return;


            string separator = "";
            if (!DA.GetData(11, ref separator)) return;

            string identifier3DSketch = "";
            if (!DA.GetData(12, ref identifier3DSketch)) return;

            string identifierView = "";
            if (!DA.GetData(13, ref identifierView)) return;

            string identifierLOD = "";
            if (!DA.GetData(14, ref identifierLOD)) return;

            bool exportRFA = false;
            DA.GetData(15, ref exportRFA);


            CustomOptions options = new CustomOptions
            {
                separator = separator,
                identifier3DSketch = identifier3DSketch,
                identifierView = identifierView,
                identifierLOD = identifierLOD,
                exportRFA = exportRFA
            };

            ProcessController mainProcess = new ProcessController()
            {
                rootPath = rootPath,
                parentLabel = parentLabel,
                parentPlane = parentPlane,
                parentiPropertyData = parentiPropertyData,
                childAssemblyFile = childAssemblyFile,
                childLabel = childLabel,
                childParameterData = childParameterData,
                childiPropertyData = childiPropertyData,
                childPlane = childPlane,
                childPoints = childPoints,
                ghOutput = DA,
                options = options
            };

            mainProcess.startProcess();


        }

        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                return Resources.Rhinventor2021AssemblyGroupBuilder24;
            }
        }
        public override Guid ComponentGuid
        {
            get { return new Guid("7942b2c4-cb09-4fc2-9c38-1e5fa02f781d"); }
        }
    }
}

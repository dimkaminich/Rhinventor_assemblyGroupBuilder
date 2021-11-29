using System;
using System.Collections.Generic;
using System.Linq;
using Gh_Data = Grasshopper.Kernel.Data;
using Gh_Types = Grasshopper.Kernel.Types;
using System.IO;
using Inventor;

namespace Rhinventor2021AssemblyGroupBuilder
{
    class ProcessController
    {
        public string rootPath { get; set; }
        public Gh_Data.GH_Structure<Gh_Types.GH_String> parentLabel { get; set; }
        public Gh_Data.GH_Structure<Gh_Types.GH_Plane> parentPlane { get; set; }
        public Gh_Data.GH_Structure<Gh_Types.IGH_Goo> parentiPropertyData { get; set; }

        public Gh_Data.GH_Structure<Gh_Types.GH_String> childAssemblyFile { get; set; }
        public Gh_Data.GH_Structure<Gh_Types.GH_String> childLabel { get; set; }
        public Gh_Data.GH_Structure<Gh_Types.IGH_Goo> childParameterData { get; set; }
        public Gh_Data.GH_Structure<Gh_Types.IGH_Goo> childiPropertyData { get; set; }
        public Gh_Data.GH_Structure<Gh_Types.GH_Plane> childPlane { get; set; }
        public Gh_Data.GH_Structure<Gh_Types.GH_GeometryGroup> childPoints { get; set; }

        public Grasshopper.Kernel.IGH_DataAccess ghOutput { get; set; }

        public CustomOptions options { get; set; }

        public void startProcess()
        {
            CustomApprenticeServer apprenticeServer = new CustomApprenticeServer(options);
            InventorAssemblySession inventorAssemblySession = new InventorAssemblySession(options);
            try
            {
                List<ParentComponent> parentComponents = createParentComponentInstances();
                parentComponents.ForEach(parentComponent =>
                {
                    string path = parentComponent.createDirectory();
                    parentComponent.childComponents.ForEach(childComponent => apprenticeServer.build(childComponent, path));
                });
                apprenticeServer.close();

                inventorAssemblySession.openSession();
                parentComponents.ForEach(parentcomponent =>
                {
                    AssemblyDocument assemblyDocument = inventorAssemblySession.createAssemblyDocument(parentcomponent);
                    parentcomponent.childComponents.ForEach(childComponent => inventorAssemblySession.placeComponent(assemblyDocument, childComponent, parentcomponent.compPlane));
                    inventorAssemblySession.saveDocument(assemblyDocument);
                    if (options.exportRFA)
                    {
                        inventorAssemblySession.exportToRevitFamily(assemblyDocument, parentcomponent);
                        inventorAssemblySession.saveDocument(assemblyDocument);
                    }
                    inventorAssemblySession.closeDocument(assemblyDocument);
                });

                inventorAssemblySession.closeSession();
                ghOutput.SetData(0, "Building process successfully finished");
            }
            catch (Exception error)
            {
                ghOutput.SetData(0, error);
                apprenticeServer.close();
                inventorAssemblySession.closeSession();
            }
        }

        private List<ParentComponent> createParentComponentInstances()
        {
            List<ParentComponent> parentComponents = new List<ParentComponent>();

            if (!validateParentStructure()) throw new ArgumentException("Wrong tree structure in parent tree");
            if (!validateChildStructure()) throw new ArgumentException("Wrong tree structure in child tree");
            if (!validateRootPath()) throw new FileNotFoundException("Given root folder not found");
            if (!validateParentDirectories()) throw new ArgumentException("Folder already exist");
            if (!validateChildFiles()) throw new FileNotFoundException("Given child assembly file not exist");

            for (int i = 0; i < parentLabel.Branches.Count; i++)
            {
                parentComponents.Add(new ParentComponent()
                {
                    compRootFolderPath = rootPath + parentLabel.Branches[i].First(),
                    compLabel = parentLabel.Branches[i].First().ToString(),
                    compPlane = parentPlane.Branches[i].First().Value,
                    compiPropertyData = parentiPropertyData.Branches[i].Select(x =>
                    {
                        List<Dictionary<string, object>> result;
                        x.CastTo<List<Dictionary<string, object>>>(out result);
                        return result;
                    }).First(),

                    childComponents = createChildComponentInstances(
                        childAssemblyFile.Branches[i].Select(x => x.Value).ToList(),

                        childLabel.Branches[i].Select(x => x.Value).ToList(),

                        childPlane.Branches[i].Select(x => x.Value).ToList(),

                        childPoints.Branches[i].Select(x => x.Objects.Select(y =>
                        {
                            Rhino.Geometry.GeometryBase p = Grasshopper.Kernel.GH_Convert.ToGeometryBase(y);
                            return (Rhino.Geometry.Point)p;
                        }).ToList()).ToList(),

                        childParameterData.Branches[i].Select(x =>
                        {
                            Dictionary<string, string> result;
                            x.CastTo<Dictionary<string, string>>(out result);
                            return result;
                        }).ToList(),

                        childiPropertyData.Branches[i].Select(x =>
                        {
                            List<Dictionary<string, object>> result;
                            x.CastTo<List<Dictionary<string, object>>>(out result);
                            return result;
                        }).ToList()

                       )
                }
                );
            }

            return parentComponents;
        }

        private List<ChildComponent> createChildComponentInstances(List<string> childAssemblyFiles, List<string> childLabels, List<Rhino.Geometry.Plane> childPlanes, List<List<Rhino.Geometry.Point>> childPoints, List<Dictionary<string, string>> childParameterData, List<List<Dictionary<string, object>>> childiPropertyData)
        {
            List<ChildComponent> childComponents = new List<ChildComponent>();

            for (int i = 0; i < childAssemblyFiles.Count(); i++)
            {
                childComponents.Add(
                        new ChildComponent()
                        {
                            compAssemblyFileSource = childAssemblyFiles[i],
                            compLabel = childLabels[i],
                            compPlane = childPlanes[i],
                            compRefPoints = childPoints[i],
                            compParameterData = childParameterData[i],
                            compiPropertyData = childiPropertyData[i],
                        }
                );
            }
            return childComponents;

        }

        private bool validateParentStructure()
        {
            return (parentLabel.Branches.Count == parentPlane.Branches.Count);
        }

        private bool validateChildStructure()
        {
            if (
                childAssemblyFile.Branches.Count != parentLabel.Branches.Count ||
                childLabel.Branches.Count != parentLabel.Branches.Count ||
                childParameterData.Branches.Count != parentLabel.Branches.Count ||
                childiPropertyData.Branches.Count != parentLabel.Branches.Count ||
                childPlane.Branches.Count != parentLabel.Branches.Count ||
                childPoints.Branches.Count != parentLabel.Branches.Count) return false;


            for (int i = 0; i < parentLabel.Branches.Count; i++)
            {
                HashSet<int> branchLength = new HashSet<int>(){
                    childAssemblyFile.Branches[i].Count,
                    childLabel.Branches[i].Count,
                    childParameterData.Branches[i].Count,
                    childiPropertyData.Branches[i].Count,
                    childPlane.Branches[i].Count,
                    childPoints.Branches[i].Count};

                if (branchLength.Count > 1) return false;

            }

            return true;
        }

        private bool validateRootPath()
        {
            if (!System.IO.Directory.Exists(rootPath))
            {
                return false;
            }
            return true;
        }

        private bool validateParentDirectories()
        {
            foreach (var branch in parentLabel.Branches)
            {
                string parentFolderPath = rootPath + branch.First();
                if (System.IO.Directory.Exists(parentFolderPath))
                {
                    return false;
                }
            }
            return true;
        }

        private bool validateChildFiles()
        {
            HashSet<string> filePaths = new HashSet<string>();

            foreach (var branch in childAssemblyFile.Branches)
            {
                foreach (var filepath in branch.ToList())
                {
                    filePaths.Add(filepath.Value);
                }
            }

            foreach (string filepath in filePaths)
            {
                if (!System.IO.File.Exists(filepath))
                {
                    return false;
                }
            }

            return true;
        }

    }
}


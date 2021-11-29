using Inventor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace Rhinventor2021AssemblyGroupBuilder
{
    class CustomApprenticeServer
    {
        public ApprenticeServerComponent apprenticeServer;
        public CustomOptions options;

        public CustomApprenticeServer(CustomOptions options)
        {
            this.apprenticeServer = new ApprenticeServerComponent();
            this.options = options;
        }

        public void build(ChildComponent childComponent, string parentComponentPath)
        {
            string targetPath;

            try
            {
                targetPath = copy(
                    childComponent.getDrawingFiles(),
                    childComponent.compLabel,
                    parentComponentPath
                    );
            }
            catch (Exception)
            {
                throw new FileNotFoundException("Something goes wrong in apprentice server while copying Inventor files. " +
                    "Please check the Inventor version of the inventor files. The version MUST be the same like the version of your executing Inventor. " +
                    "Apprentice Server is not allowed to migrate Inventor files!");
            }

            childComponent.compAssemblyFileTarget = getAssemblyFileTarget(targetPath);
        }

        private string copy(List<string> files, string compLabel, string parentComponentPath)
        {
            List<string> mappedFiles = remapInventorDrawingPaths(files, compLabel, parentComponentPath);

            for (int i = 0; i < files.Count; i++)
            {
                string sourcefile = files[i];
                string newfile = mappedFiles[i];

                ApprenticeServerDocument drawingDoc = apprenticeServer.Open(sourcefile);
                FileSaveAs fileSaveAs = apprenticeServer.FileSaveAs;
                fileSaveAs.AddFileToSave(drawingDoc, newfile);
                ApprenticeServerDocuments referenceDocs = drawingDoc.AllReferencedDocuments;
                replaceReferences(
                    referenceDocs,
                    System.IO.Path.GetDirectoryName(sourcefile),
                    System.IO.Path.GetDirectoryName(newfile),
                    compLabel,
                    ref fileSaveAs
                    );
                fileSaveAs.ExecuteSaveCopyAs();
                closeReferences(referenceDocs);
            }

            return System.IO.Path.GetDirectoryName(mappedFiles[0]);
        }

        private void replaceReferences(ApprenticeServerDocuments referenceDocs, string sourcepath, string newpath, string compLabel, ref FileSaveAs fileSaveAs)
        {
            foreach (ApprenticeServerDocument referenceDoc in referenceDocs)
            {
                //FullFileName: C:\opath1\opath2\opath3\a.ipt
                //FullFileName: C:\opath1\opath2\opath3\part\a.ipt
                //-------------------------------------
                //C:\opath1\opath3
                //C:\opath1\opath3\part
                string referenceDocFullPath = System.IO.Path.GetDirectoryName(referenceDoc.FullFileName);

                //sourcepath: C:\opath1\opath2\opath3
                //-------------------------------------
                string sourcepathDirectory = System.IO.Path.GetFileName(sourcepath);//opath3
                int strIndex = referenceDocFullPath.IndexOf(sourcepathDirectory) + sourcepathDirectory.Length;//(C:\opath1\opath2\).length + (opath3).length
                string newReferenceDocPath = newpath;//newpath: D:\new
                if (strIndex < referenceDocFullPath.Length)
                {
                    string referenceDocPath = System.IO.Path.GetFileName(referenceDocFullPath.Substring(strIndex));//part
                    newReferenceDocPath += $"\\{compLabel}{options.separator}{referenceDocPath}";//D:\new\part

                }

                //a.ipt
                string referenceDocFileName = System.IO.Path.GetFileName(referenceDoc.FullFileName);

                //D:\new\a.ipt or D:\new\part\a.ipt
                string newReferenceDocFullFilePath = $"{newReferenceDocPath}\\{compLabel}{options.separator}{referenceDocFileName}";

                if (!System.IO.Directory.Exists(newReferenceDocPath)) System.IO.Directory.CreateDirectory(newReferenceDocPath);

                fileSaveAs.AddFileToSave(referenceDoc, newReferenceDocFullFilePath);
                if (System.IO.File.Exists(newReferenceDocFullFilePath))
                {
                    System.IO.File.Delete(newReferenceDocFullFilePath);
                }
            }
        }

        private void closeReferences(ApprenticeServerDocuments referenceDocs)
        {
            foreach (ApprenticeServerDocument referenceDoc in referenceDocs)
            {
                referenceDoc.Close();
            }
        }

        private List<string> remapInventorDrawingPaths(List<string> files, string compLabel, string parentComponentPath)
        {
            files.ForEach(x =>
            {
                if (!apprenticeServer.FileManager.IsInventorDWG(x))
                {
                    throw new IOException("File is not supported by Inventor");
                }
            });

            List<string> mappedFiles = files.Select(x => mapFilename(x, compLabel, parentComponentPath)).ToList();

            mappedFiles.ForEach(f =>
            {
                if (System.IO.File.Exists(f))
                {
                    throw new IOException("File already exist, delete it and repeat process");
                }
            });

            return mappedFiles;
        }

        private string mapFilename(string fullFilename, string preffix, string newpath)
        {
            string filename = System.IO.Path.GetFileName(fullFilename);
            string path = System.IO.Path.GetDirectoryName(fullFilename);
            path = System.IO.Path.GetFileName(path);
            filename = $"{newpath}\\{preffix}{options.separator}{path}\\{preffix}{options.separator}{filename}";
            return filename;
        }

        private string getAssemblyFileTarget(string path)
        {
            List<string> assemblyFiles = System.IO.Directory.GetFiles(path, "*.iam").ToList();
            if (assemblyFiles.Count == 0)
            {
                throw new FileNotFoundException("Assembly file (iam) for child component not found");
            }
            else if (assemblyFiles.Count > 1)
            {
                throw new IOException("It must be only one assembly file (iam) exist in the child component directory");
            }
            else
            {
                return assemblyFiles[0];
            }
        }

        public void close()
        {
            apprenticeServer.Close();
        }
    }
}



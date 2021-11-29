using System.Collections.Generic;
using System.IO;

namespace Rhinventor2021AssemblyGroupBuilder
{
    class ParentComponent: BaseComponent
    {
        public string compRootFolderPath { get; set; }
        public List<ChildComponent> childComponents { get; set; }


        public string createDirectory()
        {
            if (System.IO.Directory.Exists(compRootFolderPath)) throw new IOException("Folder aready exist");
            System.IO.Directory.CreateDirectory(compRootFolderPath);
            return compRootFolderPath;
        }
    }
}

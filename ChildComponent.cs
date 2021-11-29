using System.Collections.Generic;
using System.IO;
using System.Linq;
using RHG = Rhino.Geometry;

namespace Rhinventor2021AssemblyGroupBuilder
{
    class ChildComponent: BaseComponent
    {
        public string compAssemblyFileSource { get; set; }
        public List<RHG.Point> compRefPoints { get; set; }
        public Dictionary<string, string> compParameterData { get; set; }
        public string compAssemblyFileTarget { get; set; }

        public List<string> getDrawingFiles()
        {
            List<string> files = new List<string>();
            string parentPath = System.IO.Directory.GetParent(compAssemblyFileSource).FullName;
            files.AddRange(System.IO.Directory.GetFiles(parentPath, "*.dwg").ToList());
            files.AddRange(System.IO.Directory.GetFiles(parentPath, "*.idw").ToList());

            if (files.Count == 0) throw new FileNotFoundException("Child folder doesn't contain dwg or idw files");

            files = files
                .Where(x => System.IO.Path.GetExtension(x) != ".lck")
                .Select(x => x)
                .ToList();

            return files;
        }

        public RHG.Plane transformPlaneToParent(RHG.Plane parentPlane)
        {
            RHG.Transform transform = RHG.Transform.PlaneToPlane(parentPlane, RHG.Plane.WorldXY);
            Rhino.Geometry.Plane plane = new Rhino.Geometry.Plane(compPlane);
            plane.Transform(transform);
            return plane;
        }

        public List<RHG.Point3d> remapPointsToPlane()
        {
            List<RHG.Point3d> remapedPoints = compRefPoints.Select(p =>
            {
                RHG.Point3d mp;
                compPlane.RemapToPlaneSpace(new RHG.Point3d(p.Location.X, p.Location.Y, p.Location.Z), out mp);
                return mp;
            }).ToList();

            return remapedPoints;
        }
    }
}

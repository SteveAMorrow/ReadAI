using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bentley.Plant.App.Pid.AIImport
{
    public class Size
    {
        public int width { get; set; }
        public int height { get; set; }
    }

    public class Asset
    {
        public string format { get; set; }
        public string id { get; set; }
        public string name { get; set; }
        public Size size { get; set; }
        public string taskDefinitionId { get; set; }
    }

    public class BoundingBox
    {
        public int height { get; set; }
        public int width { get; set; }
        public int left { get; set; }
        public int top { get; set; }
    }

    public class Point
    {
        public int x { get; set; }
        public int y { get; set; }
    }

    public class AssociatedTo
    {
        public string id { get; set; }
        public int confidence { get; set; }
    }

    public class Region
    {
        public string id { get; set; }
        public string type { get; set; }
        public List<string> tags { get; set; }
        public BoundingBox boundingBox { get; set; }
        public List<Point> points { get; set; }
        public string orientation { get; set; }
        public string symmetry { get; set; }
        public double confidence { get; set; }
        public List<AssociatedTo> associatedTo { get; set; }
        public string text { get; set; }
        public string iModelClass { get; set; }
        public int textStatus { get; set; }
        public string userLabel { get; set; }
        public List<string> userLabelTextsIds { get; set; }
        public double? userLabelConfidence { get; set; }
        //public Properties properties { get; set; }
        public object properties { get; set; }
    }

    public class ConnectionPoints
    {

    }

    public class Properties
    {
        public string type { get; set; }

    }

    public class IModelProperties
    {
    }

    public class Links
    {
    }

    public class AIRoot
    {
        public Asset asset { get; set; }
        public List<Region> regions { get; set; }
        public ConnectionPoints connectionPoints { get; set; }
        public Links links { get; set; }
    }

}

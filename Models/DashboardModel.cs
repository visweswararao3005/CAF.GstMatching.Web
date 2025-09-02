using System.Runtime.Serialization;

namespace CAF.GstMatching.Web.Models
{
    public class DashboardModel
    {
    }

    public class DataPointPie1
    {
        public DataPointPie1(string label, double y)
        {
            this.Label = label;
            this.Y = y;
        }

        //Explicitly setting the name to be used while serializing to JSON.
        [DataMember(Name = "label")]
        public string Label = "";

        //Explicitly setting the name to be used while serializing to JSON.
        [DataMember(Name = "y")]
        public Nullable<double> Y = null;
    }
}
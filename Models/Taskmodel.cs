using System.Runtime.Serialization;

namespace CAF.GstMatching.Web.Models
{
    public class Taskmodel
    {
        [DataContract]
        public class DataPoint
        {
            public DataPoint(string label, double y)
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

        public class Taskdetails
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public string City { get; set; }
            public string Address { get; set; }
        }
    }
}
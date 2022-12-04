using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Web;
using System;
using System.Globalization;
using System.IO;

namespace ms_graph_app_auth
{
    public class NewEvent
    {
        public string Id { get; set; }
        public string ICalUId { get; set; }
        public string Subject { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        [DataType(DataType.MultilineText)]
        public string Body { get; set; }
        public List<string> Attendees { get; set; }
        public string RequestType { get; set; }
        public string Committee { get; set; }
        public string Location { get; set; }
        public string Status { get; set; }
        public string Classification { get; set; }
        public string Priority { get; set; }
        public string MeetingLocation { get; set; }
        public string Period { get; set; }

        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }

        public bool EventRecurring { get; set; }

        public static NewEvent ReadFromJsonFile(string path)
        {
            IConfigurationRoot Configuration;

            var builder = new ConfigurationBuilder()
             .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile(path);

            Configuration = builder.Build();
            return Configuration.Get<NewEvent>();
        }
    }
}

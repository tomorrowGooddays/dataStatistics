using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace GDDataStatistics.Helper
{
    public static class FileHelper
    {
        public static T GetJsonFileFromEmbedResource<T>(string resourceName)
        {
            StreamReader r = new StreamReader(resourceName);
            string jsonString = r.ReadToEnd();
            return JsonConvert.DeserializeObject<T>(jsonString);
        }
    }
}

using System;
using System.Linq;
using System.Xml.Linq;

namespace ElmahHelper.Model
{
    public class Error
    {
        public DateTime Time { get; set; }

        private string ErrorID { get; set; }

        private string Application { get; set; }

        private string Host { get; set; }

        private string Type { get; set; }

        public string Message { get; set; }

        private string Source { get; set; }

        private string Detail { get; set; }

        public void SetInfo(XDocument info)
        {
            var errorElement = info.Descendants("error").FirstOrDefault();
            if (errorElement == null) return;

            ErrorID = errorElement.Attribute("errorId")?.Value;
            Application = errorElement.Attribute("application")?.Value;
            Host = errorElement.Attribute("host")?.Value;
            Type = errorElement.Attribute("type")?.Value;
            Message = errorElement.Attribute("message")?.Value;
            Source = errorElement.Attribute("source")?.Value;
            Detail = errorElement.Attribute("detail")?.Value;
            string timeAttr = errorElement.Attribute("time")?.Value;
            if (DateTime.TryParse(timeAttr, out DateTime time))
            {
                Time = time;
            }
        }

        public void SetEmptyInfo(string elmahName)
        {
            Message = "讀取失敗";
            Detail = $"{elmahName}讀取失敗";
        }

        public string GetDetail()
        {
            return Detail;
        }
    }
}

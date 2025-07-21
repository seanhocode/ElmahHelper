using ElmahHelper.Tools;
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

namespace ElmahHelper.Model
{
    public class Elmah
    {
        /// <summary>
        /// ErrorInfo
        /// </summary>
        public Error ElmahError { get; set; }

        /// <summary>
        /// ElmahFileName
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// ElmahTime
        /// </summary>
        public DateTime Time { get; set; }

        /// <summary>
        /// ElmahGUID
        /// </summary>
        public string GUID { get; set; }

        /// <summary>
        /// 來自哪個ZIP(完整路徑)
        /// </summary>
        public string SourceZIPPath { get; set; }

        /// <summary>
        /// 上層資料夾
        /// </summary>
        public string ParentFolderPath { get; set; }

        private XMLTool xmlTool = new XMLTool();
        private ZipTool zipTool = new ZipTool();
        private FileTool fileTool = new FileTool();

        /// <summary>
        /// 建構元(來自ZIP)
        /// </summary>
        /// <param name="elmahName"></param>
        /// <param name="zipPath"></param>
        public Elmah(string elmahName, string zipPath)
        {
            Initial(elmahName);
            SourceZIPPath = zipPath;
            ParentFolderPath = Path.GetDirectoryName(zipPath);
            LoadErrorTimeAndGUIDByName();
            GetElmahErrorInfoByZip();
        }

        /// <summary>
        /// 建構元(來自File)
        /// </summary>
        /// <param name="elmahFilePath"></param>
        public Elmah(string elmahFilePath)
        {
            Initial(Path.GetFileName(elmahFilePath));
            ParentFolderPath = Path.GetDirectoryName(elmahFilePath);
            LoadErrorTimeAndGUIDByName();
            GetElmahErrorInfoByFile();
        }

        /// <summary>
        /// Initial
        /// </summary>
        /// <param name="elmahName"></param>
        public void Initial(string elmahName)
        {
            FileName = elmahName;
            ElmahError = new Error();
        }

        /// <summary>
        /// 從ZIP中取得ElmahInfo
        /// </summary>
        public void GetElmahErrorInfoByZip()
        {
            if (!fileTool.CheckFileExist(SourceZIPPath))
                Debug.WriteLine($"找不到ZIP檔，請檢查此路徑{SourceZIPPath}");

            byte[] elmahBytes = zipTool.ExtractSingleFileToMemory(SourceZIPPath, FileName);

            if(elmahBytes == null)
                Debug.WriteLine($"ZIP檔裡面無{FileName}");
            else
                try
                {
                    ElmahError.SetInfo(xmlTool.GetXDocumentFromBytes(elmahBytes));
                }
                catch
                {
                    ElmahError.SetEmptyInfo(FileName);
                }
                
        }

        /// <summary>
        /// 從File取得ElmahInfo
        /// </summary>
        public void GetElmahErrorInfoByFile()
        {
            string filePath = Path.Combine(ParentFolderPath, FileName);

            if (!fileTool.CheckFileExist(filePath))
                Debug.WriteLine($"找不到elmah檔，請檢查此路徑{filePath}");

            ElmahError.SetInfo(xmlTool.GetXDocument(filePath));
        }

        public static (DateTime ElmahTime, string GUID)? GetElmahFileNameData(string elmahName)
        {
            // Regex說明：
            // 1. 時間戳記：4位數年份-2位月-2位日 + 6位時分秒 + 'Z'
            // 2. GUID：標準 8-4-4-4-12 格式的 UUID
            string pattern = @"(\d{4}-\d{2}-\d{2}\d{6}Z)-([0-9a-fA-F\-]{36})";

            Match match = Regex.Match(elmahName, pattern);

            if (match.Success && match.Groups.Count >= 3)
                return (ConvertZFormatToTaiwanTime(match.Groups[1].Value).Value, match.Groups[2].Value);

            return null;
        }

        /// <summary>
        /// 從檔案名稱中取得錯誤時間戳記與GUID
        /// </summary>
        private void LoadErrorTimeAndGUIDByName()
        {
            var elmahData = GetElmahFileNameData(FileName);
            Time = elmahData.Value.ElmahTime;
            GUID = elmahData.Value.GUID;
        }

        /// <summary>
        /// 將特殊格式的UTC時間字串轉換為台灣時間(UTC+8)。
        /// </summary>
        /// <param name="zTimeStr">時間字串，格式:yyyy-MM-ddHHmmssZ</param>
        /// <returns>轉換後的台灣時間(DateTime)，若格式錯誤則回傳 null</returns>
        private static DateTime? ConvertZFormatToTaiwanTime(string zTimeStr)
        {
            //定義輸入時間的格式('Z' 是字面值，需加上單引號)
            const string format = "yyyy-MM-ddHHmmss'Z'";

            // 嘗試將輸入字串解析為 DateTime 物件(UTC)
            if (DateTime.TryParseExact(
                    zTimeStr
                    , format
                    , CultureInfo.InvariantCulture //文化資訊，InvariantCulture確保解析不受本地文化影響
                    , DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal //說明輸入字串為UTC，並將結果轉為UTC
                    , out DateTime utcDateTime
             ))
            {
                try
                {
                    // 取得台灣的時區資訊(Windows時區ID)
                    TimeZoneInfo taiwanTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Taipei Standard Time");

                    // 將UTC時間轉換為台灣時間
                    DateTime taiwanTime = TimeZoneInfo.ConvertTimeFromUtc(utcDateTime, taiwanTimeZone);

                    return taiwanTime;
                }
                catch (TimeZoneNotFoundException)
                {
                    Debug.WriteLine("找不到台灣時區資訊(Taipei Standard Time)");
                }
                catch (InvalidTimeZoneException)
                {
                    Debug.WriteLine("台灣時區設定無效");
                }
            }
            else
            {
                Debug.WriteLine("時間格式解析失敗：" + zTimeStr);
            }

            return null;
        }
    }
}

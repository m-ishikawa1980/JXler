using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using JXler.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;

namespace JXler.Libraries
{
    public static class Utils
    {

        public enum ExecAction
        {
            Rl,
            Lr,
            None
        }

        public const string Rl = "<<";
        public const string Lr = ">>";
        public const string None = "<>";

        public const string cellNull = "{null}";
        public const string cellBlank = "{blank}";

        public const string settingsPath = "Settings";
        public const string settingsFile = "Settings.json";

        public static Settings GetSettings()
        {
            //エクセルシート情報
            var settingJsonPath =
                Path.Combine(ComplementRelativeDir(settingsPath), settingsFile);
            var file = File.ReadAllText(settingJsonPath); // ファイル内容をjson変数に格納
            return file.Deserialize<Settings>();
        }

        public static void SaveSettings(this Settings settings)
        {
            var settingJsonPath =
                Path.Combine(ComplementRelativeDir(settingsPath), settingsFile);
            using (StreamWriter sw = new StreamWriter(settingJsonPath, false))
            {
                sw.WriteLine(settings.Serialize(Formatting.Indented));
            }
        }

        public static string InputFileCheck(string inputFile)
        {            
            if (!File.Exists(inputFile))
            {
                return "入力ファイルが存在しません。";
            }
            return null;
        }

        public static string OutputFileCheck(string outputFile)
        {
            if (!File.Exists(outputFile))
            {
                return null;
            }

            try
            {
                // 書き込みモードでファイルを開けるか確認
                using (FileStream fp = File.Open(outputFile, FileMode.Open, FileAccess.Write))
                {
                    // 開ける
                    return null;
                }
            }
            catch
            {
                // 開けない
                return "出力ファイルチェックエラー　書き込み出来ません。";
            }
        }

        public static List<JsonXls> RenumberingJsonXlsList(this List<JsonXls> jsonToExcelList)
        {
            int index = 1;
            foreach(var jsonToExcel in jsonToExcelList)
            {
                jsonToExcel.No = index;
                index++;
            }
            return jsonToExcelList;
        }

        public static string GetExecAction(ExecAction execAction)
        {
            switch (execAction)
            {
                case ExecAction.Lr:
                    return Lr;
                case ExecAction.Rl:
                    return Rl;
                case ExecAction.None:
                    return None;
                default:
                    return None;
            }
        }

        public static ExecAction CheckExecAction(string execAction)
        {
            switch (execAction)
            {
                case Lr:
                    return ExecAction.Lr;
                case Rl:
                    return ExecAction.Rl;
                case None:
                    return ExecAction.None;
                default:
                    return ExecAction.None;
            }
        }

        public static string ComplementRelativePath(string dir, string file)
        {
            if(string.IsNullOrEmpty(dir) || string.IsNullOrEmpty(file))
            {
                return null;
            }else
            {
                return Path.Combine(ComplementRelativeDir(dir: dir), file);
            }          
        }

        public static string ComplementRelativeDir(string dir)
        {
            if (string.IsNullOrEmpty(dir))
            {
                return null;
            }
            else
            {
                return Path.IsPathRooted(dir) ?
                    dir :
                    Path.GetFullPath(dir);
            }
        }

        public static void ConvIXLCell(
            this IXLCell cell,
            JToken value,
            XlsSettings xlsSettings)
        {
            switch (value.Type)
            {
                case JTokenType.String:
                    cell.DataType = XLDataType.Text;
                    if (!string.IsNullOrEmpty(xlsSettings.XlsCellStyles.Text))
                    {
                        cell.Style.NumberFormat.Format = xlsSettings.XlsCellStyles.Text;
                    }
                    cell.Value = string.IsNullOrEmpty(value.ToString())
                        ? cellBlank
                        : value.ToString();
                    break;
                case JTokenType.Integer:
                    cell.DataType = XLDataType.Number;
                    if (!string.IsNullOrEmpty(xlsSettings.XlsCellStyles.NumberInt))
                    {
                        cell.Style.NumberFormat.Format = xlsSettings.XlsCellStyles.NumberInt;
                    }
                    cell.Value = int.Parse(value.ToString());
                    break;
                case JTokenType.Float:
                    cell.DataType = XLDataType.Number;
                    if (!string.IsNullOrEmpty(xlsSettings.XlsCellStyles.NumberFloat))
                    {
                        cell.Style.NumberFormat.Format = xlsSettings.XlsCellStyles.NumberFloat;
                    }
                    cell.Value = float.Parse(value.ToString());
                    break;
                case JTokenType.Date:
                    cell.DataType = XLDataType.DateTime;
                    if (!string.IsNullOrEmpty(xlsSettings.XlsCellStyles.DateTime))
                    {
                        cell.Style.NumberFormat.Format = xlsSettings.XlsCellStyles.DateTime;
                    }
                    cell.Value = DateTime.Parse(value.ToString());
                    break;
                case JTokenType.Null:
                    cell.DataType = XLDataType.Text;
                    cell.Value = cellNull;
                    break;
                case JTokenType.Boolean:
                    cell.DataType = XLDataType.Boolean;
                    cell.Value = bool.Parse(value.ToString());
                    break;
                default:
                    cell.Value = value.ToString();
                    break;
            }
        }

        public static object ConvJsonValue(this IXLCell cell)
        {
            switch (cell.DataType)
            {
                case XLDataType.Text:
                    switch (cell.Value.ToString())
                    {
                        case cellNull:
                            return null;
                        case cellBlank:
                            return string.Empty;
                        default:
                            return cell.Value.ToString();
                    }
                case XLDataType.Number:
                    float f;
                    int i;
                    if (int.TryParse(cell.Value.ToString(), out i))
                    {
                        return i;
                    }
                    else if (float.TryParse(cell.Value.ToString(), out f))
                    {
                        return f;
                    }
                    else
                    {
                        return cell.Value.ToString();
                    }
                case XLDataType.DateTime:
                    return DateTime.Parse(cell.Value.ToString());
                case XLDataType.Boolean:
                    return bool.Parse(cell.Value.ToString());
                default:
                    return cell.Value.ToString();
            }
        }

        public static string Serialize(
            this object obj,
            Formatting formatting = Formatting.None)
        {
            var settings = new JsonSerializerSettings();
            settings.NullValueHandling = NullValueHandling.Ignore;
            settings.Formatting = formatting;
            return JsonConvert.SerializeObject(obj, settings);
        }

        public static T Deserialize<T>(this string str)
        {
            return JsonConvert.DeserializeObject<T>(str);
        }
    }
}

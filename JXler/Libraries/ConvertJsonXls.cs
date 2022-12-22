using ClosedXML.Excel;
using JXler.Models;
using Newtonsoft.Json;
using System.IO;

namespace JXler.Libraries
{
    public static class ConvertJsonXls
    {

        public static void Convert(MainWindow mainWindow)
        {
            var settings = Utils.GetSettings();
            mainWindow.WriteLogJsonXls(msg: "Convert 開始");
            foreach (var jsonXls in settings.JsonXlsHash)
            {
                switch (Utils.CheckExecAction(execAction: jsonXls.Action))
                {
                    case Utils.ExecAction.Lr:
                        ConvJsonToXls(
                            mainWindow: mainWindow,
                            jsonXls: jsonXls,
                            settings: settings);
                        break;
                    case Utils.ExecAction.Rl:
                        ConvXlsToJson(
                            mainWindow: mainWindow,
                            jsonToXls: jsonXls,
                            settings: settings);
                        break;
                }
            }
            settings.SaveSettings();
            settings = Utils.GetSettings();
            mainWindow.dataGridJsonXls.ItemsSource = settings.JsonXlsHash;
            mainWindow.dataGridJsonXls.Items.Refresh();
            mainWindow.WriteLogJsonXls(msg: "Convert 終了");
        }

        public static void ConvJsonToXls(
            MainWindow mainWindow,
            JsonXls jsonXls,
            Settings settings)
        {
            var inputFileName =
                Utils.ComplementRelativePath(
                    dir: jsonXls.JsonPath,
                    file: jsonXls.JsonName);
            mainWindow.WriteLogJsonXls(msg: $"Json入力 入力ファイル={inputFileName}");
            var inputMsg = Utils.InputFileCheck(inputFile: inputFileName);
            if (!string.IsNullOrEmpty(inputMsg))
            {
                mainWindow.WriteLogJsonXls(msg: $"{inputMsg}", LogLevel.Error);
                return;
            }
            jsonXls.ComplementXlsFilePath(settings: settings);
            var outFileName =
                Utils.ComplementRelativePath(
                    dir: jsonXls.XlsPath,
                    file: jsonXls.XlsName);
            mainWindow.WriteLogJsonXls(msg: $"Excel変換 出力ファイル={outFileName}");
            var outputMsg = Utils.OutputFileCheck(outputFile: outFileName);
            if (!string.IsNullOrEmpty(outputMsg))
            {
                mainWindow.WriteLogJsonXls(
                    msg: $"{outputMsg}",
                    logLevel: LogLevel.Error);
                return;
            }
            var json = File.ReadAllText(inputFileName);
            var xlWorkBook = ToXls.Convert(
                json: json,
                xlsSettings: new XlsSettings(settings: settings));
            if (xlWorkBook == null)
            {
                mainWindow.WriteLogJsonXls(
                    msg: $"Excel変換エラー 対象ファイル={inputFileName}",
                    logLevel: LogLevel.Error);
                return;
            }
            xlWorkBook.SaveAs(outFileName);
            mainWindow.WriteLogJsonXls(msg: $"Excel変換終了");
        }

        public static JsonXls ComplementXlsFilePath(
            this JsonXls jsonXls,
            Settings settings)
        {
            switch (settings.ExecPtn)
            {
                case ExecPtn.SameInput:
                    jsonXls.XlsPath = jsonXls.JsonPath;
                    jsonXls.XlsName = jsonXls.JsonName.Replace("json", "xlsx");
                    break;
                case ExecPtn.SpecifyPath:
                    jsonXls.XlsPath = settings.Path;
                    jsonXls.XlsName = jsonXls.JsonName.Replace("json", "xlsx");
                    break;
                case ExecPtn.SetIndividually:
                    if (string.IsNullOrEmpty(jsonXls.XlsPath))
                    {
                        jsonXls.XlsPath = settings.BasePath;
                    }
                    if (string.IsNullOrEmpty(jsonXls.XlsName))
                    {
                        jsonXls.XlsName = jsonXls.JsonName.Replace("json", "xlsx");
                    }
                    break;
            }
            return jsonXls;
        }

        public static JsonXls ComplementJsonFilePath(
            this JsonXls jsonXls,
            Settings settings)
        {
            switch (settings.ExecPtn)
            {
                case ExecPtn.SameInput:
                    jsonXls.JsonPath = jsonXls.XlsPath;
                    jsonXls.JsonName = jsonXls.XlsName.Replace("xlsx", "json");
                    break;
                case ExecPtn.SpecifyPath:
                    jsonXls.JsonPath = settings.Path;
                    jsonXls.JsonName = jsonXls.XlsName.Replace("xlsx", "json");
                    break;
                case ExecPtn.SetIndividually:
                    break;
            }
            return jsonXls;
        }

        private static void ConvXlsToJson(
            MainWindow mainWindow,
            JsonXls jsonToXls,
            Settings settings)
        {
            var inputFileName =
                Utils.ComplementRelativePath(
                    dir: jsonToXls.XlsPath,
                    file: jsonToXls.XlsName);
            mainWindow.WriteLogJsonXls(msg: $"Excel入力 入力ファイル={inputFileName}");
            var inputMsg = Utils.InputFileCheck(inputFile: inputFileName);
            if (!string.IsNullOrEmpty(inputMsg))
            {
                mainWindow.WriteLogJsonXls(
                    msg: $"{inputMsg}",
                    logLevel: LogLevel.Error);
                return;
            }
            jsonToXls.ComplementJsonFilePath(settings: settings);
            var outFileName =
                Utils.ComplementRelativePath(
                    dir: jsonToXls.JsonPath,
                    file: jsonToXls.JsonName);
            mainWindow.WriteLogJsonXls(msg: $"Json変換 出力ファイル={outFileName}");
            var outputMsg = Utils.OutputFileCheck(outputFile: outFileName);
            if (!string.IsNullOrEmpty(outputMsg))
            {
                mainWindow.WriteLogJsonXls(
                    msg: $"{outputMsg}",
                    logLevel: LogLevel.Error);
                return;
            }
            FileStream fs = new FileStream(
                inputFileName,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite);            
            var jsonObject = ToJson.Convert(
                xlWorkBook: new XLWorkbook(fs, XLEventTracking.Disabled),
                xlsSettings: new XlsSettings(settings: settings));
            if(jsonObject == null)
            {
                mainWindow.WriteLogJsonXls(
                    msg: $"Json変換エラー 対象ファイル={inputFileName}",
                    logLevel: LogLevel.Error);
                return;
            }
            using (StreamWriter sw = new StreamWriter(outFileName, false))
            {
                sw.WriteLine(jsonObject.Serialize(Formatting.Indented));
            }
            mainWindow.WriteLogJsonXls(msg: $"Json変換終了");
        }
    }
}

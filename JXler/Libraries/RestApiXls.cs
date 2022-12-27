using ClosedXML.Excel;
using JXler.Models;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Policy;
using System.Text;
using System.Text.Json.Nodes;
using System.Threading.Tasks;

namespace JXler.Libraries
{
    public static class RestApiXls {

        public static async Task Rest(MainWindow mainWindow)
        {

            var settings = Utils.GetSettings();
            mainWindow.WriteLogJsonXls(msg: "RestApi 開始");
            foreach (var apiXls in settings.ApiXlsHash)
            {

                var inputFileName =
                    Utils.ComplementRelativePath(
                        dir: apiXls.ReqPath,
                        file: apiXls.ReqName);
                mainWindow.WriteLogApiXls(msg: $"Excel入力 入力ファイル={inputFileName}");
                var inputMsg = Utils.InputFileCheck(inputFile: inputFileName);
                if (!string.IsNullOrEmpty(inputMsg))
                {
                    mainWindow.WriteLogApiXls(
                        msg: $"{inputMsg}",
                        logLevel: LogLevel.Error);
                    return;
                }
                //https://qiita.com/c-yan/items/6e506399675e3cc56732
                //https://teratail.com/questions/219712
                FileStream fs = new FileStream(
                    inputFileName,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.ReadWrite);




                var reqParams = ToJson.Convert(
                    xlWorkBook: new XLWorkbook(fs, XLEventTracking.Disabled),
                    xlsSettings: new XlsSettings(settings: settings),
                    sheet: "param").Serialize().Deserialize<JObject>();
                var reqHeaders = ToJson.Convert(
                    xlWorkBook: new XLWorkbook(fs, XLEventTracking.Disabled),
                    xlsSettings: new XlsSettings(settings: settings),
                    sheet: "headers").Serialize().Deserialize<JObject> ();
                var reqBody = ToJson.Convert(
                    xlWorkBook: new XLWorkbook(fs, XLEventTracking.Disabled),
                    xlsSettings: new XlsSettings(settings: settings),
                    sheet: "body").Serialize();
                var url = reqParams.Value<string>("url");
                var contentType = reqHeaders.Value<string>("content-type");
                var content = new StringContent(reqBody, Encoding.UTF8);
                var request = new HttpRequestMessage(HttpMethod.Post, url);
                request.Headers.Add("ContentType", contentType);
                request.Content = content;
                var Http = new HttpClient();
                var response = await Http.SendAsync(request);
                var responseContent = await response.Content.ReadAsStringAsync();



                var outFileName =
                    Utils.ComplementRelativePath(
                        dir: apiXls.ResPath,
                        file: apiXls.ResName);
                mainWindow.WriteLogApiXls(msg: $"Excel変換 出力ファイル={outFileName}");
                var outputMsg = Utils.OutputFileCheck(outputFile: outFileName);
                if (!string.IsNullOrEmpty(outputMsg))
                {
                    mainWindow.WriteLogJsonXls(
                        msg: $"{outputMsg}",
                        logLevel: LogLevel.Error);
                    return;
                }
                var xlWorkBook = ToXls.Convert(
                    json: responseContent,
                    xlsSettings: new XlsSettings(settings: settings));
                xlWorkBook.SaveAs(outFileName);
                //switch (Utils.CheckExecAction(execAction: jsonXls.Action))
                //{
                //    case Utils.ExecAction.Lr:
                //        ConvJsonToXls(
                //            mainWindow: mainWindow,
                //            jsonXls: jsonXls,
                //            settings: settings);
                //        break;
                //    case Utils.ExecAction.Rl:
                //        ConvXlsToJson(
                //            mainWindow: mainWindow,
                //            jsonToXls: jsonXls,
                //            settings: settings);
                //        break;
                //}
            }
            settings.SaveSettings();
            settings = Utils.GetSettings();
            mainWindow.dataGridApiXls.ItemsSource = settings.ApiXlsHash;
            mainWindow.dataGridApiXls.Items.Refresh();
            mainWindow.WriteLogJsonXls(msg: "RestApi 終了");

        }
    }
}

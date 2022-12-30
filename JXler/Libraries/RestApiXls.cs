using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using JXler.Models;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
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






                //var Http = new HttpClient();
                //var response = await Http.SendAsync(SetRequest(inputFileName:inputFileName));
                //var responseContent = await response.Content.ReadAsStringAsync();

                var responseString =
                    await HttpClient(
                        SetRequest(
                            inputFileName: inputFileName));

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
                    json: responseString,
                    xlsSettings: new XlsSettings(settings: settings));
                xlWorkBook.SaveAs(outFileName);

            }
            settings.SaveSettings();
            settings = Utils.GetSettings();
            mainWindow.dataGridApiXls.ItemsSource = settings.ApiXlsHash;
            mainWindow.dataGridApiXls.Items.Refresh();
            mainWindow.WriteLogJsonXls(msg: "RestApi 終了");

        }

        private static HttpRequestMessage SetRequest(string inputFileName)
        {
            var settings = Utils.GetSettings();
            FileStream fs = new FileStream(
                inputFileName,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite);

            var xlsWorkBook = new XLWorkbook(fs, XLEventTracking.Disabled);
            var xlsSettings = new XlsSettings(settings: settings);

            var reqParams = ToJson.Convert(
                xlWorkBook: xlsWorkBook,
                xlsSettings: xlsSettings,
                sheet: "param")
                    .Serialize()
                        .Deserialize<JObject>();
            var reqHeaders = ToJson.Convert(
                xlWorkBook: xlsWorkBook,
                xlsSettings: xlsSettings,
                sheet: "headers")
                    .Serialize()
                        .Deserialize<JToken>();
            var reqBody = ToJson.Convert(
                xlWorkBook: xlsWorkBook,
                xlsSettings: xlsSettings,
                sheet: "body")
                    .Serialize();

            var url = string.Empty;

            if (reqParams.ContainsKey("url"))
            {
                url = reqParams.Value<string>("url");
            }
            HttpMethod method;// = HttpMethod.Post;
            if (reqParams.ContainsKey("method"))
            {
                switch (reqParams.Value<string>("method"))
                {
                    case "post":
                        method = HttpMethod.Post;
                        break;
                    case "get":
                        method = HttpMethod.Get;
                        break;
                    default:
                        method = HttpMethod.Get;
                        break;
                }
            }
            else
            {
                method = HttpMethod.Get;
            }
            var request = new HttpRequestMessage(method, url);
            var jsonDic = reqHeaders.ToObject<Dictionary<string, string>>();
            foreach (var data in jsonDic)
            {
                request.Headers.Add(data.Key, data.Value);
            }
            request.Content = new StringContent(reqBody, Encoding.UTF8);

            return request;
        }

        private static async Task<string> HttpClient(HttpRequestMessage requestMessage)
        {
            var Http = new HttpClient();
            var response = await Http.SendAsync(requestMessage);
            return await response.Content.ReadAsStringAsync();            
        }
    }
}

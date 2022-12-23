using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using JXler.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JXler.Libraries
{
    public static class ToXls
    {
        public static XLWorkbook Convert(
            string json,
            XlsSettings xlsSettings)
        {
            var xlWorkBook = new XLWorkbook();            
            try
            {
                var jo = JsonConvert.DeserializeObject<JToken>(json);
                CheckType(
                    jo: jo,
                    xlWorkBook: xlWorkBook,
                    sheetName: xlsSettings.IndexSheet,
                    xlsSettings: xlsSettings);
                return xlWorkBook;
            }
            catch
            {
                return null;
            };
            
        }

        public static XLWorkbook CreateWorkSheet(
            this XLWorkbook xlWorkBook,
            string sheetName,
            JTokenType joType)
        {
            IXLWorksheet sheet;
            if (xlWorkBook.TryGetWorksheet(sheetName, out sheet))
            {
                return xlWorkBook;
            }

            var ws = xlWorkBook.Worksheets.Add(sheetName);
            ws.Cell(1, 1).Value = "No";
            switch (joType)
            {
                case JTokenType.Array:
                    ws.Cell(1, 1).CreateComment().AddText("Array");
                    break;
                default:
                    ws.Cell(1, 1).CreateComment().AddText("Object");
                    break;
            }
            return xlWorkBook;
        }

        public static int CheckType(
            JToken jo,
            string sheetName,
            XLWorkbook xlWorkBook,
            XlsSettings xlsSettings)
        {
            switch (jo.Type)
            {
                case JTokenType.Array:
                    return CheckArray(
                        jo: jo,
                        xlWorkBook: xlWorkBook,
                        sheetName: sheetName,
                        xlsSettings: xlsSettings);
                case JTokenType.Object:
                    return CheckObject(
                        jo: jo,
                        xlWorkBook: xlWorkBook,
                        sheetName: sheetName,
                        xlsSettings: xlsSettings);
                default:
                    return 0;
            }        
        }

        private static int CheckArray(
            JToken jo,
            XLWorkbook xlWorkBook,
            string sheetName,
            XlsSettings xlsSettings)
        {
            var ws = xlWorkBook.CreateWorkSheet(sheetName, jo.Type).Worksheet(sheetName);
            var range = ws.Range("A:A").RangeUsed();
            var maxNum = RangeMax(range) + 1;
            foreach (var array in (JArray)jo)
            {
                //ここで行番号を指定する
                //indexの場合は連番
                //以外は受け取ったグループ番号を設定
                CheckObject(
                    jo: array,
                    xlWorkBook: xlWorkBook,
                    sheetName: sheetName,
                    num: maxNum,
                    xlsSettings: xlsSettings);
                if (sheetName == xlsSettings.IndexSheet)
                {
                    maxNum++;
                }
            }
            return maxNum;
        }

        private static int CheckObject(
            JToken jo,
            XLWorkbook xlWorkBook,
            string sheetName,
            XlsSettings xlsSettings,
            int num = 0)
        {
            var ws = xlWorkBook.CreateWorkSheet(sheetName, jo.Type).Worksheet(sheetName);
            var rowNum = ws.LastRowUsed().RowNumber() + 1;
            if (num <= 0)
            {
                var range = ws.Range("A:A").RangeUsed();
                num = RangeMax(range) + 1;
            }

            ws.Cell(rowNum, 1).Value = num;

            if (jo.Type != JTokenType.Array && jo.Type != JTokenType.Object)
            {
                var colNum = GetColumnNum(ixlWorkSheet: ws, key: "List");
                ws.Cell(rowNum, colNum).Value = jo.ToString();
            }
            else
            {
                foreach (var item in (JObject)jo)
                {
                    var colNum = GetColumnNum(ixlWorkSheet: ws, key: item.Key.ToString());
                    if ((item.Value.Type == JTokenType.Array && item.Value.ToList().Count > 0)
                        || item.Value.Type == JTokenType.Object)
                    {
                        xlsSettings.SheetName = item.Key.ToString();
                        var indexNum = CheckType(
                            jo: item.Value,
                            xlWorkBook: xlWorkBook,
                            sheetName: item.Key.ToString(),
                            xlsSettings: xlsSettings);
                        xlsSettings.IndexNum = indexNum;
                        xlsSettings.ChildSheet = item.Key.ToString();
                        WriteChildSheetName(
                            childSheet: item.Key.ToString(),
                            sheetName: sheetName,
                            xlWorkBook: xlWorkBook,
                            xlsSettings: xlsSettings,
                            rowNum: rowNum,
                            colNum: colNum);
                    }
                    else
                    {
                        WriteValue(
                            wb: xlWorkBook,
                            value: item.Value,
                            sheetName: sheetName,
                            xlsSettings: xlsSettings,
                            rowNum: rowNum,
                            colNum: colNum);
                    }
                }
            }
            return num;
        }

        private static int GetColumnNum(
            IXLWorksheet ixlWorkSheet,
            string key)
        {
            int colNum;
            //タイトル行にKeyが存在するかチェック
            var range = ixlWorkSheet.Range("1:1").RangeUsed();
            var searchResultCell = RangeSearch(range, key);
            if (searchResultCell != null)
            {
                colNum = searchResultCell.Address.ColumnNumber;
            }
            else
            {
                //タイトル行にKeyが存在しない場合、最終Column+1にKey文字列を書き込み
                colNum = ixlWorkSheet.Range("1:1").LastColumnUsed().ColumnNumber() + 1;
                ixlWorkSheet.Cell(1, colNum).Value = key;
            }
            return colNum;
        }

        private static void WriteValue(
            XLWorkbook wb,
            JToken value,
            string sheetName,
            XlsSettings xlsSettings,
            int rowNum,
            int colNum)
        {
            if (value.ToString().Length > 32767)
            {
                wb.Worksheet(sheetName).Cell(rowNum, colNum).Value = "文字数オーバー";
            }
            else
            {
                wb.Worksheet(sheetName).Cell(rowNum, colNum)
                    .ConvIXLCell(
                        value: value,
                        xlsSettings: xlsSettings);
            }
        }

        private static void WriteChildSheetName(
            string childSheet,
            string sheetName,
            XLWorkbook xlWorkBook,
            XlsSettings xlsSettings,
            int rowNum,
            int colNum)
        {
            var ws = xlWorkBook.Worksheet(sheetName);
            var value = $"{{{childSheet}_No.{xlsSettings.IndexNum}}}";
            var linkText = LinkFomura(value, xlWorkBook);
            if (linkText != null && xlsSettings.OutputXlsLink)
            {
                ws.Cell(rowNum, colNum).Value = value;
                ws.Cell(rowNum, colNum).SetHyperlink(new XLHyperlink(linkText));
            }
            else
            {
                ws.Cell(rowNum, colNum).Value = value;
            }
        }

        private static string LinkFomura(
            string linkText,
            XLWorkbook xlWorkBook)
        {
            var strArr = linkText.Replace("{", "").Replace("}", "").Split("_No.");
            var cell = RangeSearch(xlWorkBook.Worksheet(strArr[0]).Range("A:A"), strArr[1]);
            if (cell != null)
            {
                return $"{strArr[0]}!{cell.Address.ToString()}";
            }
            else
            {
                return $"{strArr[0]}!A1";
            }
        }

        private static IXLCell RangeSearch(
            IXLRange range,
            string searchText)
        {
            var cells = range.Search(searchText: searchText);
            foreach (var cell in cells)
            {
                if (cell.Value.ToString() == searchText) { return cell; }
            }
            return null;
        }

        private static int RangeMax(IXLRange range)
        {
            var cellList = new List<int>();
            foreach (var cell in range.Cells())
            {
                int i;
                if (int.TryParse(cell.Value.ToString(), out i))
                {
                    cellList.Add(i);
                }
            }
            if (cellList.Count <= 0)
            {
                return 0;
            }
            else
            {
                return cellList.Max();
            }
        }
    }
}

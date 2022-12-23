using ClosedXML;
using ClosedXML.Excel;
using JXler.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JXler.Libraries
{
    public static class ToJson
    {

        public static object Convert(XLWorkbook xlWorkBook, XlsSettings xlsSettings)
        {
            try
            {
                var workSheets = xlWorkBook.Worksheets.Select(ws => ws.Name);
                var tableHash = new Dictionary<string, IXLWorksheet>();
                foreach (var ws in workSheets)
                {
                    var workSheet = xlWorkBook.Worksheet(ws);
                    tableHash.Add(workSheet.Name, workSheet);
                }
                //インデックスシートを取得
                var wsIndex = xlWorkBook.Worksheet(xlsSettings.IndexSheet);
                var jo = EditJson(
                    sheet: wsIndex,
                    tableHash: tableHash,
                    xlsSettings: xlsSettings);
                return jo;
            }
            catch
            {
                return null;
            };
        }

        private static object EditJson(
            IXLWorksheet sheet,
            Dictionary<string, IXLWorksheet> tableHash,
            XlsSettings xlsSettings,
            string key = null)
        {
            //編集パターン判定
            object jo = new object();
            switch (EditPtnCheck(sheet: sheet))
            {
                case JsonType.Array:
                    jo = JsonArray(
                        sheet: sheet,
                        tableHash: tableHash,
                        xlsSettings: xlsSettings,
                        key: key);
                    break;
                case JsonType.Object:
                    jo = JsonObject(
                        sheet: sheet,
                        tableHash: tableHash,
                        xlsSettings: xlsSettings,
                        key: key);
                    break;
                case JsonType.ArrayObject:
                    jo = JsonArrayObject(
                        sheet: sheet,
                        tableHash: tableHash,
                        xlsSettings: xlsSettings,
                        key: key);
                    break;
                default:
                    break;
            }
            return jo;
        }

        private static List<object> JsonArrayObject(
            IXLWorksheet sheet,
            Dictionary<string, IXLWorksheet> tableHash,
            XlsSettings xlsSettings,
            string key = null)
        {
            //オブジェクト配列を返却
            var jo = new List<object>();
            var list = new List<object>();
            foreach (var tableRow in sheet.Rows())
            {
                //1行目（タイトル行）は対象外
                //keyに値が入ってきた場合は、対象Noの行だけ対象とする
                if (tableRow.RowNumber() > 1
                    && (key == null
                        || key != null
                        && sheet.Cell(tableRow.RowNumber(), 1).Value.ToString() == key))
                {
                    var array = new Dictionary<string, object>();
                    foreach (var tableCell in tableRow.Cells())
                    {
                        if (tableCell.Address.ColumnNumber > 1 &&
                            !string.IsNullOrEmpty(tableCell.Value.ToString()))
                        {
                            if (sheet.Cell(1, tableCell.Address.ColumnNumber).Value.ToString() == "List")
                            {
                                list.Add(
                                    AddCellValue(
                                        tableCell: tableCell,
                                        xlsSettings: xlsSettings,
                                        tableHash: tableHash));
                            }
                            else
                            {
                                array.Add(
                                    sheet.Cell(1, tableCell.Address.ColumnNumber).Value.ToString(),
                                        AddCellValue(
                                            tableCell: tableCell,
                                            xlsSettings: xlsSettings,
                                            tableHash: tableHash));

                            }
                        }
                    }
                    jo.Add(array);
                }
            }
            if (list.Count > 0)
            {
                return list;
            }
            else
            {
                return jo;
            }
        }

        private static object GetChildSheet(
            string value,
            Dictionary<string, IXLWorksheet> tableHash,
            XlsSettings xlsSettings)
        {
            //子要素のシートを取得しjsonを作成する
            var str = value.Replace("{", "").Replace("}", "").Split("_No.");
            var childSheetName = str[0];
            string childSheetNo = null;
            if (str.Count() > 1) { childSheetNo = str[1]; };
            var childSheet = tableHash.Where(tableList => tableList.Key == childSheetName)
                            .Select(tableList => tableList.Value).FirstOrDefault();
            return EditJson(
                sheet: childSheet,
                tableHash: tableHash,
                key: childSheetNo,
                xlsSettings: xlsSettings);
        }

        private static Dictionary<string, object> JsonObject(
            IXLWorksheet sheet,
            Dictionary<string, IXLWorksheet> tableHash,
            XlsSettings xlsSettings,
            string key = null)
        {
            //jsonオブジェクトを返却
            var jo = new Dictionary<string, object>();
            foreach (var tableRow in sheet.Rows())
            {
                if (tableRow.RowNumber() > 1
                    && (key == null
                        || key != null
                        && sheet.Cell(tableRow.RowNumber(), 1).Value.ToString() == key))
                {
                    foreach (var tableCell in tableRow.Cells())
                    {
                        if (tableCell.Address.ColumnNumber > 1
                            && !string.IsNullOrEmpty(tableCell.Value.ToString()))
                        {
                            jo.Add(
                                sheet.Cell(1, tableCell.Address.ColumnNumber).Value.ToString(),
                                AddCellValue(
                                    tableCell: tableCell,
                                    xlsSettings: xlsSettings,
                                    tableHash: tableHash));
                        }
                    }
                    break;
                }
            }
            return jo;
        }

        private static object AddCellValue(
            IXLCell tableCell,
            XlsSettings xlsSettings,
            Dictionary<string, IXLWorksheet> tableHash)
        {
            var cellValue = tableCell.Value.ToString();
            if (cellValue.StartsWith("{") && cellValue.EndsWith("}") && cellValue.Contains("_No"))
            {
                return GetChildSheet(
                        value: cellValue,
                        tableHash: tableHash,
                        xlsSettings: xlsSettings);
            }
            else
            {
                return tableCell.ConvJsonValue();
            }
        }

        private static List<object> JsonArray(
            IXLWorksheet sheet,
            Dictionary<string, IXLWorksheet> tableHash,
            XlsSettings xlsSettings,
            string key = null)
        {
            //配列を返却
            var jo = new List<object>();
            foreach (var tableRow in sheet.Rows())
            {
                if (tableRow.RowNumber() >= 2
                    && (key == null ||
                        key != null && sheet.Cell(tableRow.RowNumber(), 1).Value.ToString() == key))
                {
                    if (sheet.Cell(tableRow.RowNumber(), 2).Value.ToString().StartsWith("{")
                        && sheet.Cell(tableRow.RowNumber(), 2).Value.ToString().EndsWith("}")
                        && sheet.Cell(tableRow.RowNumber(), 2).Value.ToString().Contains("_No"))
                    {
                        jo.Add(GetChildSheet(
                            value: sheet.Cell(tableRow.RowNumber(), 2).Value.ToString(),
                            tableHash: tableHash,
                            xlsSettings: xlsSettings));
                    }
                    else
                    {
                        jo.Add(
                            AddCellValue(
                                tableCell: sheet.Cell(tableRow.RowNumber(), 2),
                                xlsSettings: xlsSettings,
                                tableHash: tableHash));
                    }
                }

            }
            return jo;
        }

        private static JsonType EditPtnCheck(IXLWorksheet sheet)
        {
            //A1セルのメモ内容よりJsonタイプを判断
            //Array or Object
            //Arrayの場合、2列目のタイトルがListの場合、値のみの配列、以外の場合はObject配列と判断
            var list = sheet.Cell(1, 2).Value.ToString();
            if (sheet.Cell(1, 1).GetComment().Text == "Object")
            {
                //オブジェクト
                return JsonType.Object;
            }
            else if (list == "List")
            {
                //2カラム目タイトルが"List"の場合、値のみの配列
                return JsonType.Array;
            }
            else
            {
                //以外の場合オブジェクト配列
                return JsonType.ArrayObject;
            }
        }
    }
}

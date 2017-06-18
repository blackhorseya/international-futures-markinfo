using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace Stock
{
    class ExportExcel
    {
        public void WriteExcel(string filepath, string field)
        {
            try
            {
                using (FileStream fs = new FileStream(filepath, FileMode.OpenOrCreate, FileAccess.Read, FileShare.ReadWrite))
                //FileInfo fi = new FileInfo(filepath);
                {
                    using (ExcelPackage ep = new ExcelPackage(fs))
                    {
                        ep.Save();
                        Process[] ps = Process.GetProcesses();
                        foreach (Process p in ps)
                        {
                            if (p.ProcessName == "EXCEL")
                            {
                                p.Kill();
                                p.StartInfo.FileName = "EXCEL";
                                p.Start();
                            }
                        }
                    }
                    using (ExcelPackage ep = new ExcelPackage(fs))
                    {
                        ExcelWorksheet sheet = ep.Workbook.Worksheets[1];//取得Sheet1
                        //int startRowNumber = sheet.Dimension.Start.Row;//起始列編號，從1算起
                        //int endRowNumber = sheet.Dimension.End.Row;//結束列編號，從1算起
                        //int startColumn = sheet.Dimension.Start.Column;//開始欄編號，從1算起
                        //int endColumn = sheet.Dimension.End.Column;//結束欄編號，從1算起
                        sheet.Cells["A1:H1"].AutoFilter = true;
                        sheet.Cells[field].Value = "test";
                        ep.Save();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void FuturesWriteExcel(List<string> header, List<string> _P, string time)
        {
            try
            {
                string filepath = Application.StartupPath + "/報價/" + DateTime.Now.ToString("yyyy-MM-dd") + "-國際期指.xlsx";
                FileInfo fi = new FileInfo(filepath);
                int startRowNumber, endRowNumber, startColumn, endColumn;

                using (ExcelPackage ep = new ExcelPackage(fi))
                {
                    ExcelWorksheet sheet;
                    try
                    {
                        sheet = ep.Workbook.Worksheets[1];//取得Sheet1
                    }
                    catch
                    {
                        sheet = ep.Workbook.Worksheets.Add(DateTime.Now.ToString("yyyy-MM-dd") + "-國際期指.xlsx");
                    }

                    if (sheet.Dimension != null)
                    {
                        startRowNumber = sheet.Dimension.Start.Row;//起始列編號，從1算起
                        endRowNumber = sheet.Dimension.End.Row;//結束列編號，從1算起
                        startColumn = sheet.Dimension.Start.Column;//開始欄編號，從1算起
                        endColumn = sheet.Dimension.End.Column;//結束欄編號，從1算起

                        sheet.Cells[endRowNumber + 1, 1].Value = time;
                        for (int j = 0; j < header.Count; j++)
                        {
                            bool foundheader = false;
                            for (int i = startColumn; i <= endColumn; i++)
                            {
                                if (header[j] == sheet.Cells[1, i].Text)
                                {
                                    sheet.Cells[endRowNumber + 1, i].Value = _P[j];
                                    foundheader = true;
                                }
                            }
                            if (foundheader == false)
                            {
                                endColumn++;
                                sheet.Cells[1, endColumn].Value = header[j];
                                sheet.Cells[endRowNumber + 1, endColumn].Value = _P[j];
                            }
                        }

                        using (var range = sheet.Cells[startRowNumber + 1, 1, endRowNumber + 1, endColumn])
                        {
                            range.AutoFitColumns();
                            range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                            range.Style.Font.Size = 12;
                        }

                        using (var range = sheet.Cells[1, 1, 1, endColumn])
                        {
                            range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(Color.DimGray);
                            range.Style.Font.Color.SetColor(Color.White);
                            range.Style.Font.Size = 12;
                        }
                    }
                    else
                    {
                        sheet.Cells["A1"].Value = "時間";

                        sheet.Cells["A2"].Value = time;
                        int i = 2;
                        foreach (string str in _P)
                        {
                            sheet.Cells[1, i].Value = header[i-2];
                            sheet.Cells[2, i].Value = str;
                            i++;
                        }

                        using (var range = sheet.Cells[1, 1, 1, i - 1])
                        {
                            range.AutoFitColumns();
                            range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(Color.DimGray);
                            range.Style.Font.Color.SetColor(Color.White);
                            range.Style.Font.Size = 12;
                        }

                        using (var range = sheet.Cells[2, 1, 2, i - 1])
                        {
                            range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                            range.Style.Font.Size = 12;
                        }
                    }

                    ep.Save();
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("檔案開啟中，無法寫入");
                MessageBox.Show(ex.ToString());
            }
        }

        public void FuturesWriteExcel(string tx1, string txel1, string txfi11, string time)
        {
            try
            {
                string filepath = Application.StartupPath + "/報價/" + DateTime.Now.ToString("yyyy-MM-dd") + "-台電金.xlsx";
                FileInfo fi = new FileInfo(filepath);
                int startRowNumber, endRowNumber, startColumn, endColumn;

                using (ExcelPackage ep = new ExcelPackage(fi))
                {
                    ExcelWorksheet sheet;
                    try
                    {
                        sheet = ep.Workbook.Worksheets[1];//取得Sheet1
                    }
                    catch
                    {
                        sheet = ep.Workbook.Worksheets.Add(DateTime.Now.ToString("yyyy-MM-dd") + "-台電金.xlsx");
                    }

                    if (sheet.Dimension != null)
                    {
                        startRowNumber = sheet.Dimension.Start.Row;//起始列編號，從1算起
                        endRowNumber = sheet.Dimension.End.Row;//結束列編號，從1算起
                        startColumn = sheet.Dimension.Start.Column;//開始欄編號，從1算起
                        endColumn = sheet.Dimension.End.Column;//結束欄編號，從1算起

                        sheet.Cells[endRowNumber + 1, 1].Value = time;
                        sheet.Cells[endRowNumber + 1, 2].Value = tx1;
                        sheet.Cells[endRowNumber + 1, 3].Value = txel1;
                        sheet.Cells[endRowNumber + 1, 4].Value = txfi11;

                        using (var range = sheet.Cells[startRowNumber + 1, 1, endRowNumber + 1, endColumn])
                        {
                            range.AutoFitColumns();
                            range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                            range.Style.Font.Size = 12;
                        }
                    }
                    else
                    {
                        sheet.Cells["A1"].Value = "時間";
                        sheet.Cells["B1"].Value = "台指期";
                        sheet.Cells["C1"].Value = "電子期";
                        sheet.Cells["D1"].Value = "金融期";

                        using (var range = sheet.Cells["A1:D1"])
                        {
                            range.AutoFitColumns();
                            range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(Color.DimGray);
                            range.Style.Font.Color.SetColor(Color.White);
                            range.Style.Font.Size = 12;
                        }

                        sheet.Cells["A2"].Value = time;
                        sheet.Cells["B2"].Value = tx1;
                        sheet.Cells["C2"].Value = txel1;
                        sheet.Cells["D2"].Value = txfi11;

                        using (var range = sheet.Cells["A2:D2"])
                        {
                            range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                            range.Style.Font.Size = 12;
                        }
                    }

                    ep.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("檔案開啟中，無法寫入");
            }
        }
    }
}

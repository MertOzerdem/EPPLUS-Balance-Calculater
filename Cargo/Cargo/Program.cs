using System;
using System.Drawing;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;

namespace Cargo
{
    class Program
    {
        static void Main(string[] args)
        {
            double[,] costs = new double[2, 6];
            int[,] ranges = new int[5, 2];
            int[] DistanceArr = new int[5];
            Formula formulas = new Formula(costs, ranges);
            Console.WriteLine("Loading...");

            // FORMUL EXCEL SHEET

            string path = "../FORMUL-GIRDI.xlsx";
            FileInfo fileInfo = new FileInfo(path);
            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

            // get number of rows and columns in the sheet
            int rows = worksheet.Dimension.Rows; // 7
            int columns = worksheet.Dimension.Columns; // 3

            // loop through the worksheet rows and columns
            for (int i = 0; i < worksheet.Dimension.Rows - 1; i++)
            {
                string temp = worksheet.Cells[i + 2, 1].Value.ToString();
                if (temp != "Artan Her Desi İçin")
                {
                    string[] tempArr = temp.Split('-');
                    formulas.desiRange[i, 0] = Int32.Parse(tempArr[0]);
                    formulas.desiRange[i, 1] = Int32.Parse(tempArr[1]);
                }

                for (int j = 0; j < worksheet.Dimension.Columns - 1; j++)
                {
                    formulas.cost[j, i] = Convert.ToDouble(worksheet.Cells[i + 2, j + 2].Value.ToString());
                }
            }
            package.Dispose();
            // EKSTRE EXCEL SHEET

            string ekstrePath = "../EKSTRE-GIRDI.xlsx";
            FileInfo ekstreFileInfo = new FileInfo(ekstrePath);
            ExcelPackage ekstrePackage = new ExcelPackage(ekstreFileInfo);
            ExcelWorksheet ekstreWorksheet = ekstrePackage.Workbook.Worksheets.FirstOrDefault();

            Console.WriteLine("Loading More...");
            int cargoRowCount = ekstreWorksheet.Dimension.Rows;
            double[] results = new double[cargoRowCount - 1];
            
            for (var i = 0; i < cargoRowCount - 1; i++)
            {
                // ücretleri hesaplamak için
                string[] cells = new string[2];
                //var cells = ekstreWorksheet[$"C{i + 2}:D{i + 2}"].ToList();
                cells[0] = ekstreWorksheet.Cells[i + 2, 3].Value.ToString();
                cells[1] = ekstreWorksheet.Cells[i + 2, 4].Value.ToString();
                if ("UZAK" == cells[1] || "ORTA" == cells[1])
                {
                    int desiWeight = Int32.Parse(cells[0]);
                    int desiIndex = desiCalculator(desiWeight, ranges);
                    if (desiIndex == 5)
                    {
                        results[i] = costs[1, 4] + ((desiWeight - ranges[4, 1]) * costs[1, 5]);
                    }
                    else
                    {
                        results[i] = costs[1, desiIndex];
                    }
                }
                else
                {
                    int desiWeight = Int32.Parse(cells[0]);
                    int desiIndex = desiCalculator(desiWeight, ranges);
                    if (desiIndex == 5)
                    {
                        results[i] = costs[0, 4] + ((desiWeight - ranges[4, 1]) * costs[0, 5]);
                    }
                    else
                    {
                        results[i] = costs[0, desiIndex];
                    }
                }

                // total Mesafelari hesaplamak için
                var tempDistance = ekstreWorksheet.Cells[i + 2, 4].Value.ToString();
                DistanceArr[distanceCheck(tempDistance)] = DistanceArr[distanceCheck(tempDistance)] 
                                    + Int32.Parse(ekstreWorksheet.Cells[i + 2, 2].Value.ToString());
            }

            int grandTotal = 0;
            foreach (var distance in DistanceArr)
            {
                grandTotal = grandTotal + distance;
            }


            ExcelPackage rapor = new ExcelPackage();

            // name of the sheet 
            var firstSheet = rapor.Workbook.Worksheets.Add("İşlem Sonucu", ekstreWorksheet);
            var secondSheet = rapor.Workbook.Worksheets.Add("Pivot Rapor");
            ekstrePackage.Dispose(); // dispose the ekstre sheet

            // FILL FIRST WORKSHEET
            firstSheet.Cells[1, 5].Value = "ÜCRET";
            for (int i = 0; i < cargoRowCount - 1; i++)
            {
                firstSheet.Cells[i+2,5].Value = results[i];
            }
            // style the first worksheet
            firstSheet.Column(1).AutoFit();
            firstSheet.Column(2).AutoFit();
            firstSheet.Column(3).AutoFit();
            firstSheet.Column(4).AutoFit();
            firstSheet.Column(5).AutoFit();

            
            // Add data to secondSheet
            secondSheet.Cells[2,2].Value = "Mesafe";
            secondSheet.Cells[2,3].Value = "Kargo Adeti";
            secondSheet.Cells[2, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            secondSheet.Cells[2, 2].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            secondSheet.Cells[2, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            secondSheet.Cells[2, 3].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            secondSheet.Cells[2, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            secondSheet.Cells[2, 3].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            secondSheet.Cells[2, 2].Style.Font.Bold = true;
            secondSheet.Cells[2, 3].Style.Font.Bold = true;

            secondSheet.Cells[3,2].Value = "UZAK";
            secondSheet.Cells[3,3].Value = DistanceArr[0];
            secondSheet.Cells[3, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            secondSheet.Cells[3, 2].Style.Fill.BackgroundColor.SetColor(Color.White);
            secondSheet.Cells[3, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            secondSheet.Cells[3, 3].Style.Fill.BackgroundColor.SetColor(Color.White);

            secondSheet.Cells[4,2].Value = "ORTA";
            secondSheet.Cells[4,3].Value = DistanceArr[1];
            secondSheet.Cells[4, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            secondSheet.Cells[4, 2].Style.Fill.BackgroundColor.SetColor(Color.White);
            secondSheet.Cells[4, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            secondSheet.Cells[4, 3].Style.Fill.BackgroundColor.SetColor(Color.White);

            secondSheet.Cells[5,2].Value = "KISA";
            secondSheet.Cells[5,3].Value = DistanceArr[2];
            secondSheet.Cells[5, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            secondSheet.Cells[5, 2].Style.Fill.BackgroundColor.SetColor(Color.White);
            secondSheet.Cells[5, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            secondSheet.Cells[5, 3].Style.Fill.BackgroundColor.SetColor(Color.White);

            secondSheet.Cells[6, 2].Value = "YAKIN";
            secondSheet.Cells[6, 3].Value = DistanceArr[3];
            secondSheet.Cells[6, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            secondSheet.Cells[6, 2].Style.Fill.BackgroundColor.SetColor(Color.White);
            secondSheet.Cells[6, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            secondSheet.Cells[6, 3].Style.Fill.BackgroundColor.SetColor(Color.White);

            secondSheet.Cells[7,2].Value = "ŞEHİRİÇİ";
            secondSheet.Cells[7,3].Value = DistanceArr[4];
            secondSheet.Cells[7, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            secondSheet.Cells[7, 2].Style.Fill.BackgroundColor.SetColor(Color.White);
            secondSheet.Cells[7, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            secondSheet.Cells[7, 3].Style.Fill.BackgroundColor.SetColor(Color.White);

            secondSheet.Cells[8,2].Value = "Grand Total";
            secondSheet.Cells[8,3].Value = grandTotal;
            secondSheet.Cells[8, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            secondSheet.Cells[8, 2].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            secondSheet.Cells[8, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            secondSheet.Cells[8, 3].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            secondSheet.Cells[8, 2].Style.Border.Top.Style = ExcelBorderStyle.Thick;
            secondSheet.Cells[8, 3].Style.Border.Top.Style = ExcelBorderStyle.Thick;
            secondSheet.Cells[8, 2].Style.Font.Bold = true;
            secondSheet.Cells[8, 3].Style.Font.Bold = true;

            // fit the sheet
            secondSheet.Column(2).AutoFit();
            secondSheet.Column(3).AutoFit();

            // push to file
            string p_strPath = "../MERTRAPOR.xlsx";

            if (File.Exists(p_strPath))
                File.Delete(p_strPath);

            // Create excel file on physical disk  
            FileStream objFileStrm = File.Create(p_strPath);
            objFileStrm.Close();

            // Write content to excel file  
            File.WriteAllBytes(p_strPath, rapor.GetAsByteArray());
            //Close package 
            rapor.Dispose();
        }

        public enum Distances
        {
            UZAK = 0,
            ORTA = 1,
            KISA = 2,
            YAKIN = 3,
            ŞEHİRİÇİ = 4
        }

        private static int distanceCheck(string distance)
        {

            switch (distance)
            {
                case "UZAK":
                    return (int)Distances.UZAK;
                case "ORTA":
                    return (int)Distances.ORTA;
                case "KISA":
                    return (int)Distances.KISA;
                case "YAKIN":
                    return (int)Distances.YAKIN;
                default:
                    return (int)Distances.ŞEHİRİÇİ;
            }
        }

        private static int desiCalculator(int desiWeight, int[,] ranges)
        {

            for (int i = 0; i < ranges.Length / 2; i++)
            {
                if (ranges[i, 0] <= desiWeight && ranges[i, 1] >= desiWeight)
                {
                    return i;
                }
            }
            return (ranges.Length / 2);
        }

        public class Formula
        {
            public double[,] cost { get; set; }
            public int[,] desiRange { get; set; }

            public Formula(double[,] arr, int[,] arr2)
            {
                this.cost = arr;
                this.desiRange = arr2;
            }

        }
    }
}

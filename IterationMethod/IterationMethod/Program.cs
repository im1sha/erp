using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IterationMethod
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            const int TOTAL_ADDITIONAL_UNITS = 4;
            const int TOTAL_MAIN_UNITS = 2;

            static int TotalUnits()
            {
                return TOTAL_ADDITIONAL_UNITS + TOTAL_MAIN_UNITS;
            }

            decimal[] primaryExpenses = new[]
            {
                58m,
                33m,
                60m,
                40m,
            };

            decimal[][] volumes = new decimal[][]
            {
                new decimal[] { 0m,   110m, 55m, 60m,  450m, 350m },
                new decimal[] { 360m, 0m,   0m,  200m, 300m, 800m },
                new decimal[] { 200m, 58m,  0m,  100m, 120m, 100m },
                new decimal[] { 0m,   0m,   0m,  0m,   87m,  40m },
            };

            //decimal[] primaryExpenses = new[]
            //{
            //    20m,
            //    80m,
            //    50m,
            //};

            //decimal[][] volumes = new decimal[][]
            //{
            //    new decimal[] { 0m,   110m, 60, 100,  400,  },
            //    new decimal[] { 200, 0m,   400,  1200, 200m,  },
            //    new decimal[] { 50, 50,  0m,  300, 100,  },
            //};

            decimal[][] structure = new decimal[TOTAL_ADDITIONAL_UNITS][];
            for (int i = 0; i < TOTAL_ADDITIONAL_UNITS; i++)
            {
                structure[i] = new decimal[TotalUnits()];
            }

            for (int i = 0; i < structure.Length; i++)
            {
                for (int j = 0; j < structure[i].Length; j++)
                {
                    structure[i][j] = volumes[i][j] / volumes[i].Sum();
                }
            }

            decimal maxDelta = decimal.MaxValue;
            decimal epsilon = 0.0000000000000000000001m;

            decimal[] xOverallExpensess = primaryExpenses.Select(i => i).ToArray();

            var xLog = new List<decimal[]>();
            var deltasLog = new List<decimal[]>();
            var maxDeltaLog = new List<decimal>();

            while (maxDelta > epsilon)
            {
                xLog.Add(xOverallExpensess.Select(i => i).ToArray());

                for (int i = 0; i < xOverallExpensess.Length; i++)
                {
                    xOverallExpensess[i] = primaryExpenses[i] +
                        Enumerable.Range(0, xOverallExpensess.Length)
                        .Except(new[] { i })
                        .Select(j => xOverallExpensess[j] * structure[i][j])
                        .Sum();
                }

                var newDeltas = new decimal[TOTAL_ADDITIONAL_UNITS];
                xLog.Last().Aggregate(0, (index, next) =>
                {
                    newDeltas[index] = Math.Abs(next - xOverallExpensess[index]);
                    return ++index;
                });
                deltasLog.Add(newDeltas);

                maxDelta = newDeltas.Max();
                maxDeltaLog.Add(maxDelta);
            }

            var tariffs = Enumerable.Range(0, TOTAL_ADDITIONAL_UNITS)
                .Select(i => xOverallExpensess[i] / volumes[i].Sum()).ToArray();

            var result = volumes.Select(i => i.Select(j => j).ToArray()).ToArray();
            for (int i = 0; i < result.Length; i++)
            {
                for (int j = 0; j < result[i].Length; j++)
                {
                    result[i][j] *= tariffs[i];
                }
            }

            using FileStream stream = File.Create("lab7.xlsx");
            using var workbook = new XLWorkbook();           

            const int horizontalShift = 2;    
            const int deltaVerticalShift = 7;
  
            var verticalShift = 1;

            var worksheet = workbook.Worksheets.Add("Data");
            for (int i = 0; i < TOTAL_ADDITIONAL_UNITS; i++)
            {
                worksheet.Cell(verticalShift, i + horizontalShift).Value 
                    = $"Item {1+i}";
                worksheet.Cell(1 + verticalShift, i + horizontalShift).Value
                    = primaryExpenses[i].ToString();
            }

            verticalShift += deltaVerticalShift;
            Print(worksheet, Enumerable.Range(1, TOTAL_ADDITIONAL_UNITS).Select(i => $"Item{i}").ToArray(), volumes, horizontalShift, verticalShift, new[] { "Volumes" });

            verticalShift += deltaVerticalShift;
            Print(worksheet, Enumerable.Range(1, TOTAL_ADDITIONAL_UNITS).Select(i => $"Item{i}").ToArray(), structure, horizontalShift, verticalShift,new[] { "Structure" });

            verticalShift += deltaVerticalShift;

            for (int i = 0; i < xLog.Count; i++)
            {
                Print(worksheet, new[] { $"#{i}"}, new decimal[][] { xLog[i].Concat(deltasLog[i]).Concat(new[] { maxDeltaLog[i] }).ToArray() }, horizontalShift, verticalShift, null);
                ++verticalShift;
            }

            verticalShift += deltaVerticalShift;
            Print(worksheet, Enumerable.Range(1, TOTAL_ADDITIONAL_UNITS).Select(i => $"Item{i}").ToArray(), result, horizontalShift, verticalShift, new[] { "Result" });

            workbook.SaveAs(stream);
        }

        static void Print(IXLWorksheet worksheet, string[] headers, decimal[][] values, int horizontalShift, int verticalShift, string[] mainHeader)
        {
            if (mainHeader != null)
            {
                for (int i = 0; i < mainHeader.Length; i++)
                {
                    worksheet.Cell(verticalShift - 1, horizontalShift - 1 + i).Value = mainHeader[i];
                }
            }

            for (int i = 0; i < values.Length; i++)
            {
                worksheet.Cell(i + verticalShift, horizontalShift - 1).Value
                           = headers[i];

                for (int j = 0; j < values[i].Length; j++)
                {
                    worksheet.Cell(i + verticalShift, j + horizontalShift).Value
                        = values[i][j].ToString();
                }
            }
        }
    }
}

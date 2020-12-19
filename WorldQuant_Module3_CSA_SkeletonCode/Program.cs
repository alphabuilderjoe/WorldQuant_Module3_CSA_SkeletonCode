using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace WorldQuant_Module3_CSA_SkeletonCode
{
    class Program
    {
        static Excel.Workbook workbook;
        static Excel.Application app;

        static void Main(string[] args)
        {
            app = new Excel.Application();
            app.Visible = true;
            try
            {
                workbook = app.Workbooks.Open("property_pricing.xlsx", ReadOnly: false);
            }
            catch
            {
                SetUp();
            }

            var input = "";
            while (input != "x")
            {
                PrintMenu();
                input = Console.ReadLine();
                try
                {
                    var option = int.Parse(input);
                    switch (option)
                    {
                        case 1:
                            try
                            {
                                Console.Write("Enter the size: ");
                                var size = float.Parse(Console.ReadLine());
                                Console.Write("Enter the suburb: ");
                                var suburb = Console.ReadLine();
                                Console.Write("Enter the city: ");
                                var city = Console.ReadLine();
                                Console.Write("Enter the market value: ");
                                var value = float.Parse(Console.ReadLine());

                                AddPropertyToWorksheet(size, suburb, city, value);
                            }
                            catch
                            {
                                Console.WriteLine("Error: couldn't parse input");
                            }
                            break;
                        case 2:
                            Console.WriteLine("Mean price: " + CalculateMean());
                            break;
                        case 3:
                            Console.WriteLine("Price variance: " + CalculateVariance());
                            break;
                        case 4:
                            Console.WriteLine("Minimum price: " + CalculateMinimum());
                            break;
                        case 5:
                            Console.WriteLine("Maximum price: " + CalculateMaximum());
                            break;
                        default:
                            break;
                    }
                } catch { }
            }

            // save before exiting
            workbook.Save();
            workbook.Close();
            app.Quit();
        }

        static void PrintMenu()
        {
            Console.WriteLine();
            Console.WriteLine("Select an option (1, 2, 3, 4, 5) " +
                              "or enter 'x' to quit...");
            Console.WriteLine("1: Add Property");
            Console.WriteLine("2: Calculate Mean");
            Console.WriteLine("3: Calculate Variance");
            Console.WriteLine("4: Calculate Minimum");
            Console.WriteLine("5: Calculate Maximum");
            Console.WriteLine();
        }

        static void SetUp()
        {
            // TODO: Implement this method
            app.Workbooks.Add();
            workbook = app.ActiveWorkbook;
            
            //Worksheet which stores property data
            workbook.Worksheets.Add();
            Excel.Worksheet currentSheet = workbook.Worksheets[1];
            currentSheet.Name = "Property";

            workbook.SaveAs("property_pricing.xlsx");
        }

        static void AddPropertyToWorksheet(float size, string suburb, string city, float value)
        {
            int row = 1;

            Excel.Worksheet currentSheet = workbook.Worksheets[1];

            while(true)
            {
                if(currentSheet.Cells[row, "A"].Value == null)
                {
                    currentSheet.Cells[row, "A"] = size;
                    currentSheet.Cells[row, "B"] = suburb;
                    currentSheet.Cells[row, "C"] = city;
                    currentSheet.Cells[row, "D"] = value;
                    currentSheet.Cells[row, "E"] = row;
                    return;

                }
                row++;
            }

        }

        static float CalculateMean()
        {
            int row = 1;
            float sum = 0.0f;

            Excel.Worksheet currentSheet = workbook.Worksheets[1];

            while (currentSheet.Cells[row, "D"].Value != null)
            {

                sum += currentSheet.Cells[row, "D"].Value;
                row++;
            }

            
            return sum/(row-1);

        }

        static float CalculateVariance()
        {
            float mean = CalculateMean();
            
            int row = 1;
            float sum_sqred_diff = 0.0f;

            Excel.Worksheet currentSheet = workbook.Worksheets[1];

            while (currentSheet.Cells[row, "D"].Value != null)
            {

                sum_sqred_diff += (currentSheet.Cells[row, "D"].Value - mean) * (currentSheet.Cells[row, "D"].Value - mean);
                row++;
            }

            return sum_sqred_diff/(row-2); //Using variance formula for a sample
        }

        static float CalculateMinimum()
        {
            Excel.Worksheet currentSheet = workbook.Worksheets[1];


            int row = 2;
            float min_value = 0.0f;
            min_value = (float)currentSheet.Cells[1, "D"].Value;


            while (currentSheet.Cells[row, "D"].Value != null)
            {

                min_value = Math.Min(min_value, (float)currentSheet.Cells[row, "D"].Value);

                row++;
            }
            
            return min_value;


        }

        static float CalculateMaximum()
        {
            Excel.Worksheet currentSheet = workbook.Worksheets[1];

            int row = 2;
            float max = (float)currentSheet.Cells[1, "D"].Value;

            while (currentSheet.Cells[row, "D"].Value != null)
            {

                max = Math.Max(max, (float)currentSheet.Cells[row, "D"].Value);
                row++;
            }

            return max;
        }
    }
}

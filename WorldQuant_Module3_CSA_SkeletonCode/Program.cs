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
            // Create new Excel application, and try to open property_pricing.xlsx spreadsheet. 
            // If it doesn't exist, we will run SetUp() function to create a new spreadsheet

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

            // Use a While loop to allow user to enter new property data or to run queries
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
            // SetUp is called if property_prices.xlsx spreadsheet is not detected, thus we need to create and save a new spreadsheet.
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
            // This function is called when a new property is added to the spreadsheet. 

            int row = 1;

            Excel.Worksheet currentSheet = workbook.Worksheets[1];

            // We search the spreadsheet, going down row by row to find the first row which has a blank/null value. Once found, we add all the property data to that row

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
            // Calculate the mean of all property prices

            int row = 1;
            float sum = 0.0f;

            Excel.Worksheet currentSheet = workbook.Worksheets[1];

            // We go down the spreadsheet row-by-row, continuing to add the property price of each row into the variable sum. 

            while (currentSheet.Cells[row, "D"].Value != null)
            {

                sum += currentSheet.Cells[row, "D"].Value;
                row++;
            }

            // We calculate the mean by dividing the sum by the value of row-1. We deduct 1 because the row variable count stops at the first row with a null value. 
            return sum/(row-1);

        }

        static float CalculateVariance()
        {
            // Calculate the variance of all property prices

            // We obtain the mean of all property prices first using the CalculateMean() function
            float mean = CalculateMean();
            
            int row = 1;
            float sum_sqred_diff = 0.0f;

            Excel.Worksheet currentSheet = workbook.Worksheets[1];

            // We go down the spreadsheet row-by-row, continuing to add the squared difference between the property price and mean price of each row into the variable sum-sqred_diff. 
            while (currentSheet.Cells[row, "D"].Value != null)
            {

                sum_sqred_diff += (currentSheet.Cells[row, "D"].Value - mean) * (currentSheet.Cells[row, "D"].Value - mean);
                row++;
            }

            // We calculate the mean by dividing the sum of squared differences by the value of row-2. The formula we use is for the variance of a sample, which is the number of itmes minus 1.
            // We deduct 2 instead of just 1 because the row variable count stops at the first row with a null value. 
            return sum_sqred_diff /(row-2); 
        }

        static float CalculateMinimum()
        {

            // Calculate the minimum of all property prices

            Excel.Worksheet currentSheet = workbook.Worksheets[1];

            // We start our row variable from 2, as will begin our loop assuming the min_value is the first row. 
            int row = 2;
            float min_value = 0.0f;
            min_value = (float)currentSheet.Cells[1, "D"].Value;

            // We then compare the next row with our minimum, and will update our min value if the new row has a smaller value.
            while (currentSheet.Cells[row, "D"].Value != null)
            {

                min_value = Math.Min(min_value, (float)currentSheet.Cells[row, "D"].Value);

                row++;
            }
            
            return min_value;


        }

        static float CalculateMaximum()
        {
            // Calculate the maximum of all property prices

            Excel.Worksheet currentSheet = workbook.Worksheets[1];

            // We start our row variable from 2, as will begin our loop assuming the max_value is the first row. 
            int row = 2;
            float max_value = (float)currentSheet.Cells[1, "D"].Value;

            // We then compare the next row with our max_value, and will update our max_value if the new row has a larger value.
            while (currentSheet.Cells[row, "D"].Value != null)
            {

                max_value = Math.Max(max_value, (float)currentSheet.Cells[row, "D"].Value);
                row++;
            }

            return max_value;
        }
    }
}

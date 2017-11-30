using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;

namespace MyTeamApp
{
    class MyExcel
    {
        public static string DB_PATH = @"";
        public static BindingList<Employee> EmpList = new BindingList<Employee>();
        public static double[,] tablo = new double[2, 6];
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        private static int lastRow=0;
        public static void InitializeExcel()
        {
            MyApp = new Excel.Application();
            MyApp.Visible = false;
           // MyExcel.copyFile();
            MyBook = MyApp.Workbooks.Open(DB_PATH);
            MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explict cast is not required here
            lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            //((Excel.Range) MySheet.Rows[2, System.Reflection.Missing.Value]).Delete(XlDeleteShiftDirection.xlShiftUp);
            //MyBook.Save();
            //lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            calculateTable();
        }
        public static BindingList<Employee> ReadMyExcel()
        {
            EmpList.Clear();
            for (int index = 2; index <= lastRow; index++)
            {
                System.Array MyValues = (System.Array)MySheet.get_Range("A" + index.ToString(), "D" + index.ToString()).Cells.Value;
                EmpList.Add(new Employee { 
                                            Name = MyValues.GetValue(1,1).ToString(),
                                            Employee_ID = MyValues.GetValue(1,2).ToString(),
                                            Email_ID = MyValues.GetValue(1,3).ToString(),
                                            Number = MyValues.GetValue(1,4).ToString()
                                        });
            }
            return EmpList;
        }
        public static void WriteToExcel(Employee emp)
        {
            try
            {
                lastRow += 1;
                MySheet.Cells[lastRow, 1] = emp.Name;
                MySheet.Cells[lastRow, 2] = emp.Employee_ID;
                MySheet.Cells[lastRow, 3] = emp.Email_ID;
                MySheet.Cells[lastRow, 4] = emp.Number;
                EmpList.Add(emp);
                MyBook.Save();
            }
            catch (Exception ex)
            { }

        }

        public static List<Employee> FilterEmpList(string searchValue, string searchExpr)
        {
            List<Employee> FilteredList = new List<Employee>();
            switch (searchValue.ToUpper())
            {
                case "NAME":
                    FilteredList = EmpList.ToList().FindAll(emp => emp.Name.ToLower().Contains(searchExpr));
                    break;
                case "MOBILE_NO":
                    FilteredList = EmpList.ToList().FindAll(emp => emp.Number.ToLower().Contains(searchExpr));
                    break;
                case "EMPLOYEE_ID":
                    FilteredList = EmpList.ToList().FindAll(emp => emp.Employee_ID.ToLower().Contains(searchExpr));
                    break;
                case "EMAIL_ID":
                    FilteredList = EmpList.ToList().FindAll(emp => emp.Email_ID.ToLower().Contains(searchExpr));
                    break;
                default:
                    break;
            }
            return FilteredList;
        }
        public static void CloseExcel()
        {
            MyBook.Saved = true;
            MyApp.Quit();

        }


        public static void calculateTable()
        {
            // Calculate where Risk is "1.Seviye" and Sonuc is Evet
            string[] range1 = { "A2:A" + lastRow, "D2:D" + lastRow };
            string[] criteria1 = { "1.Seviye", "Evet" };
            tablo[0, 0] = count(range1, criteria1);



            // Calculate where Risk is "1.Seviye" and Sonuc is Evet
            string[] range2 = { "A2:A" + lastRow, "A2:A" + lastRow, "D2:D" + lastRow };
            string[] criteria2 = { "2.Seviye", "3.Seviye", "Evet" };
            tablo[0, 1] = count(range2, criteria2);
            
            
            // Calculate where Risk is "1.Seviye" and Sonuc is Evet
            string[] range3 = { "B2:B" + lastRow, "D2:D" + lastRow };
            string[] criteria3 = { "İyi", "Evet" };
            tablo[0, 2] = count(range3, criteria3);



            // Calculate where Risk is "1.Seviye" and Sonuc is Evet
            string[] range4 = { "B2:B" + lastRow, "B2:B" + lastRow, "D2:D" + lastRow };
            string[] criteria4 = { "Orta", "Kötü", "Evet" };
            tablo[0, 3] = count(range4, criteria4);


            // Calculate where Risk is "1.Seviye" and Sonuc is Evet
            string[] range5 = { "C2:C" + lastRow, "D2:D" + lastRow };
            string[] criteria5 = { "Erkek", "Evet" };
            tablo[0, 4] = count(range5, criteria5);



            // Calculate where Risk is "1.Seviye" and Sonuc is Evet
            string[] range6 = { "C2:C" + lastRow, "D2:D" + lastRow };
            string[] criteria6 = { "Kız", "Evet" };
            tablo[0, 5] = count(range6, criteria6);


            // Calculate where Risk is "1.Seviye" and Sonuc is Evet
            string[] range11 = { "A2:A" + lastRow, "D2:D" + lastRow };
            string[] criteria11 = { "1.Seviye", "Hayır" };
            tablo[0, 0] = count(range11, criteria11);



            // Calculate where Risk is "1.Seviye" and Sonuc is Evet
            string[] range22 = { "A2:A" + lastRow, "A2:A" + lastRow, "D2:D" + lastRow };
            string[] criteria22 = { "2.Seviye", "3.Seviye", "Hayır" };
            tablo[0, 1] = count(range22, criteria22);


            // Calculate where Risk is "1.Seviye" and Sonuc is Evet
            string[] range33 = { "B2:B" + lastRow, "D2:D" + lastRow };
            string[] criteria33 = { "İyi", "Hayır" };
            tablo[0, 2] = count(range33, criteria33);



            // Calculate where Risk is "1.Seviye" and Sonuc is Evet
            string[] range44 = { "B2:B" + lastRow, "B2:B" + lastRow, "D2:D" + lastRow };
            string[] criteria44 = { "Orta", "Kötü", "Hayır" };
            tablo[0, 3] = count(range44, criteria44);


            // Calculate where Risk is "1.Seviye" and Sonuc is Evet
            string[] range55= { "C2:C" + lastRow, "D2:D" + lastRow };
            string[] criteria55 = { "Erkek", "Hayır" };
            tablo[0, 4] = count(range55, criteria55);



            // Calculate where Risk is "1.Seviye" and Sonuc is Evet
            string[] range66 = { "C2:C" + lastRow, "D2:D" + lastRow };
            string[] criteria66 = { "Kız", "Hayır" };
            tablo[0, 5] = count(range66, criteria66);
        }


        public static double count(string[] range , string[] criteria)
        {
            int numOfParameters = range.Length;
            string formula = "= COUNTIFS(";
            for (int i=0; i<numOfParameters; i++)
            {
                formula += range[i] + ", \"" + criteria[i] + "\" ,"; 
            }
            formula = formula.Substring(0, formula.Length - 1);
            formula += ")";
            MySheet.Range["Z2"].Formula = formula;
            double MyValue = MySheet.get_Range("Z" + 2.ToString(), "Z" + 2.ToString()).Cells.Value;
            return MyValue;
        }



        public static bool copyFile()
        {
            //MySheet.Cells[1, 1].Formula = "=SUM(A2,B2)";

            int length = DB_PATH.Length;
            string before2 = DB_PATH.Substring(0, length - 5) + "2";
            string after2 = DB_PATH.Substring(length - 5, 5);
            string destination = before2 + after2;
            System.IO.File.Copy(DB_PATH,destination,true);
            MyApp.Quit();
            DB_PATH = destination;
            return true;
        }

    }
    
}

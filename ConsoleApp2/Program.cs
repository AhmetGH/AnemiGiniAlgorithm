using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.IO;

namespace ExcelConverter
{

    public class Dugum
    {
        public List<List<float>> dugumValue { get; set; }
        public List<List<float>> leftDugumValue { get; set; }
        public List<List<float>> rightDugumValue { get; set; }
        public List<string> resultList { get; set; }
        public List<string> leftResultList { get; set; }
        public List<string> rightResultList { get; set; }
        public Dugum()
        {
            dugumValue = new List<List<float>>();
            leftDugumValue = new List<List<float>>();
            rightDugumValue = new List<List<float>>();
            resultList = new List<string>();
            leftResultList = new List<string>();
            rightResultList = new List<string>();
        }
        public Dugum(List<List<float>> dugumValue, List<List<float>> leftDugumValue, List<List<float>> rightDugumValue, List<string> resultList, List<string> leftResultList, List<string> rightResultList) {
            this.dugumValue = dugumValue;
            this.leftDugumValue = leftDugumValue;
            this.rightDugumValue = rightDugumValue;
            this.resultList = resultList;
            this.leftResultList = leftResultList;
            this.rightResultList = rightResultList;
        }
        public Dugum(List<List<float>> leftDugumValue, List<string> leftResultList)
        {
            dugumValue= leftDugumValue;
            resultList= leftResultList;
        }
        public Dugum(List<string> rightResultList,List<List<float>> rightDugumValue)
        {
            dugumValue = rightDugumValue;
            resultList = rightResultList;
        }

    }
    class Program
    {
        [DllImport("User32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int ProcessId);
        private static void KillExcel(Excel.Application theApp)
        {
            int id = 0;
            IntPtr intptr = new IntPtr(theApp.Hwnd);
            System.Diagnostics.Process p = null;
            try
            {
                GetWindowThreadProcessId(intptr, out id);
                p = System.Diagnostics.Process.GetProcessById(id);
                if (p != null)
                {
                    p.Kill();
                    p.Dispose();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Hata");
            }
        }
        static float aOrtalama(List<float> columnValues)
        {
            float total = 0;
            float ortalama;
            // Elde edilen float listesini kullanabilirsiniz
            foreach (float rbcValue in columnValues)
            {
                //Console.WriteLine(rbcValue);
                total += rbcValue;
            }
            ortalama = total / columnValues.Count;
            return ortalama;
        }

        static List<string> resultColumn()
        {
            string sonuc = "SONUÇ";
            List<string> resultValues = new List<string>();

            //Excel başlatılır.
            Excel.Application excel = new Excel.Application();

            try
            {
            excel.Visible = false;
            //Excel dosyası açılır.
            Excel.Workbook workbook = excel.Workbooks.Open(@"C:\Users\ahmet\source\repos\ConsoleApp2\anemi.xlsx");
            //Çalışma sayfası seçilir.
            Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[1];
            //Excel.Range range = sheet.get_Range("B1:B" + sheet.UsedRange.Rows.Count);

            int resultColumnNumber = -1; // Varsayılan olarak -1



            // Sütun arama yapılır
            Excel.Range headerRow = sheet.get_Range("A1:Z1"); // İlk satırda sütun başlıkları

            foreach (Excel.Range cell in headerRow)
            {
                if (cell.Value2 != null && cell.Value2.ToString() == sonuc)
                {
                    resultColumnNumber = cell.Column;
                    break;
                }
            }

            
            // "RBC" adında bir sütun bulundu mu?
            if (resultColumnNumber != -1)
            {
                // "RBC" sütununun tamamını alın
                Excel.Range resultColumn = sheet.get_Range(sheet.Cells[1, resultColumnNumber], sheet.Cells[sheet.UsedRange.Rows.Count, resultColumnNumber]);

                // "RBC" sütunundaki tüm değerleri bir float listesine atın


                foreach (Excel.Range cell in resultColumn)
                {
                    resultValues.Add(cell.Value2.ToString());
                }

                // Elde edilen string listesini kullanabilirsiniz
                /*foreach (string resultValue in resultValues)
                {
                    Console.WriteLine(resultValue);
                }*/
            }
            else
            {
                Console.WriteLine("SONUÇ sütunu bulunamadı.");
            }
            resultValues.RemoveAt(0);

            }
            catch (Exception ex) { Console.WriteLine("Hata"); }
            finally
            {
                KillExcel(excel);
                //System.Threading.Thread.Sleep(100);
            }
            return resultValues;
        }
        static List<float> categoryColumn(int category)
        {
            string[] kategori = { "RBC", "HGB", "HCT", "MCV", "MCH", "MCHC" };
            int ColumnNumber = -1; // Varsayılan olarak -1
            List<float> columnValues = new List<float>();
            Excel.Application excel = new Excel.Application();
            try {
                excel.Visible = false;
                //Excel dosyası açılır.
                Excel.Workbook workbook = excel.Workbooks.Open(@"C:\Users\ahmet\source\repos\ConsoleApp2\anemi.xlsx");
                //Çalışma sayfası seçilir.
                Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[1];
                Excel.Range headerRow = sheet.get_Range("A1:Z1");



                //Sütun değerlerini floata çevirme

                foreach (Excel.Range cell in headerRow)
                {
                    if (cell.Value2 != null && cell.Value2.ToString() == kategori[category])
                    {
                        ColumnNumber = cell.Column;
                        break;
                    }
                }
                //Sütun arama



                //Sütun değerlerini floata çevirme
                
                // "RBC" adında bir sütun bulundu mu?
                if (ColumnNumber != -1)
                {
                    // "RBC" sütununun tamamını alın
                    Excel.Range rbcColumn = sheet.get_Range(sheet.Cells[1, ColumnNumber], sheet.Cells[sheet.UsedRange.Rows.Count, ColumnNumber]);

                    // "RBC" sütunundaki tüm değerleri bir float listesine atın

                    foreach (Excel.Range cell in rbcColumn)
                    {
                        float value;
                        if (float.TryParse(cell.Value2.ToString(), out value))
                        {
                            columnValues.Add(value);
                        }
                    }
                    /*foreach (float columnValue in columnValues)
                    {
                        Console.WriteLine(columnValue);
                    }*/
                }
                else
                {
                    Console.WriteLine("sütun bulunamadı.");
                }
            }catch (Exception ex) { Console.WriteLine("Hata"); }
            finally
            {
                KillExcel(excel);
                System.Threading.Thread.Sleep(100);
            }
            return columnValues;
        }

        static float giniCalculator(List<float> columnValues,List<string> resultValues)
        {
            int upYesCount = 0, upNoCount = 0, downYesCount = 0, downNoCount = 0;
            float aMean = aOrtalama(columnValues);
            for (int i = 0; resultValues.Count > i; i++)
            {
                if (columnValues[i] >= aMean)
                {
                    if (resultValues[i] == "yes") { upYesCount++; } else { upNoCount++; }
                }
                else
                {
                    if (resultValues[i] == "yes") { downYesCount++; } else { downNoCount++; }
                }
            }
            Console.WriteLine(upYesCount + " " + upNoCount + " " + downYesCount + " " + downNoCount + " ");

            double leftGiniDouble, rightGiniDouble, giniDouble;

            float upYesRatio = (float)upYesCount / (upYesCount + upNoCount);
            float upNoRatio = (float)upNoCount / (upYesCount + upNoCount);
            float downYesRatio = (float)downYesCount / (downYesCount + downNoCount);
            float downNoRatio = (float)downNoCount / (downYesCount + downNoCount);

            float totalRatio = (float)upYesCount + upNoCount + downYesCount + downNoCount;
            float downCount = (float)downNoCount + downYesCount;
            float upCount = (float)upNoCount + upYesCount;

            leftGiniDouble = 1 - (Math.Pow(downYesRatio, 2) + Math.Pow(downNoRatio, 2));
            rightGiniDouble = 1 - (Math.Pow(upYesRatio, 2) + Math.Pow(upNoRatio, 2));
            giniDouble = ((leftGiniDouble * downCount) + (rightGiniDouble * upCount)) / totalRatio;
            float gini=(float)giniDouble;
            return gini;
        }

        static List<List<float>> getdugum()
        {
            List<List<float>> dugumValue = new List<List<float>>();
            for (int i = 0; i < 6; i++)
            {
                dugumValue.Add(categoryColumn(i));
            }
            return dugumValue;
        }
        /*static List<string> getResult()
        {
            List<string> resultValues = resultColumn();
            return resultValues;
        }*/
        static List<float> getGiniResult(List<List<float>> dugumValue)
        {
            List<float> giniResult = new List<float>();
            for (int i = 0; i < 6; i++)
            {
                giniResult.Add(giniCalculator(dugumValue[i], resultColumn()));
            }
            return giniResult;
        }
        static Dugum gini(List<List<float>> dugumValue, List<float> giniResult, List<string> resultValues)
        {
            List<List<float>> leftDugumValue = new List<List<float>>();
            List<List<float>> rightDugumValue = new List<List<float>>();
            List<string> leftResultValue = new List<string>();
            List<string> rightResultValue = new List<string>();


            for (int i = 0; i < 6; i++)
            {
                rightDugumValue.Add(new List<float>());
                leftDugumValue.Add(new List<float>());
            }

            float minElement;
            int minElementIndex = 0;
            if (giniResult.Count > 0)
            {
                minElement = giniResult.Min();
                minElementIndex = giniResult.IndexOf(minElement);
            }
            else
            {
                Console.WriteLine("Liste boş.");
            }

            float arithmeticMean = aOrtalama(dugumValue[minElementIndex]);
            for (int i = 0; i < resultValues.Count; i++)
            {
                if (dugumValue[minElementIndex][i] >= arithmeticMean)
                {
                    for (int j = 0; j < 6; j++)
                    {
                        rightDugumValue[j].Add(dugumValue[j][i]);


                    }
                    rightResultValue.Add(resultValues[i]);
                }
                else
                {
                    for (int j = 0; j < 6; j++)
                    {
                        leftDugumValue[j].Add(dugumValue[j][i]);
                    }
                    leftResultValue.Add(resultValues[i]);
                }
            }
            Dugum dugum = new Dugum(dugumValue, leftDugumValue, rightDugumValue, resultValues, leftResultValue, rightResultValue);
            return dugum;
        }

        static void Main(string[] args)
        {
            //Excel verileri okunur.
            List<Dugum> dugumler = new List<Dugum>();
            List<List<float>> mainDugum=getdugum();
            Dugum node = gini(mainDugum, getGiniResult(mainDugum), resultColumn());
            dugumler.Add(node);

            for(int i =0; node.dugumValue.Count > 0;i++)
            {
                while(node.leftDugumValue.Count > 0) { 
                    node = gini(node.leftDugumValue, getGiniResult(node.leftDugumValue), node.leftResultList);
                    dugumler.Add(node);
                }
                if(node.leftDugumValue.Count== 0)
                {
                    node = dugumler[i - 1];
                    node=gini(node.rightDugumValue,getGiniResult(node.rightDugumValue),node.rightResultList);

                }

            }


            //Dictionary<int, List<Dugum>> dugum = new Dictionary<int, List<Dugum>>();

            

            //Dugum dugum = gini();
            






        }
    }
}
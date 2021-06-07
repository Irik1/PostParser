using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.Threading;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OpenQA.Selenium.Interactions;

namespace Selenium
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> cities = new List<string>();
            try
            {
                string path = "Отделения_почты_Почта_Донбасса.txt";
                using (StreamReader sr = new StreamReader(path, System.Text.Encoding.Default))
                {
                    string line;
                    string[] array;

                    while ((line = sr.ReadLine()) != null)
                    {
                        array = line.Split(',');
                        cities.Add(array[2].Replace("г.", "").Replace("пгт.", "").Replace("пос. ", "").Replace("с.", "").Replace("п. ", "").Trim());
                    }

                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            cities = cities.Distinct().ToList();
            IWebDriver driver = new ChromeDriver();
            //File.Delete(@"Почтовые_индексы_Почта_Донбасса.txt");

            // Получить объект приложения Excel.
            Excel.Application excel = new Excel.Application();
            excel.SheetsInNewWorkbook = cities.Count;
            //Добавить рабочую книгу
            Excel.Workbook workBook = excel.Workbooks.Add(Type.Missing);
            //Отключить отображение окон с сообщениями
            excel.DisplayAlerts = false;
            try
            {
                var listcount = 1;
                foreach (var city in cities)
                {
                    try
                    {  
                        driver.Navigate().GoToUrl("https://postdonbass.com/delivery");
                        //driver.Navigate().GoToUrl("https://postdonbass.com/offices");
                        int timeout = 5000;
                        int delay = 2500;
                        IWebElement element = null;
                        IWebElement input;
                        Console.WriteLine("Starting " + city);
                        element = driver.FindElement(By.LinkText("Загрузить ещё"));
                        input = driver.FindElement(By.Id("edit-combine"));
                        input.SendKeys(city + Keys.Enter);
                        try
                        {
                            while (true)
                            {
                                Actions actions = new Actions(driver);
                                actions.MoveToElement(driver.FindElement(By.LinkText("Загрузить ещё")));
                                actions.Perform();
                                Thread.Sleep(delay);
                                driver.FindElement(By.LinkText("Загрузить ещё")).Click();
                            }
                        }
                        catch (Exception e)
                        {
                            //Console.WriteLine(e);
                            Console.WriteLine("Listing finished");
                        }
                        string writePath = @"Почтовые_индексы_Почта_Донбасса.txt";

                        List<string[]> arraylist = new List<string[]>();
                        string[] rowArray;
                        try
                        {
                            var table = driver.FindElement(By.TagName("table"));
                            var rows = table.FindElements(By.TagName("tr"));
                            using (StreamWriter sw = new StreamWriter(writePath, true, System.Text.Encoding.Default))
                            {
                                foreach (var row in rows)
                                {
                                    int i = 0;
                                    var cells = row.FindElements(By.TagName("td"));
                                    rowArray = new string[cells.Count];
                                    foreach (var cell in cells)
                                    {
                                        sw.WriteAsync(cell.Text + "|");
                                        rowArray[i] = cell.Text;
                                        i++;
                                    }
                                    arraylist.Add(rowArray);
                                    sw.WriteLineAsync("");
                                }
                                sw.WriteLineAsync("@");
                            }

                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }

                        //Получаем первый лист документа (счет начинается с 1)
                        Excel.Worksheet sheet = (Excel.Worksheet)excel.Worksheets.get_Item(listcount);
                        //Название листа (вкладки снизу)
                        sheet.Name = city;
                        //Пример заполнения ячеек
                        int l = 0;
                        int j = 0;
                        foreach (var str in arraylist)
                        {
                            l++;
                            j = 0;
                            foreach (var elem in str)
                            {
                                j++;
                                sheet.Cells[l, j] = String.Format(elem);
                            }
                        }
                        //ex.Application.ActiveWorkbook.Save();
                        Console.WriteLine(city+" Finished");
                        listcount++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("exception: {0}", ex);
                    }
                }
                excel.ActiveWorkbook.SaveAs("Почтовые_индексы_Почта_Донбасса.xls",  //object Filename
                    Excel.XlFileFormat.xlHtml,          //object FileFormat
                    Type.Missing,                       //object Password 
                    Type.Missing,                       //object WriteResPassword  
                    Type.Missing,                       //object ReadOnlyRecommended
                    Type.Missing,                       //object CreateBackup
                    Excel.XlSaveAsAccessMode.xlNoChange,//XlSaveAsAccessMode AccessMode
                    Type.Missing,                       //object ConflictResolution
                    Type.Missing,                       //object AddToMru 
                    Type.Missing,                       //object TextCodepage
                    Type.Missing,                       //object TextVisualLayout
                    Type.Missing);                      //object Local
                                                        // Закройте сервер Excel.
                excel.Quit();
                Console.WriteLine("The end");
                Thread.Sleep(5000);
                driver.Close();
            }
            catch (Exception ex)
            {
                driver.Close();
            }
        }
    }
}





//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using OpenQA.Selenium.Chrome;
//using OpenQA.Selenium;
//using OpenQA.Selenium.Support.UI;
//using System.Threading;
//using System.IO;

//namespace Selenium
//{
//    class Program
//    {
//        static void Main(string[] args)
//        {

//            IWebDriver driver = new ChromeDriver();
//            try
//            {
//                driver.Navigate().GoToUrl("https://postdonbass.com/delivery");
//                //driver.Navigate().GoToUrl("https://postdonbass.com/offices");
//                int timeout = 5000;
//                int delay = 2000;
//                IWebElement element;
//                try
//                {

//                    string[] lines_ = File.ReadAllLines(@"Отделения_почты_Почта_Донбасса.txt");
//                    string[] _lines = { }; 
//                    for(int i = 0; i < lines_.Length; i++)
//                    {
//                        _lines[i] = lines_[i].Split(',')[2];
//                    }

//                    element = driver.FindElement(By.LinkText("Загрузить ещё"));
//                    //element = driver.FindElement(By.LinkText("Ещё отделения"));
//                    while (element != null)
//                    {
//                        element.Click();
//                        Console.WriteLine("Click");
//                        Thread.Sleep(delay);
//                        //element = driver.FindElement(By.LinkText("Загрузить ещё"));
//                        element = driver.FindElement(By.LinkText("Ещё отделения"));

//                        //Thread.Sleep(1000);
//                        //button = driver.FindElement(By.LinkText("Загрузить ещё"));
//                    }
//                }
//                catch (Exception ex)
//                {
//                    Console.WriteLine("exception: {0}", ex);
//                }

//                //string writePath = @"Отделения_почты.txt";


//                //try
//                //{
//                //    using (StreamWriter sw = new StreamWriter(writePath, false, System.Text.Encoding.Default))
//                //    {
//                //        foreach (var item in CityList_Pred)
//                //        {
//                //            sw.WriteLineAsync(item.Text);
//                //        }
//                //    }
//                //    Console.WriteLine("Запись выполнена");
//                //}
//                //catch (Exception e)
//                //{
//                //    Console.WriteLine(e.Message);
//                //}

//                //List<string> CityList_Post = new List<string>();
//                //var table = driver.FindElement(By.TagName("table"));
//                //var rows = table.FindElements(By.TagName("tr"));


//                //foreach (var row in rows)
//                //{
//                //    if (row.Text.Contains("Bergen"))
//                //    {
//                //        //Console.WriteLine(row.Text);

//                //        var tds = row.FindElements(By.TagName("a"));
//                //        foreach (var entry in tds)
//                //        {
//                //            Console.WriteLine(entry.Text);
//                //            entry.Click();
//                //        }
//                //    }
//                //}

//                Thread.Sleep(5000);
//                driver.Close();
//            }
//            catch(Exception ex)
//            {
//                driver.Close(); 
//            }
//        }
//    }
//}

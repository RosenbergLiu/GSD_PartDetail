using Microsoft.Win32;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Drawing;
using System.IO;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace PartDetail
{
    public partial class MainWindow : Window
    {


        //string outpath = "C:/Users/xygen/Documents/GitHub/";
        string outpath = "C:/Users/roshan.liu/Scripts/DATA/PartsToPrint/";
        public MainWindow()
        {
            InitializeComponent();
        }


        public (Boolean, string) PrintPartMode1(string PartN, string JobN, string Qty)
        {
            ChromeDriver driver = new ChromeDriver();
            driver.Manage().Window.FullScreen();
            driver.Navigate().GoToUrl(@"https://partners.gorenje.com/GSD/");
            driver.FindElement(By.Id("tbUsr")).SendKeys("liuro_sh");
            driver.FindElement(By.Id("tbPwd")).Clear();
            driver.FindElement(By.Id("tbPwd")).SendKeys("gorenje1");
            driver.FindElement(By.Id("btnLogIn")).Click();
            driver.Navigate().GoToUrl(@"https://partners.gorenje.com/GSD/gsd_iskanje_rd.aspx");
            driver.FindElement(By.Id("ContentPlaceHolder1_tbRD")).Clear();
            driver.FindElement(By.Id("ContentPlaceHolder1_tbRD")).SendKeys(PartN);
            driver.FindElement(By.Id("ContentPlaceHolder1_btnIsk_CD")).Click();
            string originalWindow = driver.CurrentWindowHandle;
            Assert.AreEqual(driver.WindowHandles.Count, 1);
            try
            {
                driver.FindElement(By.Id("ContentPlaceHolder1_gvCenikRd_cell0_0_lnk_0")).Click();
                WebDriverWait w = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
                w.Until(wd => wd.WindowHandles.Count == 2);

                foreach (string window in driver.WindowHandles)
                {
                    if (originalWindow != window)
                    {
                        driver.SwitchTo().Window(window);
                        break;
                    }
                }
                WebElement ImgEle = (WebElement)driver.FindElement(By.Id("ASPxImage1"));
                string fileName1 = "provisional-" + PartN + "@" + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + ".png";
                driver.GetScreenshot().SaveAsFile(outpath + fileName1);
                string fileName2 = PartN + "@" + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + ".png";

                string DetailText = " " + IDOSS.Text;

                driver.Close();
                driver.SwitchTo().Window(originalWindow);
                driver.Quit();
                ChromeDriver driver2 = new ChromeDriver();
                driver2.Navigate().GoToUrl(@"https://partners.gorenje.com/sagCC/Login.aspx");
                driver2.FindElement(By.Id("usr")).SendKeys("liuro_sh");
                driver2.FindElement(By.Id("pwd")).SendKeys("gorenje1");
                driver2.FindElement(By.Id("btnPrijava")).Click();
                driver2.FindElement(By.Id("ctl00_tbOss")).SendKeys(JobN);
                originalWindow = driver2.CurrentWindowHandle;
                driver2.FindElement(By.Id("ctl00_btnOss")).Click();
                Assert.AreEqual(driver2.WindowHandles.Count, 1);

                WebDriverWait w2 = new WebDriverWait(driver2, TimeSpan.FromSeconds(20));
                w2.Until(wd => wd.WindowHandles.Count == 2);

                foreach (string window in driver2.WindowHandles)
                {
                    if (originalWindow != window)
                    {
                        driver2.SwitchTo().Window(window);
                        break;
                    }
                }

                string machine = driver2.FindElement(By.Id("ctl00_ContentPlaceHolder1_lblNazivIzdelka")).Text;
                string ART = driver2.FindElement(By.Id("ctl00_ContentPlaceHolder1_dd_izd_sifra2_dd_izd_sifra_I")).GetAttribute("value");
                string Index = driver2.FindElement(By.Id("ctl00_ContentPlaceHolder1_dd_izd_si_I")).GetAttribute("value");
                string JobNo = driver2.FindElement(By.Id("ctl00_ContentPlaceHolder1_lbl_st_naloga")).Text;
                string Tech = driver2.FindElement(By.Id("ctl00_ContentPlaceHolder1_txtServiser")).GetAttribute("value");
                string JOB = "Job: " + JobNo.Split(".")[1];
                string ArtIndex = "ART: " + ART + "/" + Index;
                string Quantity = "x " + Qty;

                PointF fLocation1 = new PointF(10f, 500f);
                PointF fLocation2 = new PointF(10f, 580f);
                PointF fLocation3 = new PointF(10f, 660f);
                PointF fLocation4 = new PointF(10f, 740f);
                PointF fLocation5 = new PointF(10f, 820f);

                string imgPath = outpath + fileName1;
                Bitmap bitmap = (Bitmap)Image.FromFile(imgPath);

                string imgPath2 = outpath + fileName2;
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    using (Font arialFont = new Font("Arial", 20))
                    {
                        graphics.DrawString(machine, arialFont, Brushes.Blue, fLocation1);
                        graphics.DrawString(ArtIndex, arialFont, Brushes.Blue, fLocation2);
                        graphics.DrawString(JOB, arialFont, Brushes.Blue, fLocation3);
                        graphics.DrawString(Quantity, arialFont, Brushes.Blue, fLocation4);
                        //graphics.DrawString(Tech, arialFont, Brushes.Blue, fLocation4);
                    }
                }
                bitmap.Save(imgPath2);
                bitmap.Dispose();
                File.Delete(imgPath);
                driver2.Close();
                driver2.SwitchTo().Window(originalWindow);
                driver2.Quit();
                return (true, machine);
            }
            catch (NoSuchElementException)
            {
                driver.Quit();
                return (false, "No Picture");
            }
        }


        private void GenerateOneNote_Click(object sender, RoutedEventArgs e)
        {

            string pNumber = PartNumber.Text;
            string jNumber = IDOSS.Text;
            string qty = QTY.Text;
            var result = PrintPartMode1(pNumber, jNumber, qty);
            if (!result.Item1)
            {
                MessageBox.Show("No picture for this part");
            }
        }

        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.DefaultExt = ".xlsm";
            Nullable<bool> dialogOK = openFileDialog.ShowDialog();
            if (dialogOK == true)
                ExcelPath.Text = openFileDialog.FileName;
        }

        private void btnGenerateNotes_Click(object sender, RoutedEventArgs e)
        {
            if (ExcelPath.Text == "Select an Excel")
            {
                MessageBox.Show("Please select an excel file");
            }
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open(ExcelPath.Text);
            Excel.Worksheet worksheet = workbook.Worksheets[1];
            int iniRow = 2;
            while (worksheet.Cells[iniRow, 1].Value2 != null)
            {
                string Part = worksheet.Cells[iniRow, 1].Value2.ToString();
                string Job = worksheet.Cells[iniRow, 4].Value2.ToString();
                string Qty = "1";
                if (worksheet.Cells[iniRow, 2].Value2 == null)
                {
                    Qty = "1";
                }
                else
                {
                    Qty = worksheet.Cells[iniRow, 2].Value2.ToString();
                }

                var result = PrintPartMode1(Part, Job, Qty);
                if (result.Item1 == false)
                {
                    worksheet.Cells[iniRow, 3] = "No Picture";
                    worksheet.Cells[iniRow, 6] = "Need manual solve";
                }
                else
                {
                    worksheet.Cells[iniRow, 3] = result.Item2;
                    worksheet.Cells[iniRow, 6] = "Done";
                }
                iniRow++;
            }
            workbook.Save();
            workbook.Close();
            excel.Quit();
        }


    }
}
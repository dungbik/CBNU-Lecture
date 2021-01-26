using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Threading;
using System.Xml;
using static SeleniumExtras.WaitHelpers.ExpectedConditions;

namespace CBNU_Lecture
{
    class Program
    {
        static string userId = "";
        static string userPW = "";

        static void Main(string[] args)
        {
            try
            {
                loadConfigFile();

                ChromeDriverService _driverService = ChromeDriverService.CreateDefaultService();
                _driverService.HideCommandPromptWindow = true;

                ChromeOptions _options = new ChromeOptions();
                _options.AddArgument("disable-gpu");

                ChromeDriver _driver = new ChromeDriver(_driverService, _options);
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));


                _driver.Navigate().GoToUrl("https://eis.cbnu.ac.kr/");

                var uid = wait.Until(ElementExists(By.XPath("//*[@id='uid']")));
                uid.SendKeys(userId);

                var pswd = _driver.FindElementByXPath("//*[@id='pswd']");
                pswd.SendKeys(userPW);

                var loginButton = _driver.FindElementByXPath("//*[@id='commonLoginBtn']");
                loginButton.Click();

                try
                {
                    wait.Until(ElementExists(By.XPath("//*[@id='mainframe_VFS_LoginFrame_ssoLoginTry_form_btn_y']"))).Click();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                while (true)
                {
                    try
                    {
                        wait.Until(ElementExists(By.XPath("//*[@id='mainframe_VFS_HFS_SubFrame_form_tab_subMenu_tpg_bssMenu_grd_menu_body_gridrow_8']"))).Click();
                        break;
                    }
                    catch (Exception ex)
                    {
                        Thread.Sleep(1000);
                    }
                }

                wait.Until(ElementExists(By.XPath("//*[@id='mainframe_VFS_HFS_SubFrame_form_tab_subMenu_tpg_bssMenu_grd_menu_body_gridrow_9']"))).Click();
                wait.Until(ElementExists(By.XPath("//*[@id='mainframe_VFS_HFS_SubFrame_form_tab_subMenu_tpg_bssMenu_grd_menu_body_gridrow_16']"))).Click();
                wait.Until(ElementExists(By.XPath("//*[@id='mainframe_VFS_HFS_INVFS_WorkFrame_win_2275_form_div_work_btn_search']"))).Click();


                Dictionary<int, ArrayList> list = new Dictionary<int, ArrayList>();

                var downButton = _driver.FindElementByXPath("//*[@id='mainframe_VFS_HFS_INVFS_WorkFrame_win_2275_form_div_work_grd_master_vscrollbar_incbutton']");

                while (list.Count < 3880)
                {
                    var table = wait.Until(ElementExists(By.XPath("//*[@id='mainframe_VFS_HFS_INVFS_WorkFrame_win_2275_form_div_work_grd_master_bodyGridBandContainerElement_inner']")));
                    if (table.FindElements(By.TagName("div")).Count <= 2)
                        continue;


                    var test = table.FindElement(By.XPath($"//*[@id='mainframe_VFS_HFS_INVFS_WorkFrame_win_2275_form_div_work_grd_master_body_gridrow_{(list.Count + 1) % 12}']"));
                    var index = Int32.Parse(test.FindElement(By.XPath($"//*[@id='mainframe_VFS_HFS_INVFS_WorkFrame_win_2275_form_div_work_grd_master_body_gridrow_{list.Count % 12}_cell_{list.Count % 12}_0']")).GetAttribute("aria-label"));

                    if (list.ContainsKey(index))
                        continue;

                    if (list.Count >= 3880)
                        break;

                    ArrayList row = new ArrayList();
                    for (var j = 1; j <= 20; j++)
                    {
                        try
                        {
                            row.Add(test.FindElement(By.XPath($"//*[@id='mainframe_VFS_HFS_INVFS_WorkFrame_win_2275_form_div_work_grd_master_body_gridrow_{list.Count % 12}_cell_{list.Count % 12}_{j}']")).GetAttribute("aria-label"));
                        }
                        catch (Exception ex)
                        {
                            row.Add("");
                        }
                    }
                    Console.WriteLine(row[2]);
                    list.Add(index, row);
                    if (list.Count % 12 == 0)
                    {
                        for (var i = 1; i <= 12; i++)
                            downButton.Click();
                        _driver.FindElementByXPath($"//*[@id='mainframe_VFS_HFS_INVFS_WorkFrame_win_2275_form_div_work_grd_master_body_gridrow_0_cell_0_1']").Click();
                    }

                }

                for (var i = 1; i <= 3880; i++)
                {
                    for (var j = 0; j < 20; j++)
                    {
                        Console.Write($"{((ArrayList)list[i])[j]} ");
                    }
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        static void loadConfigFile()
        {
            XmlDocument xml = new XmlDocument(); 
            xml.Load(@".\config.xml");
            XmlNodeList nodes = xml.SelectNodes("/config"); 
            
            foreach (XmlNode node in nodes) {
                userId = node["userId"].InnerText;
                userPW = node["userPW"].InnerText;
            }
            Console.WriteLine($"{userId} {userPW}");
            
        }
    }
}

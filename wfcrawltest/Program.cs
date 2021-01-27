using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace wfcrawltest
{
    static class Program
    {
        /// <summary>
        /// 해당 애플리케이션의 주 진입점입니다.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //CMD창 숨기기 : new ChromeDriver(driverService, new ChromeOptions())
            var driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;
            //크롬 창 옵션
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("headless"); //크롬 창 숨김 여부. headless : 창 숨김 
            options.AddArgument("window-size=2400x1080"); //크롬 창의 크기 결정 
            options.AddArgument("disable-gpu"); //gpu 사용하지 않음 
            options.AddArgument("user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6)"+ "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36"); //bot 이 아님을 표현


            //크롬 창 띄우기
            //using (var driver = new ChromeDriver(driverService, new ChromeOptions()))
            //크롬 창 숨김
            using (var driver = new ChromeDriver(driverService, options))
            {
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10); //화면로드 까지 최대 3초 대기




                string timecheck = DateTime.Now.ToString("yyyy-MM-dd HHmmss");
                string savePath = "logs/" + timecheck + ".txt";
                string savePath2 = "logs/" + timecheck + ".csv";
                var result = "heol123123o \n" + DateTime.Now.ToString() + "\n";

                // 접근할 웹사이트 url
                driver.Navigate().GoToUrl("https://scm.i-sens.com/login");

                // Get the page elements
                var userNameField = driver.FindElementById("UID");
                var userPasswordField = driver.FindElementById("UPW");
                var loginButton = driver.FindElementByXPath("//*[@id='btnLogin']");

                // 로그인 정보 입력
                userNameField.SendKeys("admin");
                userPasswordField.SendKeys("isens@2019");

                // 로그인
                loginButton.Click();
                // 무상납품리스트
                var freebutton = driver.FindElementByXPath("//*[@id='cssmenu']/ul/li[4]/p/a/span/b");
                freebutton.Click();
                var freelistbutton = driver.FindElementByXPath("//*[@id='cssmenu']/ul/li[4]/ul/li[1]/a");
                freelistbutton.Click();
                //조건검색
                //var optionbutton = driver.FindElementByXPath("//*[@id='search_concept']");
                //optionbutton.Click();
                //var optionponum = driver.FindElementByXPath("//*[@id='li_on']");
                //optionponum.Click();
                //var optionfield = driver.FindElementByXPath("//*[@id='txtSearch']");
                //optionfield.SendKeys("19");
                //var searchbutton = driver.FindElementByXPath("//*[@id='btnSearch']/span");
                //searchbutton.Click();

                // 페이지 테이블 내용 가져오기
                //var itemcodes = driver.FindElementByXPath("//*[@id='datatable']/tbody/tr[3]/td[6]").Text;
                //result += itemcodes;

                // 수집내용 저장할 DT
                DataTable dt = new DataTable();
                //DataColumn 생성
                DataColumn colPo = new DataColumn("PO", typeof(string));
                DataColumn colItemcode = new DataColumn("ITEMCODE", typeof(string));
                DataColumn colItemname = new DataColumn("ITEMNAME", typeof(string));
                DataColumn colTime = new DataColumn("CREATETIME", typeof(DateTime));
                //생성된 Column을 DataTable에 Add
                dt.Columns.Add(colPo);
                dt.Columns.Add(colItemcode);
                dt.Columns.Add(colItemname);
                dt.Columns.Add(colTime);

                // PK 설정으로 중복 제거
                DataColumn[] pk = new DataColumn[1];
                pk[0] = dt.Columns["PO"];
                dt.PrimaryKey = pk;

                var trows = driver.FindElementByXPath("//*[@id='datatable']/tbody/tr");
                int count = driver.FindElementsByXPath("//*[@id='datatable']/tbody/tr").Count();

                for (int i = 1; i-1 < count; i++)
                {
                    try
                    {
                        string tdpath = "//*[@id='datatable']/tbody/tr[" + i + "]";
                        string poPath = tdpath + "/td[3]";
                        string codePath = tdpath + "/td[6]";
                        string namePath = tdpath + "/td[7]";
                        string poNo = driver.FindElementByXPath(poPath).Text;
                        string itemCode = driver.FindElementByXPath(codePath).Text;
                        string itemName = driver.FindElementByXPath(namePath).Text;

                        dt.Rows.Add(poNo, itemCode, itemName, DateTime.Now);
                    }
                    catch { continue; }
                }

                // DB에 저장
                inputdata(dt);

                // CSV파일로 저장
                StringBuilder sb = new StringBuilder();
                IEnumerable<string> columnNames = dt.Columns.Cast<DataColumn>().Select(column => column.ColumnName);
                sb.AppendLine(string.Join(",", columnNames));
                foreach (DataRow row in dt.Rows)
                {
                    IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                    sb.AppendLine(string.Join(",", fields));
                }
                byte[] bytes = Encoding.UTF8.GetBytes(sb.ToString());
                string body = Encoding.UTF8.GetString(bytes);

                File.WriteAllText(savePath2, body);

                //var result = driver.FindElementByXPath("//div[@id='case_login']/h3").Text;

                File.WriteAllText(savePath, result);
                // Take a screenshot and save it into screen.png
                driver.GetScreenshot().SaveAsFile("logs/freescreen.png");
                //inputdata();

                driver.Quit();

            }
        }
        public static void inputdata(DataTable dtt)
        {
            SqlConnection con = new SqlConnection("server=171.16.1.60;database=ISENSE_SCM;user id=CASWEB;password=start@1234;Trusted_Connection=no;");
            con.Open();

            // DataTable DB에 insert
            using (var bulk = new SqlBulkCopy(con))
            {
                bulk.DestinationTableName = "CASDB2";
                try
                {
                    bulk.WriteToServer(dtt);
                }
                catch(Exception e)
                {
                    string err = e.ToString();
                    File.WriteAllText("logs/error.txt", err);
                }
            }
            //SqlCommand cmd = new SqlCommand();
            //cmd.Connection = con;
            //cmd.CommandText = "insert into CASDB2 (CASNO, NAME) values ('12223','TEST2');";
            //cmd.ExecuteReader();
            //cmd.Dispose();
        }


    }
}

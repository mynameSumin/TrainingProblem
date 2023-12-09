using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using OfficeOpenXml;

namespace Training
{
    static class Program
    {
        
        static void Main(string[] args)
        {
            // ChomeDriver 인스턴스 생성
            using (IWebDriver driver = new ChromeDriver())
            {
                driver.Url = "https://www.g2b.go.kr/index.jsp"; // 나라장터 URL로 접속
                Console.WriteLine("나라장터 접속");

                // 대기 설정. (find로 객체를 찾을 때까지 검색이 되지 않으면 대기하는 시간 초단위)
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);


                //오늘 날짜 가져오기
                DateTime startDate = DateTime.Now.AddDays(-4); //오늘 포함 최근 5일 데이터 가져오는 것으로 설정
                Console.WriteLine(startDate);
                String startDateString = startDate.ToString("yyyyMMdd");
                Console.WriteLine(startDateString);

                //시작날짜 입력란에 날짜 삽입
                var startDateBtn = driver.FindElement(By.Id("fromBidDt"));
                startDateBtn.Clear(); //기본 값 삭제
                startDateBtn.SendKeys(startDateString);

                Thread.Sleep(1000);
                //rpa를 공고명에 넣고 검색(예외처리 필요)
                driver.FindElement(By.XPath("//*[@id='bidNm']")).SendKeys("RPA");


                // xpath로 검색 버튼을 찾는다. 
                var searchBtn = driver.FindElement(By.XPath("//*[@id='searchForm']/div/fieldset[1]/ul/li[4]/dl/dd[3]/a"));
                Console.WriteLine("검색버튼 찾음");
                driver.ClickScript(searchBtn);

                Thread.Sleep(3000);
                //모든 데이터 가져오기
                GetData(driver, "rpa");

                Thread.Sleep(10000);
            }
            
            // 아무 키나 누르면 종료
            Console.WriteLine("Press any key...");
            Console.ReadKey();
        }


        // driver를 Script 실행 인터페이스로 변환
        static IJavaScriptExecutor Script(this IWebDriver driver)
        {
            return (IJavaScriptExecutor)driver;
        }
        // 스크립트 클릭 함수
        static void ClickScript(this IWebDriver driver, IWebElement element)
        {
            driver.Script().ExecuteScript("arguments[0].click();", element);
        }

        //RPA가 정확하게 포함된 데이터만 가져오기
        //예를 들어 copter parts는 rpa를 포함하고 있지만 우리가 원하는 데이터가 아니다.
        static void GetData(IWebDriver driver, string keyward)
        {
            //frame 안에 있는경우 따로 처리
            driver.SwitchTo().Frame("sub");
            driver.SwitchTo().Frame("main");

            var table = driver.FindElement(By.TagName("tbody"));
            var cols = table.FindElements(By.TagName("tr"));
            int index = 1;

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);  // 바탕화면 경로
            string path = Path.Combine(desktopPath, "장표.xlsx"); // 엑셀 파일 저장 경로
            FileInfo file = new FileInfo(path);

            
            using(ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                var dimension = worksheet.Dimension;
                int lastRow = dimension.End.Row;
                int newRow = lastRow + 1;
                
                foreach (var row in cols)
                {
                    var name = row.FindElement(By.XPath("//*[@id='resultForm']/div[2]/table/tbody/tr[" + index + "]/td[4]/div/a"));
                    if (name.Text.Contains("RPA"))
                    {
                        var announce = row.FindElement(By.XPath("//*[@id='resultForm']/div[2]/table/tbody/tr["+ index + "]/td[5]/div"));
                        var demand = row.FindElement(By.XPath("//*[@id='resultForm']/div[2]/table/tbody/tr["+ index +"]/td[6]/div"));
                        var deadLine = row.FindElement(By.XPath("//*[@id='resultForm']/div[2]/table/tbody/tr[" + index + "]/td[8]/div"));
                        var childElements = deadLine.FindElements(By.XPath("./*"));
                        string spanText = deadLine.FindElement(By.TagName("span")).Text;
                        string allText = deadLine.Text;
                        string stringDeadLine = allText.Replace(spanText, "").Trim();
                        int start = 1;
                        // 해당 행 출력 또는 저장 등 필요한 작업 수행
                        worksheet.Cells[newRow, start++].Value = name.Text;
                        worksheet.Cells[newRow, start++].Value = "RPA";
                        worksheet.Cells[newRow, start++].Value = announce.Text;
                        worksheet.Cells[newRow, start++].Value = demand.Text;
                        worksheet.Cells[newRow, start++].Value = stringDeadLine;
                        worksheet.Cells[newRow, start++].Value = DateTime.Now.ToString("yyyy/MM/dd");
                        newRow++;
                    }
                    index++;
                }
                package.Save();
            }
            Console.WriteLine("데이터가 성공적으로 추가되었습니다.");
        }

    }
}

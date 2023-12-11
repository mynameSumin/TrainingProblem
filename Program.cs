using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using System.IO;
using OfficeOpenXml;

namespace Training
{
    static class Program
    {

        static void Main(string[] args)
        {
            using (IWebDriver driver = new ChromeDriver())
            {
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2); // 대기 시간 설정
                driver.Url = "https://www.g2b.go.kr/index.jsp"; // 나라장터 URL로 접속

                //오늘 날짜 가져오기
                DateTime startDate = DateTime.Now.AddDays(-4); //오늘 포함 최근 5일 데이터 가져오는 것으로 설정
                Console.WriteLine(startDate);
                String startDateString = startDate.ToString("yyyyMMdd");
                Console.WriteLine(startDateString);

                //시작날짜 입력란에 날짜 삽입
                var startDateBtn = driver.FindElement(By.Id("fromBidDt"));
                startDateBtn.Clear(); //기본 값(한 달 전) 삭제
                startDateBtn.SendKeys(startDateString);

                //rpa를 공고명에 넣고 검색
                driver.FindElement(By.XPath("//*[@id='bidNm']")).SendKeys("RPA"); //검색창
                var searchBtn = driver.FindElement(By.XPath("//*[@id='searchForm']/div/fieldset[1]/ul/li[4]/dl/dd[3]/a")); //검색 버튼
                driver.ClickScript(searchBtn);

                //모든 데이터 가져오기
                while (true)
                {
                    try
                    {
                        GetData(driver, "RPA");
                        var plus = driver.FindElement(By.Id("pagination")).FindElement(By.ClassName("default"));
                        driver.ClickScript(plus);
                        Thread.Sleep(2000);
                        driver.SwitchTo().DefaultContent();
                    }
                    catch (NoSuchElementException)
                    {
                        Console.WriteLine("데이터 추가 완료");
                        Thread.Sleep(5000);
                        return;
                    }
                }

            }

            // 아무 키나 누르면 종료
            Console.WriteLine("프로그램 종료");
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
        static void GetData(IWebDriver driver, string keyward)
        {
            //frame 안에 있는경우 따로 처리
            driver.SwitchTo().Frame("sub");
            driver.SwitchTo().Frame("main");

            var table = driver.FindElement(By.TagName("tbody"));
            var cols = table.FindElements(By.TagName("tr"));
            int index = 1;

            foreach (var row in cols)
            {
                var name = row.FindElement(By.XPath("//*[@id='resultForm']/div[2]/table/tbody/tr[" + index + "]/td[4]/div/a")).Text;
                if (name.Contains(keyward))
                {
                    Console.WriteLine(name); //추가되는 데이터 이름
                    var announce = row.FindElement(By.XPath("//*[@id='resultForm']/div[2]/table/tbody/tr[" + index + "]/td[5]/div")).Text;
                    var demand = row.FindElement(By.XPath("//*[@id='resultForm']/div[2]/table/tbody/tr[" + index + "]/td[6]/div")).Text;
                    var deadLine = row.FindElement(By.XPath("//*[@id='resultForm']/div[2]/table/tbody/tr[" + index + "]/td[8]/div"));
                    var childElements = deadLine.FindElements(By.XPath("./*"));
                    string spanText = deadLine.FindElement(By.TagName("span")).Text;
                    string allText = deadLine.Text;
                    string stringDeadLine = allText.Replace(spanText, "").Trim();

                    // 해당 행 출력 또는 저장 등 필요한 작업 수행
                    saveExcel(name, keyward, announce, demand, stringDeadLine);
                }
                index++;
            }
        }

        static void saveExcel(string name, string keyword, string announce, string demand, string deadLine)
        {
            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);  // 바탕화면 경로
                string path = Path.Combine(desktopPath, "장표.xlsx"); // 엑셀 파일 저장 경로
                FileInfo file = new FileInfo(path);
                ExcelPackage.LicenseContext = LicenseContext.Commercial;


                using (ExcelPackage package = new ExcelPackage(file))
                {
                    if (package.Workbook.Worksheets.Count <= 0)
                    {
                        throw new Exception("엑셀 파일이 없습니다");
                    }
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    var dimension = worksheet.Dimension;
                    int lastRow = dimension.End.Row;
                    int newRow = lastRow + 1; //마지막 행 다음줄

                    int start = 1; //열 위치
                    worksheet.Cells[newRow, start++].Value = name;
                    worksheet.Cells[newRow, start++].Value = keyword;
                    worksheet.Cells[newRow, start++].Value = announce;
                    worksheet.Cells[newRow, start++].Value = demand;
                    worksheet.Cells[newRow, start++].Value = deadLine;
                    worksheet.Cells[newRow, start++].Value = DateTime.Now.ToString("yyyy/MM/dd hh:mm");
                    package.Save();
                }
            }
            catch (InvalidOperationException err)
            {
                Console.WriteLine("장표 엑셀 파일이 열려있을 확률이 높습니다.");
                Console.WriteLine("err message: " + err);
                throw new Exception();
            }
            catch (IndexOutOfRangeException err)
            {
                Console.WriteLine("장표 파일이 바탕화면에 없습니다");
                throw new Exception();
            }
        }
    }
}
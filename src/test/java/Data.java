import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class Data {


    public static void main(String args[]) throws Exception {

        System.setProperty("webdriver.chrome.driver", "src/test/resources/driver/chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        try {
            FileInputStream fileInputStream = new FileInputStream("src/test/resources/data/data.xlsx");
            Workbook wb = new XSSFWorkbook(fileInputStream);
            driver.get("https://www.brainbench.com/xml/bb/common/testcenter/consumer/alltests.xml");
            List<WebElement> list = driver.findElements(By.xpath("//tr/td[1]/a"));
            driver.manage().window().maximize();
            String job;
            Sheet sheet = wb.getSheet("Jobs");
            for (int i = 0; i < 2; i++) {
                list = driver.findElements(By.xpath("//tr/td[1]/a"));
                job = list.get(i).getText();
                sheet.createRow(i).createCell(1).setCellValue(job);
                list.get(i).click();
                Thread.sleep(5000);
                int tableCount = driver.findElements(By.xpath("//tbody")).size();

                List<WebElement> skills = driver.findElements(By.xpath("//td"));
                String skill;
                for (int j = 0; j < skills.size(); j++) {
                    skill = skills.get(j).getText();
                    sheet.getRow(i).createCell(j + 2).setCellValue(skill);
                }
                driver.navigate().back();
                Thread.sleep(5000);
            }
            FileOutputStream fileOutputStream = new FileOutputStream("src/test/resources/data/data.xlsx");
            wb.write(fileOutputStream);

        }
        finally{
            driver.quit();
        }

    }
}

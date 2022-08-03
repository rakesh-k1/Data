import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

public class Data {


    public static void main(String args[]) throws Exception {

        System.setProperty("webdriver.chrome.driver", "src/test/resources/driver/chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        try {
            FileInputStream fileInputStream = new FileInputStream("src/test/resources/data/data.xlsx");
            Workbook wb = new XSSFWorkbook(fileInputStream);
            driver.get("https://www.brainbench.com/xml/bb/common/testcenter/consumer/alltests.xml");
            Thread.sleep(5000);
            driver.manage().window().maximize();
//            List<WebElement> list = driver.findElements(By.xpath("//tr/td[1]/a"));
            String job;
            Sheet sheet = wb.getSheet("Jobs");
            for (int i = 0; i < 4; i=i+2) {
                List<WebElement> list = driver.findElements(By.xpath("//tr/td[1]/a"));
                System.out.println(list.size());
                job = list.get(i).getText();
                Row row = sheet.createRow(i);
                Row row1 = sheet.createRow(i + 1);
                row.createCell(1).setCellValue(job);
                list.get(i).click();
                Thread.sleep(5000);
                driver.findElement(By.xpath("//table[2]/tbody"));
                List<WebElement> rows = driver.findElements(By.xpath("//table[2]/tbody/tr"));
                System.out.println(rows.size());
                if(rows.size()>1){
                    int xy=0;
                    int heading=driver.findElements(By.xpath("//table[2]/tbody/tr/td/dl/dt")).size();
                for(int x=1;x<=rows.size();x++){
                    String colHeadXpath;
                    String colDataXpath;
                    String colsData;
                    WebElement col2;
                    WebElement colHead;
                    List<WebElement> colData;
                    List<WebElement> cols = driver.findElements(By.xpath("//table[2]/tbody/tr["+x+"]/td"));
                    for (int y=1;y<=cols.size();y++){
                        colHeadXpath="//table[2]/tbody/tr["+x+"]/td["+y+"]/dl/dt";
                        colDataXpath="//table[2]/tbody/tr["+x+"]/td["+y+"]/dl/dd";
                        colsData="//table[2]/tbody/tr["+x+"]/td["+y+"]";
                        col2 = driver.findElement(By.xpath(colsData));
                        if(col2.getText().isEmpty()){

                        }
                        else{
                            colHead = driver.findElement(By.xpath(colHeadXpath));
                            colData = driver.findElements(By.xpath(colDataXpath));
                            System.out.println(colHead.getText());
                            for(int z=0;z<1;z++) {
                                String skill="";
                                row.createCell(xy + 3).setCellValue(colHead.getText());
                                for (int xx = 0; xx < colData.size(); xx++) {
                                    skill = skill + " \r\n" + colData.get(xx).getText();
                                }
                                System.out.println(skill);
                                row1.createCell(xy+3).setCellValue(skill);
                            }
                            xy++;
                            }
                        }


                }}
                else{
                    int xy=0;
                    List<WebElement> cols = driver.findElements(By.xpath("//table[2]/tbody/tr/td"));
                    for (int x=1;x<= cols.size();x++) {
                        int size=driver.findElements(By.xpath("//table[2]/tbody/tr/td[1]/dl/dt")).size();
                        for (int y = 1; y <= size; y++) {
                            String colHeadXpath = "//table[2]/tbody/tr/td[1]/dl[" + y + "]/dt";
                            String colDataXpath = "//table[2]/tbody/tr/td[1]/dl[" + y + "]/dt/following-sibling::dd";
                            String colsData = "//table[2]/tbody/tr/td[" + y + "]";
                            WebElement colHead = driver.findElement(By.xpath(colHeadXpath));
                            List<WebElement> colData = driver.findElements(By.xpath(colDataXpath));
                            for(int z=0;z<1;z++) {
                                String skill="";
                                row.createCell(xy + 3).setCellValue(colHead.getText());
                                for (int xx = 0; xx < colData.size(); xx++) {
                                    skill = skill + " \r\n" + colData.get(xx).getText();
                                }
                                System.out.println(skill);
                                row1.createCell(xy+3).setCellValue(skill);
                            }
                            xy++;
                        }
                    }
                driver.navigate().back();
                Thread.sleep(5000);
            }
            FileOutputStream fileOutputStream = new FileOutputStream("src/test/resources/data/data.xlsx");
            wb.write(fileOutputStream);

        }}
        finally{
            driver.quit();
        }

    }
}

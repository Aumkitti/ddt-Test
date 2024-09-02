package ddt;

import static org.junit.jupiter.api.Assertions.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Date;
import java.text.SimpleDateFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

class TS1_verify_user {

    @BeforeAll
    static void setUpBeforeClass() throws Exception {
        // Set up system property for ChromeDriver
        System.setProperty("webdriver.chrome.driver", "D:/chromedriver.exe");
    }

    @Test
    void testCheckLogin() throws Exception {
        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
        Date date = new Date(0);
        String testDate = formatter.format(date);
        String testerName = "Kittipong Dachjit";

        String path = "C:/Users/HP-NPRU/Desktop/testdata.xlsx";
        FileInputStream fs = new FileInputStream(path);

        // Creating a workbook
        XSSFWorkbook workbook = new XSSFWorkbook(fs);
        XSSFSheet sheet = workbook.getSheetAt(0);
        int row = sheet.getLastRowNum() + 1;

        // Open a single WebDriver instance outside the loop
        WebDriver driver = new ChromeDriver();
        WebDriverWait wait = new WebDriverWait(driver, 10); // Explicit wait for 10 seconds

        for (int i = 1; i < row; i++) {
            driver.get("http://localhost:5173/advice");

            WebElement elementToClick = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div[2]/div[2]/div[3]"));
            elementToClick.click();
             
            WebElement addplanToClick = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div[2]/div[2]/div[3]/button[2]"));
            addplanToClick.click();
            
            String year = sheet.getRow(i).getCell(1).toString();
            WebElement yearField = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div[2]/div/div/form/div[1]/input"));
            yearField.sendKeys(year);

            Thread.sleep(1000);
            
            WebElement dropdown = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div[2]/div/div/form/div[2]/select"));
            dropdown.click();

            WebElement secondOption = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div[2]/div/div/form/div[2]/select/option[2]"));
            secondOption.click();
            
            Thread.sleep(1000);

            WebElement dropdownsomery = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div[2]/div/div/div[1]/div[1]/select"));
            dropdownsomery.click();

            WebElement secondOption1 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div[2]/div/div/div[1]/div[1]/select/option[2]"));
            secondOption1.click();
            
            Thread.sleep(1000);
            
            WebElement addcourseToClick = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div[2]/div/div/div[1]/div[2]/button"));
            addcourseToClick.click();
            
            Thread.sleep(1000);
            
            WebElement dropcate = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div[2]/div[2]/div/div[1]/select"));
            dropcate.click();

            WebElement secondOptioncate = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div[2]/div[2]/div/div[1]/select/option[2]"));
            secondOptioncate.click();
            
            Thread.sleep(1000);
            
            WebElement dropgroup = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div[2]/div[2]/div/div[2]/select"));
            dropgroup.click();

            WebElement secondOptiongroup = driver.findElement(By.xpath("/html/body/div/div/div[2]/div[2]/div/div[2]/select/option[2]"));
            secondOptiongroup.click();
            
            Thread.sleep(1000);
            
            WebElement dropcourse = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div[2]/div[2]/div/div[3]/select"));
            dropcourse.click();

            WebElement secondOptioncourse = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div[2]/div[2]/div/div[3]/select/option[3]"));
            secondOptioncourse.click();
            
            Thread.sleep(1000);
            
            WebElement elementToClicksave = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div[2]/div[2]/div/div[4]/button[1]"));
            elementToClicksave.click();
            
            Thread.sleep(1000);
            
            WebElement elementToClicksaveplan = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div[2]/div/div/div[2]/button[2]"));
            elementToClicksaveplan.click();


            // Wait for the specific element (button) to be visible
            WebElement resultButton = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"root\"]/div/div[2]/div[2]/div[2]/div/div/div/button")));
            String actual = resultButton.getText();
            String expected = sheet.getRow(i).getCell(2).toString();

            // Create an Excel row for results
            Row rows = sheet.getRow(i);
            Cell cell = rows.createCell(3);
            cell.setCellValue(actual);

            // Check if the result message contains the expected class group
            boolean isPlanAdded = actual.contains("64/46");

            if (expected.equals(actual)) {
                Cell cell2 = rows.createCell(4);
                cell2.setCellValue("Pass");
            } else {
                Cell cell2 = rows.createCell(4);
                cell2.setCellValue("Fail");
            }

            Cell cell3 = rows.createCell(5);
            cell3.setCellValue(testDate);
            Cell cell4 = rows.createCell(6);
            cell4.setCellValue(testerName);

            // Update the Excel file only once after the loop
            FileOutputStream fos = new FileOutputStream(path);
            workbook.write(fos);
            fos.close();
        }

        driver.quit();
        workbook.close();
        fs.close();
    }
}

package com.rating.mobile;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class excelIntegration {

    public WebDriver driver;

    // Read and write in excel sheet
    public void excelIntegration(String fileName, String playStoreUrl, String appStoreUrl) throws IOException, NullPointerException {
        excelIntegration integration = new excelIntegration();
        integration.startDriver();
        File file = new File(fileName);
        FileInputStream input = new FileInputStream(file);
        HSSFWorkbook workbook = new HSSFWorkbook(input);
        HSSFSheet sheet = workbook.getSheet("MobileApps");
        int rowsize = sheet.getLastRowNum() - sheet.getFirstRowNum();
        for (int i = 1; i <= rowsize; i++) {
            String mobileAppName = sheet.getRow(i).getCell(0).getStringCellValue().toLowerCase();
            System.out.println("MobileApp : " + mobileAppName.toUpperCase());
            integration.searchApp(mobileAppName);
            sheet.getRow(i).createCell(1).setCellValue(integration.fetch(playStoreUrl, mobileAppName));
            sheet.getRow(i).createCell(2).setCellValue(integration.fetch(appStoreUrl, mobileAppName));
            FileOutputStream outputStream = new FileOutputStream(fileName);
            workbook.write(outputStream);
            outputStream.close();
        }
        integration.closeDriver();
    }

    public void startDriver() {
        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.navigate().to("https://www.google.com");
    }

    public void closeDriver() {
        driver.quit();
    }

    public void searchApp(String appName) {
        driver.findElement(By.cssSelector("input[name='q']")).clear();
        driver.findElement(By.cssSelector("input[name='q']")).sendKeys(appName);
        driver.findElement(By.cssSelector("input[name='q']")).sendKeys(Keys.ENTER);
    }

    // Fetching the ratings
    public String fetch(String storeUrl, String mobileAppName) {
        String customerRating = "n/a";
        int currentPage = 1, maxPage = 3;
        driver.findElement(By.xpath("//table[@role='presentation']//tr/td[2]")).click();
        do {
            try {
                customerRating = driver.findElement(By.xpath("//div[contains(@data-async-context,'query')]/div//a[contains(@href,'" + storeUrl + "') and contains(@href,'" + mobileAppName + "')]//parent::div//parent::div//parent::div//span[contains(text(),'Rating: ')]")).getText();
                break;
            } catch (NoSuchElementException e) {
                driver.findElement(By.xpath("//span[text()='Next']")).click();
                System.out.println("Clicking on next " + currentPage++);
                currentPage++;
            }
        } while (currentPage < maxPage);
        return customerRating.replaceAll("Rating: ", "");
    }
}


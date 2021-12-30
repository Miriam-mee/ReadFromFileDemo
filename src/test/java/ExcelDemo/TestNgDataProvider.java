package ExcelDemo;

import Utility.XLUtility;
import io.github.bonigarcia.wdm.WebDriverManager;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.IOException;

public class TestNgDataProvider {
    private WebDriver d;


    @BeforeMethod
    public void inIt(){
        WebDriverManager.chromedriver().setup();
        d=new ChromeDriver();
    }

    @Test (dataProvider = "loginData")
    public void testBrowserLaunch(String userName, String password, String isValidData) throws InterruptedException {
//        System.out.println("UserName : " +   userName +"PassWord :  "  +password +"isValidData :" +   isValidData);
        d.get("https://opensource-demo.orangehrmlive.com/");
        WebElement email = d.findElement(By.xpath("//*[@id=\"txtUsername\"]"));
        email.clear();
        email.sendKeys(userName);
        WebElement ePassword = d.findElement(By.xpath("//*[@id=\"txtPassword\"]"));
        ePassword.clear();
        ePassword.sendKeys("Password");

        WebElement loginBtn = d.findElement(By.xpath("//*[@id=\"btnLogin\"]"));
        loginBtn.click();

        String expectedUrl = "https://opensource-demo.orangehrmlive.com/index.php/dashboard";
        String actualUrl =d.getCurrentUrl();
        if (isValidData.equals("Valid")){
            Assert.assertTrue(expectedUrl.equals(actualUrl));
            d.findElement(By.xpath("//*[@id=\"welcome\"]")).click();
            Thread.sleep(3000);
            d.findElement(By.xpath("//*[@id=\"welcome-menu\"]/ul/li[3]/a")).click();

        }
        else if(isValidData.equals("Invalid")){
            Assert.assertTrue(!expectedUrl.equals(actualUrl));
        }

    }
    @DataProvider(name= "loginData")
    public Object[][] getData() throws IOException {
        /*String loginData[][]= {
                {"Admin", "admin123", "valid"},
                {"Admin", "admin", "invalid"},
                {"Admin", "adm", "Invalid"},
                {"Adm", "admin123", "Invalid"}
        };*/
        String path = System.getProperty("user.dir") + "\\Datasource\\Data.xlsx";
        XLUtility xlUtils = new XLUtility(path);
        int totalRows = xlUtils.getRowCount("Sheet1");
        System.out.println("Row : "+ totalRows);
        int totalColumns= xlUtils.getCellCount("Sheet1", 1);
        String loginData[][] = new String[totalRows][totalColumns];
        for (int row = 0; row < totalRows; row++){
            for (int col=0; col<totalColumns; col++){
                loginData[row][col]= xlUtils.getCellData("Sheet1",row,col);
                System.out.println("Value :" + loginData[row][col] +  "  ");
            }
            System.out.println();
        }
        return loginData;
    }


}

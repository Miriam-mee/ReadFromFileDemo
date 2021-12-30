package ExcelDemo;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.FileReader;

public class JsonDemo {
    private WebDriver d;
@BeforeMethod
    public void setUp() {
        WebDriverManager.chromedriver().setup();
        d = new ChromeDriver();

    }

    @Test(dataProvider = "dp")
    public void testJsonDataProvider(Object[] data){
        d.get("https://www.saucedemo.com/");
        String user[]= new String[2];
        for(Object o:data){
            String user_pwd = ((String) o);
            System.out.println(("Miriam :"+(String) o));
            user[0]= user_pwd.split(",")[0];
            user[1]= user_pwd.split(",")[1];

        }
    }

    @DataProvider(name= "dp")
    public Object[][] readJson(){
        JSONParser parser = new JSONParser();
        Object obj = null;

        try{
            FileReader fileReader = new FileReader(".\\Datasource\\logindetails.json");
            obj=parser.parse(fileReader);
        }
        catch (Exception e){
            e.printStackTrace();
        }
        JSONObject jsonObject = (JSONObject) obj;
        Object userdetails = jsonObject.get("userdetails");
        JSONArray jsonArray = (JSONArray) userdetails;
        String userLogin[]= new String[jsonArray.size()];
        for(int i = 0; i<jsonArray.size(); i++){
            Object o = jsonArray.get(i);
            JSONObject jsonObject1 =(JSONObject) o;
            String userName= (String) jsonObject1.get("user");
            String password =(String)jsonObject1.get("password");

            System.out.println("User : "+ userName);
            System.out.println("Password :" + password);
            userLogin[i]=userName + ", " + password;
        }
        return  new Object[][]{userLogin};


    }
}
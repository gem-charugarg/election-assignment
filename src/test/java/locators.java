import org.openqa.selenium.By;

public class locators {
    public static By stateSelector = By.id("ddlState");

    public static By allStates = By.xpath("//select[@id='ddlState']//option");
    public static By constituencySelector = By.id("ddlAC");

    public static By allConstituencies = By.xpath("//select[@id='ddlAC']//option");

    public static By tableRows = By.xpath("//div[@id='div1']//tbody//tr");

    public static By tableCols = By.xpath("//div[@id='div1']//tbody//tr//th");


    public static By colName(int row,int col){
        return By.xpath("//div[@id='div1']//tbody//tr[" + row + "]//td[" + col + "]");
    }
}

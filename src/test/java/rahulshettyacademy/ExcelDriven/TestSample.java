package rahulshettyacademy.ExcelDriven;

import java.io.IOException;
import java.util.ArrayList;

import org.openqa.selenium.By;

public class TestSample {

	public static void main(String[] args) throws IOException {
		
		DataDriven d = new DataDriven();
		ArrayList data = d.getData("Add Profile");
		
		System.out.println(data.get(0));
		System.out.println(data.get(1));
		System.out.println(data.get(2));
		System.out.println(data.get(3));
		
		//Implementation example;
		//driver.findElement(By.xpath("cgjhjgy")).sendKeys(data.get(0));

	}

}

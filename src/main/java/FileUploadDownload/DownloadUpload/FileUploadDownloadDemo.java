package FileUploadDownload.DownloadUpload;

import java.time.Duration;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

import com.google.common.collect.Table.Cell;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.DataFormatter;

public class FileUploadDownloadDemo {
	public static void main(String[] args) throws IOException {
		WebDriver driver = new FirefoxDriver();
		String fileName = "C:\\\\Selenium\\\\download.xlsx";
		DataFormatter formate = new DataFormatter();
		int numOfRows;
		int numOfColumns;
		XSSFSheet sheet;
		String formattedcell = null;
		Iterator<Row> rowIterator;
		String columnNUmOfFruitName = null;
		int celValueofPrice = 0;
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(7));
//	System.setProperty("webdriver.gecko.driver", "C:\\Selenium\\Drivers\\geckodriver.exe");
//	driver.get("https://rahulshettyacademy.com/upload-download-test/");
//	driver.findElement(By.id("downloadButton")).click();
		FileInputStream fis = new FileInputStream(fileName);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int numOfSheets = workbook.getNumberOfSheets();
		for (int i = 0; i < numOfSheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("Sheet1")) {
				numOfRows = workbook.getSheetAt(i).getPhysicalNumberOfRows();
				sheet = workbook.getSheetAt(i);
				rowIterator = sheet.rowIterator();
				String fruitPrice = null;

				while (rowIterator.hasNext()) {
					Row row = rowIterator.next();
					numOfColumns = row.getPhysicalNumberOfCells();

					for (int j = 0; j < numOfColumns; j++) {
						String headerCellName = row.getCell(j).getStringCellValue();
						if (headerCellName.equalsIgnoreCase("Price")) {
							celValueofPrice = j;
							break;
						}
					}
					break;
				}

				while (rowIterator.hasNext()) {
					Row row = rowIterator.next();
					numOfColumns = row.getPhysicalNumberOfCells();

					for (int j = 0; j < numOfColumns; j++) {
						formattedcell = formate.formatCellValue(row.getCell(j));
						if (formattedcell.equalsIgnoreCase("Apple")) {

							fruitPrice = formate.formatCellValue(row.getCell(celValueofPrice));
							break;
						}
					}
					if (formattedcell.equalsIgnoreCase("Apple")) {
						break;
					}
				}
				System.out.println("Fruit Price" + fruitPrice);

			}

		}
//	WebElement upload =  driver.findElement(By.cssSelector("input[type='file']"));
//	upload.sendKeys(fileName);
//	//driver.findElement(By.xpath("//div[@class ='Toastify__toast-body' div[2]]")).sendKeys("learning");
//	//driver.findElement(By.className("Toastify__toast-body div:nth-child(2"));
//	WebDriverWait wait =  new WebDriverWait(driver,Duration.ofSeconds(5));
//	wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(".Toastify__toast-body div:nth-child(2)")));
//     
//	String message = driver.findElement(By.cssSelector(".Toastify__toast-body div:nth-child(2)")).getText();
//	
//	System.out.println(message);
//	Assert.assertEquals(message,"Updated Excel Data Successfully.");
//	wait.until(ExpectedConditions.invisibilityOfElementLocated(By.cssSelector(".Toastify__toast-body div:nth-child(2)")));
//    String priceColumn = driver.findElement(By.xpath("//div[text()='Price']")).getAttribute("data-column-id");
//	
//  String vegPrice = driver.findElement(By.xpath("//div[@class='sc-hIPBNq eXWrwD rdt_TableBody']/div[@id='row-1']/div[@id='cell-"+priceColumn+"-undefined']")).getText();
//System.out.println(vegPrice);
//Assert.assertEquals("320",vegPrice);

	}
}
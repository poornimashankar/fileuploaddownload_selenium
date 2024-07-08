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

import org.apache.poi.ss.usermodel.Cell;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.DataFormatter;

public class FileUploadDownloadDemo1 {
	public static void main(String[] args) throws IOException {
		String fruitName = "Apple";
		String fileName = "C:\\\\Selenium\\\\download.xlsx";
		String updatedPrice = "310";
		System.setProperty("webdriver.gecko.driver", "C:\\Selenium\\Drivers\\geckodriver.exe");
		WebDriver driver = new FirefoxDriver();
		driver.get("https://rahulshettyacademy.com/upload-download-test/");
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(7));

		FileInputStream fis = new FileInputStream(fileName);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		DataFormatter formate = new DataFormatter();
		// Download the file
		driver.findElement(By.id("downloadButton")).click();
		String fruitPrice = getFruitPrice(workbook, fruitName, formate,updatedPrice);
		System.out.println("Fruit price before updated" + fruitPrice);
		// Edit the fruit Price
		// Upload the file
       FileOutputStream  fos =  new FileOutputStream(fileName);
       workbook.write(fos);
  workbook.close();
       fis.close();
       
		WebElement upload = driver.findElement(By.cssSelector("input[type='file']"));
		upload.sendKeys(fileName);
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(5));
		// Wait until message appears
		wait.until(ExpectedConditions
				.visibilityOfElementLocated(By.cssSelector(".Toastify__toast-body div:nth-child(2)")));

		String message = driver.findElement(By.cssSelector(".Toastify__toast-body div:nth-child(2)")).getText();

		System.out.println(message);
		Assert.assertEquals(message, "Updated Excel Data Successfully.");
		// Wait until message disappears
		wait.until(ExpectedConditions
				.invisibilityOfElementLocated(By.cssSelector(".Toastify__toast-body div:nth-child(2)")));
		String priceColumn = driver.findElement(By.xpath("//div[text()='Price']")).getAttribute("data-column-id");

		String priceAfterUpdated = driver
				.findElement(By.xpath("//div[@class='sc-hIPBNq eXWrwD rdt_TableBody']/div[@id='row-1']/div[@id='cell-"
						+ priceColumn + "-undefined']"))
				.getText();
		System.out.println("Fruit price after updated" + priceAfterUpdated);
		Assert.assertEquals(updatedPrice, priceAfterUpdated);

	}

	private static String getFruitPrice(XSSFWorkbook workbook, String fruitName, DataFormatter formate,String updatedPrice) {
		int numOfRows;
		int numOfColumns;
		XSSFSheet sheet;
		Iterator<Row> rowIterator;
		int celValueofPrice = 0;
		String CellName = null;
		String priceBeforeUpdated = null;
		int numOfSheets = workbook.getNumberOfSheets();
		for (int i = 0; i < numOfSheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("Sheet1")) {
				numOfRows = workbook.getSheetAt(i).getPhysicalNumberOfRows();
				sheet = workbook.getSheetAt(i);
				rowIterator = sheet.rowIterator();

				while (rowIterator.hasNext()) {
					Row row = rowIterator.next();
					numOfColumns = row.getPhysicalNumberOfCells();

					for (int j = 0; j < numOfColumns; j++) {
						CellName = formate.formatCellValue(row.getCell(j));
						if (CellName.equalsIgnoreCase("Price")) {
							celValueofPrice = j;
							break;
						} else if (CellName.equalsIgnoreCase(fruitName)) {
							priceBeforeUpdated = formate.formatCellValue(row.getCell(celValueofPrice));
							row.getCell(celValueofPrice).setCellValue(updatedPrice);
							break;

						}
					}
					if (CellName.equalsIgnoreCase("Apple")) {
						break;
					}
				}

			}

		}

		return priceBeforeUpdated;
	}
}
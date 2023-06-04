package Excel;

import java.io.File;
import java.io.FileInputStream;
import java.time.Duration;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

public class Main extends ExcelWriter {

	static int WRows;
	static String formattedPrice;

	public static void main(String[] args) throws Exception {

		FileInputStream excellFile = new FileInputStream(new File(FileLocation));

		XSSFWorkbook workbook = new XSSFWorkbook(excellFile);

		XSSFSheet sheet1 = workbook.getSheet("Sheet1");

		int firstRow = sheet1.getFirstRowNum();
		int lastRow = sheet1.getLastRowNum();

		for (int i = firstRow + 1; i <= lastRow; i++) {
			System.out.print("___________________________");
			System.out.println("\nValidating Row " + i);
			System.out.println("___________________________");
			XSSFRow row = sheet1.getRow(i);
			WRows = i;
			byRows(row);

		}

	}

	public static void CheckingFmiNum(double nUMPrice, double fMIPrice) throws Exception {

		int statusCol = 0;
		double RPrice = Double.parseDouble(formattedPrice);
		if (nUMPrice == RPrice && fMIPrice == RPrice) {
			String result = "Both are Correct";
			System.out.println(result);
			ExcelWriter.updateCellData(WRows, statusCol, result);
			System.out.println(" Both Correct Cell data updated successfully in excel");
		} else if (nUMPrice == RPrice) {
			String result = "NUM Correct";
			System.out.println(result);
			ExcelWriter.updateCellData(WRows, statusCol, result);
			System.out.println("NUM data updated successfully in excel.");
		} else if (fMIPrice == RPrice) {

			String result = "FMI Correct";
			System.out.println(result);
			ExcelWriter.updateCellData(WRows, statusCol, result);
			System.out.println("FMI data updated successfully in excel.");
		} else if (nUMPrice != RPrice && fMIPrice != RPrice) {

			String result = "Both are Wrong";
			System.out.println(result);
			ExcelWriter.updateCellData(WRows, statusCol, result);
			System.out.println("Wrong Cell data updated successfully in excel.");
		} else {
			System.out.println("No Condition Satisfy");
		}
	}

	public static void byRows(XSSFRow row1) throws Exception {
		// Taking Retailer name from Excel
		XSSFCell retailer = row1.getCell(2);
		String retailName = getRetailerName(retailer);

		//Taking WebLink from Excel
		XSSFCell cell1 = row1.getCell(9);
		CellsData(cell1, retailName);

		// Taking NUMPrice from Excel
		XSSFCell NumCell = row1.getCell(6);
		double NUMPrice = getCellValue(NumCell);

		System.out.println("Num price " + NUMPrice);

		// Taking FMIPrice from Excel
		XSSFCell FMICell = row1.getCell(4);
		double FMIPrice = getCellValue(FMICell);
		System.out.println("FMI price " + FMIPrice);
		
		//Compare the NUM price and FMI price with Webpage price
		CheckingFmiNum(NUMPrice, FMIPrice);

	}

	public static void CellsData(XSSFCell cell1, String retailVal) throws Exception {
		String name = null;
		if (cell1.getCellType() == HSSFCell.CELL_TYPE_STRING) {
			name = cell1.getStringCellValue();
			System.out.println(name);
		}
		getPrice(name, retailVal);

	}

	public static String getRetailerName(XSSFCell cell1) {
		String name = null;

		if (cell1.getCellType() == HSSFCell.CELL_TYPE_STRING) {
			name = cell1.getStringCellValue();
			System.out.println(name);
		}

		return name;

	}

	public static double getCellValue(XSSFCell cell1) {
		double val = 0;

		if (cell1.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
			val = cell1.getNumericCellValue();
			// System.out.println(val);
		}

		return val;

	}

	public static void getPrice(String URL, String retailColumnVal) throws Exception {

		if (retailColumnVal.startsWith("AJ")) {

			String xpath = "//div[contains(@class, 'font-size-xxl') and contains(@class, 'ml1') and contains(@class, 'bold') ]/span";
			getPriceAction(URL, xpath);

		} else if (retailColumnVal.startsWith("Amazon")) {

			String xpath1 = "//span[@class='a-price aok-align-center']";
			String xpath2 = "//*[@id='mbc-price-1']";
			amazon(URL, xpath1, xpath2);
			//two types of websites(two xpath) in amazon
		} else if (retailColumnVal.startsWith("Home")) {
			HomeDepot(URL);

		} else if (retailColumnVal.startsWith("Lowes")) {

			String xpath = "//div[contains(@class,'main-price undefined')]";
			getPriceAction(URL, xpath);

		} else if (retailColumnVal.startsWith("Wayfair")) {

			String xpath = "//*[@id='bd']/div[1]/div[2]/div/div[2]/div/div[1]/div[2]/div[1]/div[1]/span[1]";
			getPriceAction(URL, xpath);

			// Price Temporary Unavailabale
		} else if (retailColumnVal.startsWith("MyKnobs")) {

			String xpath = "//span[@class='price']";
			getPriceAction(URL, xpath);

		} else if (retailColumnVal.startsWith("Overstock")) {

			String xpath = "//div[@class='css-1olsk4d e1eyx97t2']";
			getPriceAction(URL, xpath);

		} else if (retailColumnVal.startsWith("Walmart")) {

			String xpath = "//span[@class='inline-flex flex-column']";
			getPriceAction(URL, xpath);

		} else if (retailColumnVal.startsWith("Light")) {

			String xpath = "(//span[contains(@class,'sales')])[1]";
			getPriceAction(URL, xpath);

		}

	}

	public static void priceFormatter(String priceText) throws Exception {

		String priceValue = priceText.replaceAll("[^0-9]", "");

		formattedPrice = priceValue.substring(0, priceValue.length() - 2) + "."
				+ priceValue.substring(priceValue.length() - 2);
		System.out.println("Price value: " + formattedPrice);
		int PriceColumn = 8;
		ExcelWriter.updateCellData(WRows, PriceColumn, formattedPrice);
		System.out.println("Updated successfully in Excel");

		// System.out.println(priceValue);

	}

	public static void getPriceAction(String URL, String Xpath) {
		try {
			ChromeOptions options = new ChromeOptions();
			options.addArguments("--remote-allow-origins=*");
			WebDriver driver = new ChromeDriver(options);
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(60));
			driver.get(URL);

			Thread.sleep(5000);
			WebElement AJ = driver.findElement(By.xpath(Xpath));
			String r = AJ.getText();
			System.out.println("Prie in Web " + r);
			priceFormatter(r);

			driver.quit();
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}

	}

	public static void HomeDepot(String URL) throws Exception {
		ChromeOptions options = new ChromeOptions();
		options.addArguments("--remote-allow-origins=*");
		WebDriver driver = new ChromeDriver(options);
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(60));
		driver.get(URL);

		String pinCode = "55441";
		try {

			driver.findElement(By.xpath("//*[@id='myStore']/a/span/div[2]/div")).click();
			Thread.sleep(2000);
			driver.findElement(By.xpath("//*[@id='myStoreDropdown']/div/div[4]/a")).click(); // *[@id='myStoreDropdown']/div/div[4]/a
			driver.findElement(By.id("myStore-formInput")).sendKeys(pinCode);
			driver.findElement(By.xpath("//*[@id='myStore-formButton']/span/img")).click();
			driver.findElement(By.xpath("//*[@id='myStore-list']/div[1]/div[6]/button")).click();

			WebElement Home = driver.findElement(
					By.xpath("//*[@id='root']/div/div[3]/div/div/div[3]/div/div/div[1]/div/div/div/div/div[1]"));

			String r = Home.getText(); // Last step it didn't take text price
			System.out.println(r);
			priceFormatter(r);

		} catch (NoSuchElementException e) {
			System.out.println("Lower Price Cart Scenario");
		} catch (Exception e1) {
			System.out.println(e1.getMessage());
		}

		driver.quit();
	}

	public static void amazon(String URL, String xpath1, String xpath2) throws Exception {
		ChromeOptions options = new ChromeOptions();
		options.addArguments("--remote-allow-origins=*");
		WebDriver driver = new ChromeDriver(options);
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(60));
		driver.get(URL);
		try {

			Thread.sleep(5000);
			WebElement AJ = driver.findElement(By.xpath(xpath1));
			String r = AJ.getText();
			System.out.println("Prie in Web " + r);
			priceFormatter(r);

		} catch (Exception e) {
			try {
				Thread.sleep(1000);
				WebElement AJ = driver.findElement(By.xpath(xpath2));
				String r = AJ.getText();
				System.out.println("Prie in Web " + r);
				priceFormatter(r);
			} catch (Exception e1) {
				System.out.println(e1.getMessage());
			}
		}

		driver.quit();

	}

}



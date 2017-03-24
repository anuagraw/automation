import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import java.util.ArrayList;
import java.util.List;

 
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 
/*
 * Create Excel file and insert data from other excels.
 */
public class MyTestClass {
 
	//folder location for input .xlsx files
	static String inputFileLocation= "D:\\Users\\anuagraw\\Downloads\\excelNew\\";
	
	//parameters for judgment
	static int totalJudgementParameters = 24;
	
	//variables to write data on to the final evaluation excel
	static String outputFileName ="D:\\Users\\anuagraw\\Downloads\\excelNew\\NewFolder\\JavaTesting.xlsx";
	
	static int columnsToLEaveBeforeScoreIns=3;
	
	static short borderType = CellStyle.BORDER_THIN;
	static short fontForHeaders = Font.BOLDWEIGHT_BOLD;
	static short borderColor = IndexedColors.BLACK.getIndex();
	
	static int insertionStartRow = 5;
	static int insertionStartColumn = 1;
	
	static int mergeStartColumn = 3;
 
	public static void main(String args[]) throws Exception {
		
		List<String> valueOfparameters = new ArrayList<String>();
		List<String> listOfParameters = new ArrayList<>();
		List<String> listOfHeadings = new ArrayList<>();
		
		listOfParameters.add("ProjectName");
		listOfParameters.add("TeamStrength");
		
		listOfParameters.add("PnPCurrent");
		listOfParameters.add("SCCurrent");
		listOfParameters.add("DnTCurrent");
		listOfParameters.add("BnRCurrent");
		listOfParameters.add("DeCurrent");
		listOfParameters.add("MoCurrent");
		listOfParameters.add("CMCurrent");
		listOfParameters.add("CurrentScore");
		
		listOfParameters.add("PnPExpected");
		listOfParameters.add("SCExpected");
		listOfParameters.add("DnTExpected");
		listOfParameters.add("BnRExpected");
		listOfParameters.add("DeExpected");
		listOfParameters.add("MoExpected");
		listOfParameters.add("CMExpected");
		listOfParameters.add("ExpectedScore");
		
		listOfParameters.add("DF");
		listOfParameters.add("LT");
		listOfParameters.add("MTTR");
		listOfParameters.add("TD");
		listOfParameters.add("PFD");
		listOfParameters.add("CTV");
		
		
		listOfHeadings.add("SNo");
		listOfHeadings.add("Project Name");
		listOfHeadings.add("Team Strength");
		listOfHeadings.add("Process And Planning");
		listOfHeadings.add("Source Control");
		listOfHeadings.add("Develop And Test");
		listOfHeadings.add("Build And Release");
		listOfHeadings.add("Deploy");
		listOfHeadings.add("Monitoring");
		listOfHeadings.add("Control/Measures");
		listOfHeadings.add("Average Current Score");
		listOfHeadings.add("Process And Planning");
		listOfHeadings.add("Source Control");
		listOfHeadings.add("Develop And Test");
		listOfHeadings.add("Build And Release");
		listOfHeadings.add("Deploy");
		listOfHeadings.add("Monitoring");
		listOfHeadings.add("Control/Measures");
		listOfHeadings.add("Average Expected Score");
		listOfHeadings.add("Deployment Frequency (releases per year)");
		listOfHeadings.add("Mean Time to Deploy/ Lead Time (in weeks)");
		listOfHeadings.add("Mean Time to Recover (in hours)");
		listOfHeadings.add("Technical Debt (in days)");
		listOfHeadings.add("Percent of failed deployments (in last one year)");
		listOfHeadings.add("Customer ticket volume (in last one year)");
		
		System.out.println("Start");
		
		File dir = new File(inputFileLocation);
		List<File> excelFiles = new ArrayList<File>();
		for (File fileMatch : dir.listFiles()) {
		    if (fileMatch.getName().endsWith((".xlsx"))) {
		      excelFiles.add(fileMatch);
		    }
		  }
		
		for(int individualFile=0;individualFile<excelFiles.size();individualFile++)
		{
			
			FileInputStream file = new FileInputStream(excelFiles.get(individualFile).getPath());
		    //XSSFWorkbook workbook = new XSSFWorkbook(file);
		   // Workbook workbook = WorkbookFactory.create(file);
			Workbook workbook = WorkbookFactory.create(file, "devops2017");
		    for(int judgementParam=0; judgementParam<totalJudgementParameters; judgementParam++)
		    {
		    	int namedCellIdx = workbook.getNameIndex(listOfParameters.get(judgementParam));
			    Name aNamedCell = workbook.getNameAt(namedCellIdx);
			    AreaReference aref = new AreaReference(aNamedCell.getRefersToFormula());
			    CellReference[] crefs = aref.getAllReferencedCells();
			    // retrieve the cell at the named range and test its contents
			    
			    Sheet sheet = workbook.getSheet(crefs[0].getSheetName());
			    Row row = sheet.getRow(crefs[0].getRow());
			    Cell cell = row.getCell(crefs[0].getCol());
			    if(judgementParam==0)
			    {
			    	valueOfparameters.add(cell.getStringCellValue());	
			    }
			    else
			    {
			    	double number = Math.round(cell.getNumericCellValue() * 100);
					number = number/100;
				    valueOfparameters.add(String.valueOf(number));	
			    }			    
		    }
		    file.close();
		}
		
		Workbook wb = new XSSFWorkbook();
		Sheet sheetToCreate = wb.createSheet("Total Score");
	    FileOutputStream fileOut = new FileOutputStream(outputFileName);
	    
	    CellStyle style = wb.createCellStyle();
	    CellStyle boldStyle = wb.createCellStyle();
	    Font font = wb.createFont();//Create font
	    font.setBoldweight(fontForHeaders);//Make font bold
	    style.setBorderBottom(borderType);
	    style.setBottomBorderColor(borderColor);
	    style.setBorderLeft(borderType);
	    style.setLeftBorderColor(borderColor);
	    style.setBorderRight(borderType);
	    style.setRightBorderColor(borderColor);
	    style.setBorderTop(borderType);
	    style.setTopBorderColor(borderColor);
	    
	    boldStyle.setBorderBottom(borderType);
	    boldStyle.setBottomBorderColor(borderColor);
	    boldStyle.setBorderLeft(borderType);
	    boldStyle.setLeftBorderColor(borderColor);
	    boldStyle.setBorderRight(borderType);
	    boldStyle.setRightBorderColor(borderColor);
	    boldStyle.setBorderTop(borderType);
	    boldStyle.setTopBorderColor(borderColor);
	    boldStyle.setFont(font);
	    
	    int headerCheck=0;
	    for(int totalRows=insertionStartRow; totalRows<=(excelFiles.size()+insertionStartRow+1); totalRows++)
	    {
	    	Row row = sheetToCreate.createRow(totalRows);
	    	for(int eachCell=insertionStartColumn; eachCell<(listOfHeadings.size()+insertionStartColumn); eachCell++)
	    	{
	    		if(headerCheck==0 || headerCheck==1)
	    		{
	    			Cell cell = row.createCell(eachCell);
				    cell.setCellStyle(boldStyle);
	    		}
	    		else
	    		{
	    			Cell cell = row.createCell(eachCell);
				    cell.setCellStyle(style);	
	    		}	    		
	    	}
	    	++headerCheck;    		
	    }
	   	    /*Merge columns for current score*/
	    Row rowToMerge = sheetToCreate.getRow(insertionStartRow);
	    Cell cellMergeCurrent = rowToMerge.getCell(insertionStartColumn+columnsToLEaveBeforeScoreIns);
	    cellMergeCurrent.setCellValue("Current Score");
	    
	    Cell cellMergeExpected = rowToMerge.getCell(insertionStartColumn+columnsToLEaveBeforeScoreIns+7+1);
	    cellMergeExpected.setCellValue("Expected Score");
	    
	    Cell cellMergeKPI = rowToMerge.getCell(insertionStartColumn+columnsToLEaveBeforeScoreIns+7+1+7+1);
	    cellMergeKPI.setCellValue("KPI");
	    
	    sheetToCreate.addMergedRegion(new CellRangeAddress(
	    		insertionStartRow, //first row (0-based)
	    		insertionStartRow, //last row  (0-based)
	    		(insertionStartColumn+columnsToLEaveBeforeScoreIns), //first column (0-based)
	    		(insertionStartColumn+columnsToLEaveBeforeScoreIns+7)  //last column  (0-based)
	    ));
	    CellUtil.setAlignment(cellMergeCurrent, wb, CellStyle.ALIGN_CENTER);
	    
   	    /*Merge columns for expected score*/
	    //rowToMerge = sheetToCreate.getRow(insertionStartRow);
	    
	    
	    sheetToCreate.addMergedRegion(new CellRangeAddress(
	    		insertionStartRow, //first row (0-based)
	    		insertionStartRow, //last row  (0-based)
	    		(insertionStartColumn+columnsToLEaveBeforeScoreIns+7+1), //first column (0-based)
	    		(insertionStartColumn+columnsToLEaveBeforeScoreIns+7+1+7)  //last column  (0-based)
	    ));
	    CellUtil.setAlignment(cellMergeExpected, wb, CellStyle.ALIGN_CENTER);
	    
	    /*Merge columns for KPI scores*/
	    
	    sheetToCreate.addMergedRegion(new CellRangeAddress(
	    		insertionStartRow, //first row (0-based)
	    		insertionStartRow, //last row  (0-based)
	    		(insertionStartColumn+columnsToLEaveBeforeScoreIns+7+1+7+1), //first column (0-based)
	    		(insertionStartColumn+columnsToLEaveBeforeScoreIns+7+1+7+1+5)  //last column  (0-based)
	    ));
	    CellUtil.setAlignment(cellMergeKPI, wb, CellStyle.ALIGN_CENTER);
	    
	    
	    Row rowForHeader = sheetToCreate.getRow(insertionStartRow+1);
	    int insertionStartColumnCopy=insertionStartColumn;
	    for(int writeHeadings=0;writeHeadings<listOfHeadings.size();writeHeadings++)
	    {
	    	Cell cellForHeader = rowForHeader.getCell(insertionStartColumnCopy++);
	    	cellForHeader.setCellValue(listOfHeadings.get(writeHeadings));
	    }	    
	    
	    
	    System.out.println("------------------------------------");
	    int columnTracker=0;
	    for(int fetchIndividualProjectScore=0; fetchIndividualProjectScore< excelFiles.size(); fetchIndividualProjectScore++)
		{
	    	insertionStartColumnCopy=insertionStartColumn;
	    	Row rowToInsert = sheetToCreate.createRow(insertionStartRow+fetchIndividualProjectScore+2);
	    	Cell cellToWrite = rowToInsert.createCell(insertionStartColumnCopy++);
	    	cellToWrite.setCellValue(fetchIndividualProjectScore+1);
	    	cellToWrite.setCellStyle(style);
	    	//cellToWrite = rowToInsert.createCell(insertionStartColumnCopy++);
	    	/*cellToWrite.setCellValue(excelFiles.get(fetchIndividualProjectScore).getName());
	    	cellToWrite.setCellStyle(style);*/

	    	for(int i=0;i<totalJudgementParameters;i++)
	    	{
	    		if(i!=0)
	    		{
	    			cellToWrite = rowToInsert.createCell(insertionStartColumnCopy++);
		    		cellToWrite.setCellValue(Double.parseDouble(valueOfparameters.get(columnTracker++)));
		    		cellToWrite.setCellStyle(style);
	    		}
	    		else
	    		{
	    			cellToWrite = rowToInsert.createCell(insertionStartColumnCopy++);
		    		cellToWrite.setCellValue(valueOfparameters.get(columnTracker++));
		    		cellToWrite.setCellStyle(style);	
	    		}
	    	}
		}
	    
	    for(int eachCell=insertionStartColumn; eachCell<=listOfHeadings.size(); eachCell++)
    	{
	    	sheetToCreate.autoSizeColumn(eachCell);
	    }
	    
	    wb.write(fileOut);
	    fileOut.close();
	}
}
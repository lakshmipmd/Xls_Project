import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

//Read the contents from the excel file (with header) & write the same in another excel file
public class XLSXReaderWriter5 {

	private static final File excel = 
			new File("C:\\Users\\pradp\\Documents\\pdp_fileproject_output\\sample1.xls");
	private static final File outExcel = 
			new File("C:\\Users\\pradp\\Documents\\pdp_fileproject_output\\sampleOutput1.xls");
	
	public static void main(String[] args) throws IOException{
	
		List<ExcelHeader> contentHeader = getContentHeaderfromExcel();
		List<ExcelFile> contentList = getContentListfromExcel();
		System.out.println(contentHeader);
		System.out.println(contentList);
		writeContentListToExcel(contentList, contentHeader);
		
	}

	private static void writeContentListToExcel(List<ExcelFile> contentList, List<ExcelHeader> contentHeader) throws IOException {

		FileOutputStream fos = new FileOutputStream(outExcel);
		HSSFWorkbook outbook = new HSSFWorkbook();
		HSSFSheet sheet = outbook.createSheet("Sheet1");
		
		int rowIndex = 0;		//Start of row index
		
		HSSFRow row = sheet.createRow(rowIndex++);
		
		int HeaderIndex = 0;	
		for(ExcelHeader header : contentHeader){
			row.createCell(HeaderIndex++).setCellValue(header.getHeader());
		}

		for(ExcelFile content : contentList){
			row = sheet.createRow(rowIndex++);		//Starts from row 1
			int cellIndex = 0;
			//first place in row is ID
			row.createCell(cellIndex++).setCellValue(content.getID());
			//second place in row is name
			row.createCell(cellIndex++).setCellValue(content.getName());
			//third place in row is marks in Salary
			row.createCell(cellIndex++).setCellValue(content.getSalary());
			//fourth place in row is marks in Department
			row.createCell(cellIndex++).setCellValue(content.getDepartment());
			//fifth place in row is marks in Manager
			row.createCell(cellIndex++).setCellValue(content.getManager());
			}
			outbook.write(fos);
		 
			 System.out.println(outExcel + " is successfully written");
			 
			 outbook.close();
			 fos.close();
			 		
	}
	
	private static List<ExcelHeader> getContentHeaderfromExcel() throws IOException {
		
		List<ExcelHeader> contentHeader = new ArrayList<>();
		
		FileInputStream file = new FileInputStream(excel);
		HSSFWorkbook book = new HSSFWorkbook(file);
		HSSFSheet sheet = book.getSheetAt(0);
			
		// Assuming "column headers" are in the first row
		HSSFRow header_row = sheet.getRow(0);
		int col = (int)header_row.getLastCellNum();
		
		// Iterating over each column header
		for (int i = 0; i < col; i++) {
			ExcelHeader content = new ExcelHeader();
			HSSFCell header_cell = header_row.getCell(i);
			content.setHeader(header_cell.getStringCellValue());
			contentHeader.add(content);
			}
		
		book.close();
		file.close();
		return contentHeader;
	}

	private static List<ExcelFile> getContentListfromExcel() throws IOException{
		
		List<ExcelFile> contentList = new ArrayList<>();
		
		FileInputStream file = new FileInputStream(excel);
		HSSFWorkbook book = new HSSFWorkbook(file);
		HSSFSheet sheet = book.getSheetAt(0);
		HSSFRow row = sheet.getRow(1);
			
		//Iterate through each rows one by one
        for(int i = 1 ; i < row.getLastCellNum(); i++){
           	HSSFRow nextRow = sheet.getRow(i);
           	ExcelFile content = new ExcelFile();

           	for(int j = 0 ; j < nextRow.getLastCellNum() ; j++){
           		HSSFCell nextCell = nextRow.getCell(j);
                	
               	if(nextCell.getColumnIndex() == 0)
                	content.setID(nextCell.getNumericCellValue());
                else if(nextCell.getColumnIndex() == 1)
                	content.setName(nextCell.getStringCellValue());
                else if(nextCell.getColumnIndex() == 2)
                	content.setSalary(nextCell.getStringCellValue());
                else if(nextCell.getColumnIndex() == 3)
                	content.setDepartment(nextCell.getStringCellValue());
                else if(nextCell.getColumnIndex() == 4)
                	content.setManager(nextCell.getStringCellValue());
                	
                }//end of for loop of cells
               contentList.add(content);
              }//End of for loop of rows
			book.close();
			file.close();
			
		return contentList;
	}
}

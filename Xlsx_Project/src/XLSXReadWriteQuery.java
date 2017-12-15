import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

//Program to read Excel file & Create Insert query & Write into text file

public class XLSXReadWriteQuery {

	private static final File excel = new File("C:\\Users\\pradp\\Documents\\pdp_fileproject_output\\sample1.xls");
	private static final File outFile = new File("C:\\Users\\pradp\\Documents\\pdp_fileproject_output\\query.txt");
	private static final String table_name = "SYS.SAMPLE_TABLE";
	
	public static void main(String[] args) {

		try {
			//get the header from the excel file
			List<ExcelHeader> contentHeader = getContentHeaderfromExcel();
			System.out.println(contentHeader);
			
			//get the row values from the excel file
			List<ExcelFile> contentList = getContentListfromExcel();
			System.out.println(contentList);
			
			//create insert query & write into text file
			writeContentListToText(contentList, contentHeader);

		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	//create insert query & write into text file
	private static void writeContentListToText(List<ExcelFile> contentList, List<ExcelHeader> contentHeader)
			throws IOException {
		
		FileWriter fw = new FileWriter(outFile);;
		BufferedWriter bw = new BufferedWriter(fw);
		
		// Prepare the column Names of the table from the file header - if the names r same
		String Col_Name = "";

		for (ExcelHeader header : contentHeader) {
			Col_Name += header.toString() + ", ";
		}
		String Column_Names = Col_Name.substring(0, Col_Name.length() - 2);

		// Prepare the table values from the excel file & Prepare final query
		
		for (ExcelFile content : contentList) {
			String Query = "(" + content.getID() + ", '" + content.getName() + "', '" + content.getSalary()
					+ "', '" + content.getDepartment() + "', '" + content.getManager() + "')";

			String contentquery = "INSERT INTO " + table_name + "(" + Column_Names + ")  VALUES " 
									+ Query + "\n" ;
			System.out.println(contentquery);
			bw.write(contentquery);
		}

		bw.close();
		fw.close();
	}

	//get the header from the excel file
	private static List<ExcelHeader> getContentHeaderfromExcel() throws IOException {

		List<ExcelHeader> contentHeader = new ArrayList<>();

		FileInputStream file = new FileInputStream(excel);
		HSSFWorkbook book = new HSSFWorkbook(file);
		HSSFSheet sheet = book.getSheetAt(0);

		// Assuming "column headers" are in the first row
		HSSFRow header_row = sheet.getRow(0);
		int col = (int) header_row.getLastCellNum();

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

	//get the row values from the excel file
	private static List<ExcelFile> getContentListfromExcel() throws IOException {

		List<ExcelFile> contentList = new ArrayList<ExcelFile>();

		FileInputStream file = new FileInputStream(excel);
		HSSFWorkbook book = new HSSFWorkbook(file);
		HSSFSheet sheet = book.getSheetAt(0);

		// Iterate through each rows one by one
		for (int i = 1; i <= sheet.getLastRowNum(); i++) {
			HSSFRow nextRow = sheet.getRow(i);
			ExcelFile content = new ExcelFile();
			int cellsize = nextRow.getPhysicalNumberOfCells();

			for (int j = 0; j < cellsize; j++) {
				HSSFCell nextCell = nextRow.getCell(j);

				if (nextCell.getColumnIndex() == 0)
					content.setID(nextCell.getNumericCellValue());
				else if (nextCell.getColumnIndex() == 1)
					content.setName(nextCell.getStringCellValue());
				else if (nextCell.getColumnIndex() == 2)
					content.setSalary(nextCell.getStringCellValue());
				else if (nextCell.getColumnIndex() == 3)
					content.setDepartment(nextCell.getStringCellValue());
				else if (nextCell.getColumnIndex() == 4)
					content.setManager(nextCell.getStringCellValue());

			} // end of for loop of cells
			contentList.add(content);
		} // End of for loop of rows
		book.close();
		file.close();
		return contentList;
	}
}

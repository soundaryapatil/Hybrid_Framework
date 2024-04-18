package utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFileUtil {
	Workbook wb;
	//constructor
	public ExcelFileUtil(String ExcelPath)throws Throwable
	{
		FileInputStream fi = new FileInputStream(ExcelPath);
		wb =WorkbookFactory.create(fi);
	}
	//count no of rows in a sheet
		public int rowCount(String sheetName)
		{
			return wb.getSheet(sheetName).getLastRowNum();
		}
		//method for reading cell data
		public String getCellData(String SheetName,int row,int column)
		{
			String data;
			if(wb.getSheet(SheetName).getRow(row).getCell(column).getCellType()==CellType.NUMERIC)
			{
				int celldata =(int)wb.getSheet(SheetName).getRow(row).getCell(column).getNumericCellValue();
				data =String.valueOf(celldata);
			}
			else
			{
				data = wb.getSheet(SheetName).getRow(row).getCell(column).getStringCellValue();
			}
			return data;
		}
		//method for set cell data
		public void setCellData(String sheetName,int row,int columns,String status,String WriteExcelpath)throws Throwable
		{
			//get sheet from wb
			Sheet ws =wb.getSheet(sheetName);
			//get row from sheet
			Row rowNum =ws.getRow(row);
			//create cell in row
			Cell cell =rowNum.createCell(columns);
			//write status
			cell.setCellValue(status);
			if(status.equalsIgnoreCase("Pass"))
			{
				CellStyle style = wb.createCellStyle();
				Font font = wb.createFont();
				//color with green
				font.setColor(IndexedColors.GREEN.getIndex());
				font.setBold(true);
				style.setFont(font);
				ws.getRow(row).getCell(columns).setCellStyle(style);
			}
			else if(status.equalsIgnoreCase("Fail"))
			{
				CellStyle style = wb.createCellStyle();
				Font font = wb.createFont();
				// color with red
				font.setColor(IndexedColors.RED.getIndex());
				font.setBold(true);
				style.setFont(font);
				ws.getRow(row).getCell(columns).setCellStyle(style);
			}
			else if(status.equalsIgnoreCase("Blocked"))
			{
				CellStyle style = wb.createCellStyle();
				Font font = wb.createFont();
				// color with blue
				font.setColor(IndexedColors.BLUE.getIndex());
				font.setBold(true);
				style.setFont(font);
				ws.getRow(row).getCell(columns).setCellStyle(style);
			}
			FileOutputStream fo =new FileOutputStream(WriteExcelpath);
			wb.write(fo);
		}
		public static void main(String[] args) throws Throwable {
			ExcelFileUtil xl = new ExcelFileUtil("D://sample.xlsx");
			
			int rc = xl.rowCount("EMP");
			System.out.println(rc);
			for(int i=1;i<=rc;i++)
			{
				String fname = xl.getCellData("EMP", i, 0);
				String mname = xl.getCellData("EMP", i, 1);
				String lname = xl.getCellData("EMP", i, 2);
				String eid = xl.getCellData("EMP", i, 3);
				System.out.println(fname+" "+mname+" "+lname+" "+eid);
				//xl.setCellData("EMP", i, 4, "pass", "D:/sound.xlsx");
				//xl.setCellData("EMP", i, 4, "Fail", "D:/sound.xlsx");
				xl.setCellData("EMP", i, 4, "Blocked", "D:/sound.xlsx");
			}
		}
}





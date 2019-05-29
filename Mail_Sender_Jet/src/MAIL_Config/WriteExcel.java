package MAIL_Config;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Path;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

public class WriteExcel
{
	static String ExcelSavePath;
	static HSSFWorkbook hswb;

	public String foldercreate(String folderpath) throws IOException
	{
		Boolean f = DirectoryStream.class.equals(folderpath);
		String foldercheck = folderpath;

		if (f != true) {
			System.out.println("Exists");
			File f2 = new File(folderpath);
			f2.mkdir();
		} else if (f == true) {
			System.out.println("NOT Exists");

		}
		String ExcelSavePath = autoexcel(foldercheck);
		File f1 = new File(foldercheck);
		f1.listFiles();
		return ExcelSavePath;

	}

	@SuppressWarnings("resource")
	public String autoexcel(String foldercheck) throws IOException {
		Path ExcelSavePath = null;
		int fCount = 0;

		File f2 = new File(foldercheck);
		fCount = f2.listFiles().length;
		if (fCount != 0) {

			ExcelSavePath = java.nio.file.Paths.get(foldercheck, "RUN_" + (fCount + 1) + ".xls");
			(ExcelSavePath).toString();
			HSSFWorkbook hswb;
			hswb = new HSSFWorkbook();
			ExcelSavePath = java.nio.file.Paths.get(foldercheck, "RUN_" + (fCount + 1) + ".xls");
			String b = (ExcelSavePath).toString();
			HSSFSheet sheet1 = hswb.createSheet("Run 1");

			// Main code //
			HSSFFont font = hswb.createFont();
			((org.apache.poi.ss.usermodel.Font) font).setFontHeightInPoints((short) 9);
			((org.apache.poi.ss.usermodel.Font) font).setBold(true);
			HSSFRow row1 = sheet1.createRow(0);

			// XSSFCell cell = row1.getCell(0);
			HSSFCell cell = row1.createCell(0);
			CellStyle style9 = hswb.createCellStyle();
			
		
			cell = row1.createCell(0);
			cell.setCellValue("Customer Email ID");
			cell.setCellStyle(style9);
		 
			cell = row1.createCell(1);
			cell.setCellValue("Status");
			cell.setCellStyle(style9);
			
			cell = row1.createCell(2);
			cell.setCellValue("Time");
			cell.setCellStyle(style9);
			
			FileOutputStream fileOut14 = new FileOutputStream(b);
			hswb.write(fileOut14);
			fileOut14.close();
		} 
		else
		{

			HSSFWorkbook hswb;

			hswb = new HSSFWorkbook();

			ExcelSavePath = java.nio.file.Paths.get(foldercheck, "RUN_1" + ".xls");
			String a = (ExcelSavePath).toString();
			HSSFSheet sheet1 = hswb.createSheet("Run 1");

			HSSFRow row1 = sheet1.createRow(0);

			HSSFCell cell = row1.createCell(0);
			CellStyle style9 = hswb.createCellStyle();
			
			cell = row1.createCell(0);
			cell.setCellValue("Customer Email ID");
			cell.setCellStyle(style9);

			cell = row1.createCell(1);
			cell.setCellValue("Status");
			cell.setCellStyle(style9);

			FileOutputStream fileOut12 = new FileOutputStream(a);
			hswb.write(fileOut12);
			fileOut12.close();

		}
		String d = (ExcelSavePath).toString();
		return d;

	}

	
	public void WriteToExcel(String Receiver_Email_ID,String Status,int r, String filename,String Time) throws FileNotFoundException, IOException
	{
		HSSFWorkbook hswb;
		int col = 0;

		FileInputStream fileOut13 = new FileInputStream(filename);
		hswb = new HSSFWorkbook(fileOut13);
		fileOut13.close();

		HSSFSheet sheet1 = hswb.getSheetAt(0);
		// HSSFSheet sheet2 = hswb.getSheetAt(1);
		HSSFRow row22 = sheet1.createRow(r);
		row22.createCell(col);
		HSSFCell cell1;

		CellStyle style3 = hswb.createCellStyle();
		HSSFFont font = hswb.createFont();
		cell1 = sheet1.getRow(r).createCell(col);
		cell1.setCellValue(Receiver_Email_ID);
		cell1.setCellStyle(style3);
		style3.setFont(font);

		
		Cell c;
		CellStyle style2 = hswb.createCellStyle();
		HSSFFont font1 = hswb.createFont();
		if (Status.equals("Pass")) {

			 
			style2.setFillForegroundColor(IndexedColors.LIME.index);
			style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			c = sheet1.getRow(r).createCell(col + 1);
			c.setCellValue("Sent");
			c.setCellStyle(style2);
			font1.setFontName(HSSFFont.FONT_ARIAL);
			font1.setFontHeightInPoints((short) 10);
			font1.setColor(IndexedColors.BLACK.getIndex());
			style2.setFont(font1);

		} 
		else

		{
			 
			style2.setFillForegroundColor(IndexedColors.RED.index);
			style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			c = sheet1.getRow(r).createCell(col + 1);
			c.setCellValue("Not Sent");
			c.setCellStyle(style2);
			font.setFontName(HSSFFont.FONT_ARIAL);
			font.setFontHeightInPoints((short) 10);
			font.setColor(IndexedColors.BLACK.getIndex());
			style2.setFont(font);
		}

		CellStyle style31 = hswb.createCellStyle();
		HSSFFont font11 = hswb.createFont();
		cell1 = sheet1.getRow(r).createCell(col+2);
		cell1.setCellValue(Time);
		cell1.setCellStyle(style31);
		style31.setFont(font11);
		
		
		FileOutputStream fileOut12 = new FileOutputStream(filename);
		hswb.write(fileOut12);
		fileOut12.close();

	}


}

package com.exl;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {

	/**
	 * @param args
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	public static void main(String[] args) throws FileNotFoundException,
			IOException {
		// TODO Auto-generated method stub

		XSSFWorkbook xwb = null;
		HSSFWorkbook wb = null;
		boolean flag = true;
		try {
			POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(
					"c://成绩表1.xls"));
			wb = new HSSFWorkbook(fs);
		} catch (Exception e) {
			flag = false;
			xwb = new XSSFWorkbook(new FileInputStream("c://成绩表.xlsx"));
			
		}
		if (flag) {//2003
			HSSFSheet sheet = wb.getSheetAt(0);
			int rows = sheet.getPhysicalNumberOfRows();
			for (int i = 0; i < rows; i++) {
				HSSFRow row = sheet.getRow(i);
				int cells = row.getPhysicalNumberOfCells();
				for (int j = 0; j < cells; j++) {
					HSSFCell cell = row.getCell(j);
					cell.setCellType(Cell.CELL_TYPE_STRING);
					System.out.print(cell.getStringCellValue() + "    ");
				}
				System.out.println();
			}
		} else {//2007
			XSSFSheet sheet = xwb.getSheetAt(0);
			int rows = sheet.getPhysicalNumberOfRows();
			for (int i = 0; i < rows; i++) {
				XSSFRow row = sheet.getRow(i);
				int cells = row.getPhysicalNumberOfCells();
				for (int j = 0; j < cells; j++) {
					XSSFCell cell = row.getCell(j);
					cell.setCellType(Cell.CELL_TYPE_STRING);
					System.out.print(cell.getStringCellValue() + "    ");
				}
				System.out.println();
			}
		}

	}

}

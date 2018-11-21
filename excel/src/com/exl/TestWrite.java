package com.exl;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestWrite {
	/**
	 * @param args
	 * @throws IOException 
	 * @throws FileNotFoundException 
	 */
	public static void main(String[] args) throws FileNotFoundException, IOException {
		// TODO Auto-generated method stub
		try{
			HSSFWorkbook wb = new HSSFWorkbook();//创建一个excel文件
			HSSFSheet sheet1 = wb.createSheet("first");
			HSSFSheet sheet2 = wb.createSheet("second");
			HSSFRow row1 = sheet1.createRow(0);
			HSSFCell cell1 = row1.createCell(0);
			HSSFCell cell2 = row1.createCell(1);
			cell1.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
			cell1.setCellValue(1000);
			cell2.setCellValue(new Date());
			wb.write(new FileOutputStream("f://成绩表2.xls"));
			System.out.println("输出是2003版本");
		}catch (Exception e) {
			XSSFWorkbook xwb = new XSSFWorkbook();
			XSSFSheet sheet1= xwb.createSheet("first");
			XSSFSheet sheet2= xwb.createSheet("second");
			XSSFRow row1 = sheet1.createRow(0);
			XSSFCell cell1 = row1.createCell(0);
			cell1.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
			cell1.setCellValue(1000);
			xwb.write(new FileOutputStream("f://成绩表2.xls"));
			System.out.println("输出是2007版本");
		}
	}

}

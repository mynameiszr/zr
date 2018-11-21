package com.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcle {
	public static void main(String[] args) throws FileNotFoundException,
			IOException {
		XSSFWorkbook xwb = null;//支持2007版本的excle
		HSSFWorkbook wb = null;//支持2003版本的excle
		boolean flag = true;
		try {
			//这是2003版本的获取excle文件方式
			POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(
					"d://test2003.xls"));
			wb = new HSSFWorkbook(fs);
		} catch (Exception e) {
			//这是2007版本的获取excle方式
			flag = false;
			xwb = new XSSFWorkbook(new FileInputStream("d://test.xlsx"));
			// TODO: handle exception
		}
		if (flag) {// 2003
			HSSFSheet sheet = wb.getSheetAt(0);//读取第一个sheet
			int rows = sheet.getPhysicalNumberOfRows();//获取有效行数
			for (int i = 0; i < rows; i++) {
				HSSFRow row = sheet.getRow(i);//获取每一行
				int cells = row.getPhysicalNumberOfCells();//后去行中的列数
				for (int j = 0; j < cells; j++) {
					HSSFCell cell = row.getCell(j);//得到每个列
					System.out.print(cell.getStringCellValue() + "    ");//输出列的值
				}
				System.out.println();
			}
		} else {// 2007
			XSSFSheet sheet = xwb.getSheetAt(0);
			int rows = sheet.getPhysicalNumberOfRows();
			for (int i = 0; i < rows; i++) {
				XSSFRow row = sheet.getRow(i);
				int cells = row.getPhysicalNumberOfCells();
				for (int j = 0; j < cells; j++) {
					XSSFCell cell = row.getCell(j);
					System.out.print(cell.getStringCellValue() + "    ");
				}
				System.out.println();
			}
		}

	}
}

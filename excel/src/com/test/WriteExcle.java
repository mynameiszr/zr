package com.test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcle {
	public static void main(String[] args) throws FileNotFoundException, IOException {
			//这是2003版本输出
			HSSFWorkbook wb = new HSSFWorkbook();//创建excle对象
			HSSFSheet sheet1 = wb.createSheet("first");//创建第一个sheet
			HSSFSheet sheet2 = wb.createSheet("second");//创建第二个sheet
			HSSFRow row1 = sheet1.createRow(0);//在第一个sheet中创建第一行
			HSSFCell cell1 = row1.createCell(0);//在第一行中创建第一列
			cell1.setCellType(HSSFCell.CELL_TYPE_NUMERIC);//设置第一列的列数据类型
			cell1.setCellValue(1000);//给第一列设置值
			wb.write(new FileOutputStream("c://成绩表.xlsx"));//输出为文件
			System.out.println("输出是2003版本");
			
			//这是2007版本输出
//			XSSFWorkbook xwb = new XSSFWorkbook();
//			XSSFSheet sheet1= xwb.createSheet("first");
//			XSSFSheet sheet2= xwb.createSheet("second");
//			XSSFRow row1 = sheet1.createRow(0);
//			XSSFCell cell1 = row1.createCell(0);
//			cell1.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
//			cell1.setCellValue(1000);
//			xwb.write(new FileOutputStream("d://11.xlsx"));
//			System.out.println("输出是2007版本");
	}
}

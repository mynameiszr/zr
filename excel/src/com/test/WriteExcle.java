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
			//����2003�汾���
			HSSFWorkbook wb = new HSSFWorkbook();//����excle����
			HSSFSheet sheet1 = wb.createSheet("first");//������һ��sheet
			HSSFSheet sheet2 = wb.createSheet("second");//�����ڶ���sheet
			HSSFRow row1 = sheet1.createRow(0);//�ڵ�һ��sheet�д�����һ��
			HSSFCell cell1 = row1.createCell(0);//�ڵ�һ���д�����һ��
			cell1.setCellType(HSSFCell.CELL_TYPE_NUMERIC);//���õ�һ�е�����������
			cell1.setCellValue(1000);//����һ������ֵ
			wb.write(new FileOutputStream("c://�ɼ���.xlsx"));//���Ϊ�ļ�
			System.out.println("�����2003�汾");
			
			//����2007�汾���
//			XSSFWorkbook xwb = new XSSFWorkbook();
//			XSSFSheet sheet1= xwb.createSheet("first");
//			XSSFSheet sheet2= xwb.createSheet("second");
//			XSSFRow row1 = sheet1.createRow(0);
//			XSSFCell cell1 = row1.createCell(0);
//			cell1.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
//			cell1.setCellValue(1000);
//			xwb.write(new FileOutputStream("d://11.xlsx"));
//			System.out.println("�����2007�汾");
	}
}

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
		XSSFWorkbook xwb = null;//֧��2007�汾��excle
		HSSFWorkbook wb = null;//֧��2003�汾��excle
		boolean flag = true;
		try {
			//����2003�汾�Ļ�ȡexcle�ļ���ʽ
			POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(
					"d://test2003.xls"));
			wb = new HSSFWorkbook(fs);
		} catch (Exception e) {
			//����2007�汾�Ļ�ȡexcle��ʽ
			flag = false;
			xwb = new XSSFWorkbook(new FileInputStream("d://test.xlsx"));
			// TODO: handle exception
		}
		if (flag) {// 2003
			HSSFSheet sheet = wb.getSheetAt(0);//��ȡ��һ��sheet
			int rows = sheet.getPhysicalNumberOfRows();//��ȡ��Ч����
			for (int i = 0; i < rows; i++) {
				HSSFRow row = sheet.getRow(i);//��ȡÿһ��
				int cells = row.getPhysicalNumberOfCells();//��ȥ���е�����
				for (int j = 0; j < cells; j++) {
					HSSFCell cell = row.getCell(j);//�õ�ÿ����
					System.out.print(cell.getStringCellValue() + "    ");//����е�ֵ
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

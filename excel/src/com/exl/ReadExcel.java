package com.exl;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.PrintWriter;


import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ReadExcel extends HttpServlet {

	/**
	 * Constructor of the object.
	 */
	public ReadExcel() {
		super();
	}

	/**
	 * Destruction of the servlet. <br>
	 */
	public void destroy() {
		super.destroy(); // Just puts "destroy" string in log
		// Put your code here
	}

	
	public void doGet(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
	}


	public void doPost(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream("d://test.xlsx"));
		HSSFWorkbook wb = new HSSFWorkbook(fs);
		HSSFSheet sheet = wb.getSheetAt(0);
		int rows = sheet.getPhysicalNumberOfRows();
		for(int i=0;i<rows;i++){
			HSSFRow row = sheet.getRow(i);
			int cells = row.getPhysicalNumberOfCells();
			for(int j=0;j<cells;j++){
				HSSFCell cell = row.getCell(j);
				System.out.print(cell.getStringCellValue()+"    ");
			}
			System.out.println();
		}
	}	
}

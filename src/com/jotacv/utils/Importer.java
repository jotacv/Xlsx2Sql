package com.jotacv.utils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Importer {
	
	private List<List<Cell>> sheetData;
	private XSSFWorkbook workbook;
	private String tableName;
	private int maxwidth = 0;
	private FileOutputStream fos;
	
	private static String getStringFromCell(Cell cell,String def,boolean quoted, Boolean isId) {
		String value = def;
		if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
			if(quoted)
				value = "'"+cell.getStringCellValue().replace("\n", "\\n").replace("'", "\'")+"'";
			else
				value = cell.getStringCellValue();
		} else {
			if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
				if (isId)
					value = String.valueOf((int)cell.getNumericCellValue());
				else
					value = String.valueOf(cell.getNumericCellValue());
			}
		}
		return value;
	}
	
	public Importer (String filename) {
		try{
			
			this.sheetData = new ArrayList<List<Cell>>();
			InputStream is = new FileInputStream(filename);
			this.workbook = new XSSFWorkbook(is);
			XSSFSheet sheet = this.workbook.getSheetAt(0);
			this.tableName = sheet.getSheetName();
			
			System.out.print("Got the file... ");
			this.maxwidth = sheet.getRow(0).getLastCellNum();
			Iterator<Row> rows = sheet.rowIterator();
			while (rows.hasNext()) {
				XSSFRow row = (XSSFRow) rows.next();
				List<Cell> data = new ArrayList<Cell>();
				for (int i = 0; i < maxwidth; i++) {
					XSSFCell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					data.add(cell);
				}
				sheetData.add(data);
			}
		}catch(Exception e){
			e.printStackTrace();
		}finally{
			System.out.print("parsed.\n");
			try {
				this.workbook.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}
	
	public void generate(){
		//Get Headers
		StringBuilder sbh = new StringBuilder();
		sbh.append("(");
		for(int i=0;i<this.maxwidth;i++){
			List<Cell> firstRow = this.sheetData.get(0);
			sbh.append(getStringFromCell(firstRow.get(i),"",false,false));
			if(i<this.maxwidth-1)
				sbh.append(",");
		}
		sbh.append(")");
		
		//Output
		System.out.print("Generating.");
		try{
			int k = 0;
			this.fos = new FileOutputStream("output.sql");
			Iterator<List<Cell>> it = this.sheetData.iterator();
			it.next();	//Skip first;
			while(it.hasNext()){
				if(k++%50==0)System.out.print(".");
				List<Cell> row = it.next();
				StringBuilder sb = new StringBuilder();
				sb.append("INSERT INTO ");
				sb.append(this.tableName);
				sb.append(sbh);
				sb.append(" VALUES (");
				for(int i=0;i<this.maxwidth;i++){
					sb.append(getStringFromCell(row.get(i),"null",true,i==0));
					if(i<this.maxwidth-1)
						sb.append(",");
				}
				sb.append(");\n");
				fos.write(sb.toString().getBytes());
			}
		}catch(Exception e){
			e.printStackTrace();
		}finally{
			System.out.print("done.\n");
			try {
				this.fos.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		
	}

	public static void main(String[] args) {
		Importer importer = new Importer(args[0]);
		importer.generate();
	}

}

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

public class Converter {
	private List<List<Cell>> sheetData;
	private XSSFWorkbook workbook;
	private String tableName;
	private int maxwidth = 0;
	private FileOutputStream fos;
	private Integer sheetNum; 
	private String nullValue = "null";
	private StringBuilder sbh = null;
	private StringBuilder sbv = null;
	private StringBuilder sbt = null;
	
	private boolean OMMIT_NULL_VALUES = true;
	
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
	
	public Converter (InputStream is, int sheetNum) {
		try{

			this.sheetNum = sheetNum;
			this.sheetData = new ArrayList<List<Cell>>();

			this.workbook = new XSSFWorkbook(is);
			XSSFSheet sheet = this.workbook.getSheetAt(sheetNum);
			this.tableName = sheet.getSheetName();
			
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
		}catch(IllegalArgumentException iae){
			//Cannot iterate to next sheet
			throw iae;
		}catch(Exception e){
			e.printStackTrace();
		}finally{
			try {
				this.workbook.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}
	
	public void run(){
		try{
			System.out.print("Sheet "+sheetNum+" ");
			//Get Headers Map	
			List<String> headers = new ArrayList<String>();
			for(int i=0;i<this.maxwidth;i++){
				List<Cell> firstRow = this.sheetData.get(0);
				headers.add(getStringFromCell(firstRow.get(i),"",false,false));
			}
			
			//Iterate trough rows
			int k = 0;
			this.fos = new FileOutputStream("output"+this.sheetNum+".sql");
			Iterator<List<Cell>> it = this.sheetData.iterator(); it.next();	//Skip first;			
			while(it.hasNext()){
				
				//Clean
				if(k++%50==0)System.out.print(".");
				List<Cell> row = it.next();
				sbh = new StringBuilder();
				sbv = new StringBuilder();
				sbt = new StringBuilder();
				
				//Get Values
				List<Integer> headersHits = new ArrayList<Integer>();
				String val = null;
				for(int i=0;i<this.maxwidth;i++){
					val = (getStringFromCell(row.get(i),this.nullValue,true,i==0));
					if(!OMMIT_NULL_VALUES || !this.nullValue.equals(val)){
						sbv.append(val);
						headersHits.add(i);
						sbv.append(",");
					}
				}
				if(sbv.length()>0)
					sbv.deleteCharAt(sbv.length()-1);
				
				//Get headers by value
				sbh.append("(");
				for(Integer headHit : headersHits){
					sbh.append(headers.get(headHit));
					sbh.append(",");
				}
				sbh.deleteCharAt(sbh.length()-1);
				sbh.append(")");
		
				//Build the final string
				sbt.append("INSERT INTO ");
				sbt.append(this.tableName);
				sbt.append(sbh);
				sbt.append(" VALUES (");
				sbt.append(sbv);
				sbt.append(");\n");
				fos.write(sbt.toString().getBytes());
			}
			
		}catch(Exception e){
			e.printStackTrace();
		}finally{
			System.out.println(". done");
			try {
				this.fos.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		
	}
} 

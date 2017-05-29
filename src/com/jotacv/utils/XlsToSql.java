package com.jotacv.utils;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;

public class XlsToSql {

	public static void main(String[] args) {
		if(args.length==1){
			InputStream is = null;
			
			try{
				int count = 0; 
				is = new FileInputStream(args[0]);
				for(;;){
					Converter converter = new Converter(is, count++);
					converter.run();
				}
			}catch(FileNotFoundException fnfe){
				System.out.println("Couldn't find the sheet or bad stuff happened.");
				fnfe.printStackTrace();
			
			}catch(Exception e){
				System.out.println("All done.");
				
			}finally{
				try{
					if(is!=null)
						is.close();
				}catch(Exception e){
					e.printStackTrace();
				}
			}
		}else{
			System.out.println("Usage java -jar Xls2Sql.jar [datasheet.xlsx]");
		}
	}

} 

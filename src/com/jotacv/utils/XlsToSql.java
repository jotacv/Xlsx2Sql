package com.jotacv.utils;

public class XlsToSql {

	public static void main(String[] args) {
		if(args.length==1){
			try{
				int count = 0; 
				for(;;){
					Converter converter = new Converter(args[0], count++);
					converter.run();
				}
			}catch(Exception e){
				System.out.println("All done");
			}
		}else{
			System.out.println("Usage java -jar Xls2Sql [datasheet.xlsx]");
		}
	}

}

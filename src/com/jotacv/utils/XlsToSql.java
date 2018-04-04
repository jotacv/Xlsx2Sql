package com.jotacv.utils;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XlsToSql {

	public static void main(String[] args) throws IOException {
		if (args.length == 1) {
			InputStream is = null;
			XSSFWorkbook wb = null;
			try {
				int count = 0;
				is = new FileInputStream(args[0]);
				wb = new XSSFWorkbook(is);
				for (;;) {
					Converter converter = new Converter(wb, count++);
					converter.run();
				}
			} catch (IllegalArgumentException iae) {
				System.out.println("All done.");

			} catch (FileNotFoundException fnfe) {
				System.out.println("Couldn't find the sheet ");
				fnfe.printStackTrace();

			} catch (IOException e) {
				e.printStackTrace();

			} finally {
				wb.close();
				is.close();
			}
		} else {
			System.out.println("Usage java -jar Xls2Sql.jar [datasheet.xlsx]");
		}
	}

}

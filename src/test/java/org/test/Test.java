package org.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.SystemOutLogger;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {
	public static void main(String[] args) throws IOException {
		
		File file=new File("C:\\Users\\vshar\\eclipse-workspace\\DayTwo\\Excel7\\Excel1.xlsx");
		FileInputStream stream=new FileInputStream(file);
		Workbook workbook=new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet("Sheet1");
		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
			//System.out.println(row);
			System.out.println("----------");
			for (int j = 0; j <row.getPhysicalNumberOfCells(); j++) {
				Cell cell = row.getCell(j);
				//System.out.println(cell);
				CellType type = cell.getCellType();
				//System.out.println(type);
				switch (type) {
				case STRING:
					String stringCellValue = cell.getStringCellValue();
					System.out.println(stringCellValue);
					break;
					//Modified Branch 
					
					

                case NUMERIC:
                	if (DateUtil.isCellDateFormatted(cell)) {
                		Date date = cell.getDateCellValue();
                		SimpleDateFormat dateFormat=new SimpleDateFormat("dd/MM/yyyy");
                		String format = dateFormat.format(date);
                		System.out.println(format);
						
					} else {
						double d = cell.getNumericCellValue();
						BigDecimal b = BigDecimal.valueOf(d);
						String number = b.toString();
						System.out.println(number);
						break;

					}
                	

				default:
					break;
					}
				
			
				}
					
				}
				}
			
		
		
		
	}
	

	


package com.attendance.interfaces;

import java.text.ParseException;

import org.apache.poi.ss.usermodel.Row;

public interface ImportProcessor {
	 void importFromExcel(String filePath);
	 
	 boolean validateFields(Row row);
	 
	 void persistEntity(Row row) throws ParseException;
	 
}

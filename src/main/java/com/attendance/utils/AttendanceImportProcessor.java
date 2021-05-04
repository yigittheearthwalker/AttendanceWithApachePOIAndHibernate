package com.attendance.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hibernate.Session;
import org.hibernate.Transaction;

import com.attendance.configuration.HibernateUtils;
import com.attendance.entities.Attendance;
import com.attendance.entities.Course;
import com.attendance.entities.Student;
import com.attendance.exceptions.InvalidFieldsRowException;
import com.attendance.interfaces.ImportProcessor;

public class AttendanceImportProcessor implements ImportProcessor{

	@Override
	public void importFromExcel(String filePath) {
		File file = new File(filePath);
		
		try (Workbook workbook = new XSSFWorkbook(new FileInputStream(file))){
			Sheet sheet = workbook.getSheet("take_attendance");
			Row fieldsRow = sheet.getRow(0);
			
			int lastRow = sheet.getLastRowNum();
			if (lastRow == -1) {
				throw new InvalidFieldsRowException("Fields Row does not exist");
			} 
			if (!validateFields(fieldsRow)) {
				throw new InvalidFieldsRowException("Template does not have correct fields or correct order");
			}
			
			for(int i = 1; i < lastRow + 1; i++) {
				
				Row row = sheet.getRow(i);
				
				try {
					persistEntity(row);
				} catch (ParseException e) {
					e.printStackTrace();
				}
				
			}
			HibernateUtils.getSessionFactory().close();
			workbook.close();
			
		} catch (FileNotFoundException e) {
			
			e.printStackTrace();
		} catch (IOException e) {
			
			e.printStackTrace();
		} catch (InvalidFieldsRowException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

	@Override
	public boolean validateFields(Row row) {
		boolean valid = true;
		
		for (int i = 0; i < TemplateGenerator.ATTENDANCE_COLUMNS.length; i++) {
			if (!row.getCell(i).getStringCellValue().equalsIgnoreCase(TemplateGenerator.ATTENDANCE_COLUMNS[i])) {
				valid = false;
				break;
			}
		}
		return valid;
	}
	
	@Override
	public void persistEntity(Row row) throws ParseException {
		
		Session session = HibernateUtils.getSessionFactory().openSession();
		Transaction transaction = session.beginTransaction();
		
		Cell dateCell = row.getCell(0);
		Cell studentCell = row.getCell(1);
		Cell courseCell = row.getCell(3);
		Cell attendedCell = row.getCell(5);
		
		Attendance attendance = new Attendance();
		attendance.setDate(new SimpleDateFormat("dd-MM-yyyy").parse(dateCell.getStringCellValue()));
		attendance.setStudent(session.get(Student.class, Integer.parseInt(studentCell.getStringCellValue())));
		attendance.setCourse(session.get(Course.class, Integer.parseInt(courseCell.getStringCellValue())));
		attendance.setAttended(attendedCell.getBooleanCellValue());
		
		session.persist(attendance);
		transaction.commit();
		session.close();
	}

}

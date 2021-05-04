package com.attendance.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hibernate.Session;
import org.hibernate.Transaction;

import com.attendance.configuration.HibernateUtils;
import com.attendance.entities.Course;
import com.attendance.exceptions.InvalidFieldsRowException;
import com.attendance.interfaces.ImportProcessor;

public class CourseImportProcessor implements ImportProcessor{

	@Override
	public void importFromExcel(String filePath) {
		File file = new File(filePath);
		try (Workbook workbook = new XSSFWorkbook(new FileInputStream(file))){
			
			Sheet sheet = workbook.getSheet("create_course");
			Row fieldsRow = sheet.getRow(0);

			int lastRow = sheet.getLastRowNum();
			if (lastRow == -1) {
				throw new InvalidFieldsRowException("Fields Row does not exist");
			} 
			if (!validateFields(fieldsRow)) {
				throw new InvalidFieldsRowException("Template does not have correct fields or correct order");
			}
			
			for (int i = 1; i < lastRow + 1; i++) {
				Row row = sheet.getRow(i);
				
				persistEntity(row);

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
		Cell cell0 = row.getCell(0);
		Cell cell1 = row.getCell(1);
		return cell0.getStringCellValue().equals("Course.id") && cell1.getStringCellValue().equals("Course.courseName");
	}

	@Override
	public void persistEntity(Row row) {
		Session session = HibernateUtils.getSessionFactory().openSession();
		Transaction transaction = session.beginTransaction();
		
		Cell cell0 = row.getCell(0);
		Cell cell1 = row.getCell(1);
		if (cell1 != null) {
			if (cell0 == null) {
				Course course = new Course();
				course.setCourseName(cell1.getStringCellValue());
				session.persist(course);
			}else {
				Course course = session.get(Course.class, (int) cell0.getNumericCellValue());
				course.setCourseName(cell1.getStringCellValue());
				session.persist(course);
			}
		}
		transaction.commit();
		session.close();
	}
	
}

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
import com.attendance.entities.Student;
import com.attendance.exceptions.InvalidFieldsRowException;
import com.attendance.interfaces.ImportProcessor;

public class StudentImportProcessor implements ImportProcessor{

	@Override
	public void importFromExcel(String filePath) {
		File file = new File(filePath);
		try (Workbook workbook = new XSSFWorkbook(new FileInputStream(file))){
			
			Sheet sheet = workbook.getSheet("create_student");
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
		Cell cell2 = row.getCell(2);
		return cell0.getStringCellValue().equals("Student.id") 
				&& cell1.getStringCellValue().equals("Student.studentName")
					&& cell2.getStringCellValue().equals("Student.course");
	}

	@Override
	public void persistEntity(Row row) {
		Session session = HibernateUtils.getSessionFactory().openSession();
		Transaction transaction = session.beginTransaction();
		
		Cell idCell = row.getCell(0);
		Cell nameCell = row.getCell(1);
		Cell courseCell = row.getCell(2);
		
		Student student = null; 
		if (idCell == null) {
			if (nameCell != null) {
				student = new Student();
				student.setStudentName(nameCell.getStringCellValue());
				if(courseCell != null) student.setCourse(session.get(Course.class,(int) courseCell.getNumericCellValue()));
				session.persist(student);
			}
		}else {
			if (nameCell != null || courseCell != null) {
				student = session.get(Student.class, (int) idCell.getNumericCellValue());
				if(nameCell != null) student.setStudentName(nameCell.getStringCellValue());
				if(courseCell != null) student.setCourse(session.get(Course.class,(int) courseCell.getNumericCellValue()));
				session.persist(student);
			}
		}
		
		transaction.commit();
		session.close();
	}

}

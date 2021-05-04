package com.attendance.utils;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.List;

import javax.persistence.Query;
import javax.xml.transform.Source;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hibernate.SQLQuery;
import org.hibernate.Session;
import org.hibernate.Transaction;
import org.hibernate.query.NativeQuery;

import com.attendance.configuration.HibernateUtils;
import com.attendance.entities.Student;

public class TemplateGenerator {
	final static String[] ATTENDANCE_COLUMNS = {"Attendance.date", "Attendance.student_id", "Student Name", 
			"Attendance.course_id", "Course Name", "Attendance.attended"};
	
	public static void generate(Class<?> clazz) {
		File file = new File(clazz.getSimpleName().toLowerCase()+"ImportTemplate.xlsx");
		
		try {
			Workbook workbook = new XSSFWorkbook();
			Sheet sheet = workbook.createSheet("create_"+clazz.getSimpleName().toLowerCase());
			Row row = sheet.createRow(0);
			Font fontBold = workbook.createFont();
			fontBold.setBold(true);
			CellStyle style = workbook.createCellStyle();
			style.setFont(fontBold);
			Field[] fields = clazz.getDeclaredFields();
			for (int i = 0; i < fields.length; i++) {
				Cell cell = row.createCell(i);
				cell.setCellValue(clazz.getSimpleName()+"."+fields[i].getName());
				cell.setCellStyle(style);
			}
	
			workbook.write(new FileOutputStream(file));
			workbook.close();
			System.out.println(file.getName() + " Ready ");
		} catch (IOException e1) {
			e1.printStackTrace();
		}
		
	}
	
	public static void generateAttendanceTemplate(String date, String courseName) {
		File file = new File("attendanceTemplate.xlsx");
		
		try {	
			Workbook workbook = new XSSFWorkbook();
			Sheet sheet = workbook.createSheet("take_attendance");
			Row row = sheet.createRow(0);
			
			Font fontBold = workbook.createFont();
			fontBold.setBold(true);
			CellStyle style = workbook.createCellStyle();
			style.setFont(fontBold);
			style.setWrapText(true);
			
			for (int i = 0; i < ATTENDANCE_COLUMNS.length; i++) {
				Cell cell = row.createCell(i);
				cell.setCellValue(ATTENDANCE_COLUMNS[i]);
				cell.setCellStyle(style);
			}
			
			fillAttendanceTemplate(sheet, date, courseName);
		
			workbook.write(new FileOutputStream(file));
			workbook.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	static void fillAttendanceTemplate(Sheet sheet, String date, String courseName) {
		Session session = HibernateUtils.getSessionFactory().openSession();
		Transaction transaction = session.beginTransaction();
		
		Query query = session.createNativeQuery("Select s.id, s.student_name, s.course_id "
														+ "from student s "
														+ "where s.course_id = (Select id "
																				+ "from course "
																				+ "where course_name = '"+courseName+"');");
		List<Object[]> resultRows = query.getResultList();
		for (int i = 0; i < resultRows.size(); i++) {
			Row row = sheet.createRow(i + 1);
			Object[] obj = resultRows.get(i);
		
			Cell cell1 = row.createCell(0);
			cell1.setCellValue(date);
			Cell cell2 = row.createCell(1);
			cell2.setCellValue(obj[0].toString());
			Cell cell3 = row.createCell(2);
			cell3.setCellValue(obj[1].toString());
			Cell cell4 = row.createCell(3);
			cell4.setCellValue(obj[2].toString());
			Cell cell5 = row.createCell(4);
			cell5.setCellValue(courseName);
		}
		
		session.close();
		transaction.commit();
		HibernateUtils.getSessionFactory().close();
	}
}

package com.attendance.utils;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import javax.persistence.Query;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hibernate.Session;
import org.hibernate.Transaction;

import com.attendance.configuration.HibernateUtils;

public class ReportGenerator {
	
	List<Object> dates = null;
	String courseName = null;
	Font fontBold = null;
	CellStyle style = null;
	
	
	public void attendanceReport(String courseName) {
		
		File file = new File("Attendance_Report.xlsx");
		this.courseName = courseName;
		dates = getDates();
		
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("Attendance_Table");
		setFontAndStyle(workbook);
		setHeaderRow(sheet);
		fillAttendanceSheet(sheet);
		try {
			workbook.write(new FileOutputStream(file));
			workbook.close();
			
			System.out.println(file.getName() + " Ready");
			HibernateUtils.getSessionFactory().close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
	List<Object> getDates(){
		Session session = HibernateUtils.getSessionFactory().openSession();
		Transaction transaction = session.beginTransaction();
		
		
		Query datesQuery = session.createNativeQuery("SELECT DISTINCT(DATE(a.attendance_date)) FROM attendance a");
		
		List<Object> dates = datesQuery.getResultList();
		
		transaction.commit();
		session.close();
		return dates;
	}
	
	void setHeaderRow(Sheet sheet) {
		Row row = sheet.createRow(0);
		
		Cell cell0 = row.createCell(0);
		cell0.setCellValue("Course Name");
		cell0.setCellStyle(style);
		Cell cell1 = row.createCell(1);
		cell1.setCellValue("Student Name");
		cell1.setCellStyle(style);
		
		for (int i = 2; i < dates.size() + 2; i++) {
			Cell dateCell = row.createCell(i);
			dateCell.setCellValue(dates.get(i - 2).toString());
			dateCell.setCellStyle(style);
		}
	}
	
	void fillAttendanceSheet(Sheet sheet) {
		Session session = HibernateUtils.getSessionFactory().openSession();
		Transaction transaction = session.beginTransaction();
		
		String queryString = "SELECT c.course_name, s.student_name, ";
		for (int i = 0; i < dates.size(); i++) {
			queryString += "(SELECT CASE WHEN a.attended = TRUE THEN 'X' ELSE '' END FROM attendance a "
							+ "WHERE a.student_id = s.id "
							+ "AND a.course_id = c.id "
							+ "AND DATE(a.attendance_date) = '" + dates.get(i).toString() + "') as \"day"+i+"\", ";
		}
		queryString = queryString.substring(0, queryString.length() - 2);
		queryString += "FROM course c "
					+ "INNER JOIN student s ON s.course_id = c.id "
					+ "WHERE c.course_name = '" + this.courseName + "' "
					+ "GROUP BY c.id, s.id; ";
				
		Query query = session.createNativeQuery(queryString);
		
		List<Object[]> attendanceData = query.getResultList();
		
		for (int i = 0; i < attendanceData.size(); i++) {
			Row row = sheet.createRow(i + 1);
			Object[] objects = attendanceData.get(i);
			for (int j = 0; j < objects.length; j++) {
				Cell cell = row.createCell(j);
				cell.setCellValue((String) objects[j]);
			}
		}		
	}
	
	void setFontAndStyle(Workbook workbook) {
		fontBold = workbook.createFont();
		fontBold.setBold(true);
		style = workbook.createCellStyle();
		style.setFont(fontBold);
	}
	
}

package com.attendance.main;

import java.io.File;
import java.io.FileOutputStream;
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
import com.attendance.interfaces.ImportProcessor;
import com.attendance.utils.AttendanceImportProcessor;
import com.attendance.utils.CourseImportProcessor;
import com.attendance.utils.ReportGenerator;
import com.attendance.utils.StudentImportProcessor;
import com.attendance.utils.TemplateGenerator;

public class Driver {
	public static void main(String[] args) {
		
		/* **STEP 1**
		 * This static method Creates import templates in Excel for given class in the parameter in the project directory
		 * After generating your templates, you can locate them and fill with your own data
		 */
		//
		
		TemplateGenerator.generate(Course.class);
		TemplateGenerator.generate(Student.class);
		
		/* **STEP 2**
		 * This importFromExcel() method looks for the excel file that is specified in the parameter as pathName. 
		 * The file must be specifically generated for importing the Course or Student entities and structure must not change.
		 * To create new entities, you need to leave ID column blank. If you specify any ID, processor will look
		 * for the specific entity in the Database and update it if finds one
		 * NOTE: when importing students, you must give course id in the course column, not the course name
		 */
		
//		 new CourseImportProcessor().importFromExcel("courseImportTemplate.xlsx");
//		 new StudentImportProcessor().importFromExcel("studentImportTemplate.xlsx");
		
		/* **STEP 3**
		 * After you persist your Course and Student entities using above methods, 
		 * 
		 */
		
//		TemplateGenerator.generateAttendanceTemplate("09-05-2021", "JAVA");
		
		/* **STEP 4**
		 * After generating the Attendance Template fill attended column with Boolean values
		 * To define whether if the participant attended on the specified date or not
		 *	Save the file and Run the Method Below
		 * 
		 */
//		 new AttendanceImportProcessor().importFromExcel("attendanceTemplate.xlsx");
		
		/* **STEP 5**
		 * After you do above steps and create attendances for a couple of days, 
		 * You can run below method to create an attendance report for a Course by giving the course name in the parameters
		 * 
		 */
//		 new ReportGenerator().attendanceReport("JAVA");
	}

}

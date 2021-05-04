package com.attendance.entities;

import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.Id;
import javax.persistence.JoinColumn;
import javax.persistence.OneToOne;
import javax.persistence.Table;

@Entity
@Table(name = "attendance")
public class Attendance {
	
	@Id
	@GeneratedValue
	private int id;
	
	@Column (name = "attendance_date")
	private Date date;
	
	@OneToOne
	@JoinColumn(name = "student_id")
	private Student student;
	
	@OneToOne
	@JoinColumn(name = "course_id")
	private Course course;
	
	@Column(name = "attended")
	private Boolean attended;

	public int getId() {
		return id;
	}

	public void setId(int id) {
		this.id = id;
	}

	public Date getDate() {
		return date;
	}

	public void setDate(Date date) {
		this.date = date;
	}

	public Student getStudent() {
		return student;
	}

	public void setStudent(Student student) {
		this.student = student;
	}

	public Course getCourse() {
		return course;
	}

	public void setCourse(Course course) {
		this.course = course;
	}

	public Boolean getAttended() {
		return attended;
	}

	public void setAttended(Boolean attended) {
		this.attended = attended;
	}


	
	
}

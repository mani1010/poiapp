package com.app.poi.service;

public class Student {

	private Integer usn;
	private String name;
	public Student() {
		super();
	}
	public Student(Integer usn, String name) {
		super();
		this.usn = usn;
		this.name = name;
	}
	public Integer getUsn() {
		return usn;
	}
	public void setUsn(Integer usn) {
		this.usn = usn;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	@Override
	public String toString() {
		return "Student [usn=" + usn + ", name=" + name + "]";
	}
	
	

}

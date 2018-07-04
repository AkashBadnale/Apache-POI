package com.ab.apachePOI.ExcelReadWrite;

public class Student {
   public int id;
   public String name;
   public String surname;
   public int age;
   
   public Student(int id, String name, String surname, int age) {
	super();
	this.id = id;
	this.name = name;
	this.surname = surname;
	this.age = age;
	}//param constructor

public int getId() {
	return id;
}

public void setId(int id) {
	this.id = id;
}

public String getName() {
	return name;
}

public void setName(String name) {
	this.name = name;
}

public String getSurname() {
	return surname;
}

public void setSurname(String surname) {
	this.surname = surname;
}

public int getAge() {
	return age;
}

public void setAge(int age) {
	this.age = age;
}
}//Student

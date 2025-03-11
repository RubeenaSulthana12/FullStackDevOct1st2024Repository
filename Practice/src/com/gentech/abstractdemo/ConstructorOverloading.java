package com.gentech.abstractdemo;
abstract class School
{
    School (String sclName, String sclLocation)
    {
        System.out.println("School Name :"+sclName);
        System.out.println("School Id:"+sclLocation);
    }
    void showCourseName(String courses[])
    {
        for(int i=0;i<courses.length;i++)
        {
            System.out.println("Courses Name :"+courses[i]);
        }
    }
}
class Students extends School
{
    Students(String sclName, String sclLocation)
    {
        super(sclName, sclLocation);
    }

    void showSchoolName(String schName)
    {
        System.out.println("School Name :"+schName);
    }
    void showSchoolId(String sclLocation)
    {
        System.out.println("School Id:"+sclLocation);
    }
    void showStudentName(String stuName)
    {
        System.out.println("Student Name :"+stuName);
    }
}


public class ConstructorOverloading {
    public static void main(String[] args)
    {
        Students s1=new Students("Gates","Banglore");
        s1.showSchoolId("Banglore");
        s1.showSchoolName("Gates");
        s1.showStudentName("Rubeena");
        s1.showCourseName(new String[]{"Java","Python","CSS"} );
    }

}



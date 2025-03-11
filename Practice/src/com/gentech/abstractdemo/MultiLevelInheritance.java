package com.gentech.abstractdemo;
abstract class School5
{
    abstract void showsclLocation(String location);
    abstract void showSchoolname (String sclName);
}
abstract class Student2 extends School5
{
    void showsclLocation(String location)
    {
        System.out.println("School Location :"+location);
    }

    void showStudentMail (String stuMail)
    {
        System.out.println("Student email:"+stuMail);
    }
}
class Student3 extends Student2
{
    void showSchoolname (String sclName)
    {
        System.out.println("School Name :"+sclName);
    }
    void showStudentId (int stuId)
    {
        System.out.println("Student Id :"+stuId);
    }
}

public class MultiLevelInheritance {
    public static void main(String[] args)
    {
        Student3 s1=new Student3();
        s1.showSchoolname("Gentech");
        s1.showsclLocation("Attigupe");
        s1.showStudentId(90);
        s1.showStudentMail("pulare@mail.com");
    }

}


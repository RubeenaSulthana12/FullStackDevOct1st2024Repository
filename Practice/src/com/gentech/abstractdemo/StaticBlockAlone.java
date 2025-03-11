package com.gentech.abstractdemo;
abstract class college
{
    static
    {
        System.out.println("it is a static block");
    }
    abstract void showCollegeName(String name);
    void showStudentName(String student[])
    {
        for(int i=0;i<student.length;i++)
        {
            System.out.println("Student Name :"+student[i]);
        }
    }
}
class RVcollege extends college
{
    void showCollegeName(String name)
    {
        System.out.println("College Name :"+name);
    }
    void showDepartmentName(String dName)
    {
        System.out.println("Department Name :"+dName);
    }
}



public class StaticBlockAlone{
    public static void main(String[] args) {
        RVcollege r1=new RVcollege();
        r1.showCollegeName("Gates College");
        r1.showDepartmentName("MCA");
        r1.showStudentName(new String[]{"Ruby","Fia","Ammu"});
    }

}



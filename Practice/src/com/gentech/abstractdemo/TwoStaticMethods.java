package com.gentech.abstractdemo;
abstract class College1
{
    abstract void showCollegeName(String cName);
    abstract void showCollegeId(int cId);
    void showCoursesName(String courses[])
    {
        for(int i=0;i<courses.length;i++)
        {
            System.out.println("Courses Name :"+courses[i]);
        }
    }
    static void showDisplay1()
    {
        System.out.println("It is a First static Method");
    }
    static void showDisplay2()
    {
        System.out.println("It is a Second static Method");
    }
}
class Teacher1 extends College1
{
    void showCollegeName(String cName)
    {
        System.out.println("College Name :"+cName);
    }
    void showCollegeId(int cId)
    {
        System.out.println("College Id :"+cId);
    }
    void showTeacherName(String tName)
    {
        System.out.println("Teacher Name :"+tName);
    }
}



public class TwoStaticMethods {
    public static void main(String[] args)
    {
        College1.showDisplay1();
        College1.showDisplay2();
        Teacher1 t1=new Teacher1();
        t1.showCollegeId(91);
        t1.showTeacherName("Rubeena");
        t1.showCollegeName("Gates");
        t1.showCoursesName(new String[]{"Civil","Mech","BT","Au"});
    }

}



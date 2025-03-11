//Execution order os static members
package com.gentech.Members;
class Student
{
    static String fullname;
    static int marks;
    static
    {
        fullname="ruby";
        showFullName();
        showMarks();
    }
    static void showFullName()
    {
        System.out.println("Full Name:"+fullname);
    }
    static void showMarks()
    {
        System.out.println("Marks:"+marks);
    }

}
public class StaticBlock1 {
    public static void main(String[] args) {
        Student.marks=77;
        System.out.println("marks in main method:"+Student.marks);
    }
}

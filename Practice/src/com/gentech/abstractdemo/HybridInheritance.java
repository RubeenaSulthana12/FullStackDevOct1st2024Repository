package com.gentech.abstractdemo;
abstract class Institute1
{
    abstract void showInstituteEmail(String mail);
    abstract void showTeacherName(String tName);
    abstract void showStudentFees(double fees);
    void showInstituteName(String insName)
    {
        System.out.println("Institute Name:"+insName);
    }
}
abstract class Institute2 extends Institute1
{
    void showInstituteEmail(String mail)
    {
        System.out.println("Institute Email :"+mail);
    }
    void showTeacherName(String tName)
    {
        System.out.println("Teacher Name :"+tName);
    }
    void showInstituteAddress (String insAdd)
    {
        System.out.println("Institute Address :"+insAdd);
    }
}
class Institute3 extends Institute1
{
    void showInstituteEmail(String mail)
    {
        System.out.println("Institute Email :"+mail);
    }
    void showTeacherName(String tName)
    {
        System.out.println("Teacher Name :"+tName);
    }
    void showStudentFees(double fees)
    {
        System.out.println("Student Fees :"+fees);
    }
    void showInstituteContact (long Contact)
    {
        System.out.println("Institute Number :"+Contact);
    }
}
class Institute4 extends Institute2
{
    void showStudentFees(double fees)
    {
        System.out.println("Student Fees :"+fees);
    }
    void showNumOfStuInInstitute(int stuNum)
    {
        System.out.println("Number of Students:"+stuNum);
    }
}

public class HybridInheritance
{
    public static void main(String[] args) {
        Institute3 i1=new Institute3();
        i1.showInstituteName("Gentech");
        i1.showTeacherName("Prabhakar");
        i1.showInstituteEmail("mail.com");
        i1.showStudentFees(1890.89);
        i1.showInstituteContact(982649680L);

        Institute4 i2=new Institute4();
        i2.showNumOfStuInInstitute(89);
        i2.showInstituteName("Gentech");
        i2.showInstituteAddress("Attigupe");
        i2.showStudentFees(1728.90);
        i2.showTeacherName("Prabhakar");
        i2.showInstituteEmail("mail.com");
    }
}




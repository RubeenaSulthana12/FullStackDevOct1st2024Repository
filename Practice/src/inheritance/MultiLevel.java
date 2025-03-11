package inheritance;
class Employee
{

    void Emp(String name, int age, String jobtype)
    {
        System.out.println("Employee Name:"+name);
        System.out.println("Employee Age:"+age);
        System.out.println("Employee JobType:"+jobtype);
    }
}
class Student extends Employee
{
  void Std(String sname,int sage)
  {
      System.out.println("Student name:"+sname);
      System.out.println("Student age:"+sage);
  }
}
class BusinessMan extends Student
{
    void Man(String btype)
    {

        System.out.println("Business Type:"+btype);
    }
}
public class MultiLevel {
    public static void main(String[] args) {
        BusinessMan o=new BusinessMan();
        o.Emp("Rubeena",21,"BankManager");
        o.Std("Sameer",23);
        o.Man("Merchant");
    }
}

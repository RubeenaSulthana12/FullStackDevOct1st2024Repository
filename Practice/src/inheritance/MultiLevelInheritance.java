package inheritance;
class Employee4
{
    Employee4(String name,int age)
    {
        System.out.println("Employee Name:"+name);
        System.out.println("Employee Age:"+age);
    }
}
class Student4 extends Employee4
{
    Student4(String name,int age,String sname,int sage)
    {
        super(name,age);
        name=sname;
        age=sage;
        System.out.println("Student name:"+sname);
        System.out.println("Student age:"+sage);
    }

}

class BusinessMan4 extends Student4
{
    String businesstype;
    BusinessMan4(String name,int age,String sname,int sage,String btype)
    {
        super(name,age,sname,sage);
        this.businesstype=btype;
        System.out.println("Business Type:"+btype);

    }


}
public class MultiLevelInheritance {
    public static void main(String[] args) {
        BusinessMan4 o= new BusinessMan4("Rubeena",17,"Sulthana",21,"Manager");
        Student4 o1=new Student4("Sulthana",23,"Roshan",25);

    }
}


package inheritance;
class Student2
{
    void st(String name,String gender)
    {
        System.out.println("Student Name:"+name);
        System.out.println("Student Gender:"+gender);
    }
}
class Teacher2 extends Student2
{
    void teac(String name,int age)
    {
        System.out.println("Teacher Name:"+name);
        System.out.println("Teacher age:"+age);
    }
}
class Employee2 extends Student2
{
    void empl(String name,int age)
    {
        System.out.println("Employee Name:"+name);
        System.out.println("Employee Age:"+age);
    }
}
public class Hierarchoical {
    public static void main(String[] args) {
        Employee2 o = new Employee2();
        o.st("Jeelan", "Male");
        o.empl("Roshan", 11);

        Teacher2 o1 = new Teacher2();
        o1.teac("Rubeena", 21);
        o.st("Sameer", "Male");
    }
}

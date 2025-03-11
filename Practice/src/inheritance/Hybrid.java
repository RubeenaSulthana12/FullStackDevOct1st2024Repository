package inheritance;
class School
{
    void s(String sname,int nooffaculties)
    {
        System.out.println("Schol Name:"+sname);
        System.out.println("Number Of Faculties:"+nooffaculties);
    }
}
class Teacher extends School
{
    void t(String tname,int noofsubjectsteaches)
    {
        System.out.println("Teacher Name:"+tname);
        System.out.println("Num Of Sub:"+noofsubjectsteaches);
    }
}
class Classroom extends School
{
    void room(int noofrooms)
    {
        System.out.println("Num Of Rooms:"+noofrooms);
    }
}
class Student1 extends Classroom
{
    void std(String sname,int sage)
    {
        System.out.println("Student name:"+sname);
        System.out.println("Student age:"+sage);
    }
}
public class Hybrid {
    public static void main(String[] args) {
        Student1 o = new Student1();
        o.std("Roshan", 11);
        o.room(12);
        o.s("Rotary High school", 87);
        Teacher o1 = new Teacher();
        o1.t("Lucky", 5);
        o1.s("Narayana", 87);
    }
}

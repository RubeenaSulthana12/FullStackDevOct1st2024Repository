package com.gentech.methods;
class Outer1 {

     String fullname;
     int age;
     String gender;

    class Inner1 {
        void show(String fullname, int age, String gender)
        {
//            fullname = "Rubeena Sulthana";
//            age = 21;
//            gender = "female";
            Outer1.this.fullname = fullname;
            Outer1.this.age = age;
            Outer1.this.gender = gender;
            System.out.println("Full Name:" + Outer1.this.fullname);
            System.out.println("Age:" + Outer1.this.age);
            System.out.println("Gender:" + Outer1.this.gender);
        }
    }

}

public class Assignment2 {
    public static void main(String[] args) {
        Outer1 o=new Outer1();
        Outer1.Inner1 i1= o.new Inner1();
        i1.show("ABC", 21, "F");

        System.out.println(o.gender);

    }
}


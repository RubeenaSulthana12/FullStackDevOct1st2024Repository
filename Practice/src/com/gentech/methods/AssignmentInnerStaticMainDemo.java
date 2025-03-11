package com.gentech.methods;
class Outer11 {

    String fullname;
    int age;
    String gender;

   static class Inner11 {
        void show(String fullname, int age, String gender)
        {
//            fullname = "Rubeena Sulthana";
//            age = 21;
//            gender = "female";
            Outer11 outer11 = new Outer11();
            outer11.fullname = fullname;
            outer11.age = age;
            outer11.gender = gender;
            System.out.println("Full Name:" + outer11.fullname);
            System.out.println("Age:" + outer11.age);
            System.out.println("Gender:" + outer11.gender);
        }
    }

}

public class AssignmentInnerStaticMainDemo {
    public static void main(String[] args) {
        Outer11 o=new Outer11();
        Outer11.Inner11 i1= new Outer11.Inner11();
        i1.show("ABC", 21, "F");
    }
}

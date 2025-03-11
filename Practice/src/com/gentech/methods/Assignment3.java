package com.gentech.methods;
class Outer2 {
    Inner2 o = new Inner2();

    void display() {
        o.show();
        System.out.println("Name:" + o.name);
        System.out.println("Age:" + o.age);
    }
   private class Inner2 {
       private String name;
       private int age;
       private void show()
       {
           name="Rubeena";
           age=21;
       }
   }

   }

public class Assignment3 {
    public static void main(String[] args) {
        Outer2 o1=new Outer2();
        o1.display();
    }
}

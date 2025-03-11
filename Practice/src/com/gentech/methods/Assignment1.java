package com.gentech.methods;
class Outer {
    Inner o = new Inner();
    void show() {
        o.name = "ruby";
        o.age = 21;
        System.out.println("First Name:" + o.name);
        System.out.println("Age:" + o.age);
    }

    private class Inner
    {
        private String name;
        private int age;


    }
}
public class Assignment1 {
    public static void main(String[] args) {
        Outer o1=new Outer();
        o1.show();
    }

}

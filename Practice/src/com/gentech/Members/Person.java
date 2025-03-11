// Instance members can not access directly into a static method
package com.gentech.Members;
public class Person{
    String fullname;
    int age;

    public static void main(String[] args) {
        Person p1 = new Person();
        p1.fullname = "Rubeena";
        p1.age = 21;
        System.out.println("Full name:" + p1.fullname);
        System.out.println("Age:" + p1.age);
    }
}

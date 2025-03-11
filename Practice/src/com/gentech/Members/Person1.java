//static members of the class access directly in static method or block of same class
package com.gentech.Members;

public class Person1 {
    static String firstname;
    static int age;

    public static void main(String[] args) {
        firstname="Sulthana";
        age=21;
        System.out.println("First NAme:"+firstname);
        System.out.println("Age:"+age);
    }
}

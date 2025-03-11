//If a class contains instance block,static block,constructor
package com.gentech.Members;
class Demo6
{
    static
    {
        System.out.println("it is static");
    }
    {
        System.out.println("it is instance");
    }
    Demo6()
    {
        System.out.println("it is constructor");
    }
}

public class InstanceBlock3 {
    public static void main(String[] args) {
        Demo6 o=new Demo6();
    }
}

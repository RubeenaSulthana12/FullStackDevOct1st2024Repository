package com.gentech.Members;
class Demo3
{
    static
    {
        System.out.println("It is a third block");
    }
    static
    {
        System.out.println("It is a First block");
    }
    static
    {
        System.out.println("it is a second bock");
    }
}
public class StaticBlock4 {
    public static void main(String[] args) {
        Demo3 o=new Demo3();
    }
}

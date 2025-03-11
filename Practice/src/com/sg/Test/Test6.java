package com.sg.Test;
class Demo10
{
    {
        System.out.println("Hi I am Ruby");
    }
}
class Demo11
{
    {
        Demo10 o=new Demo10();
        System.out.println("Its Instance Block");
    }
}
public class Test6 {
    public static void main(String[] args) {
        Demo11 o1=new Demo11();
    }
}

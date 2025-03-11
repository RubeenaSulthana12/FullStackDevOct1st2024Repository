package com.sg.Test;
class Demo15
{
    void addition(int x,int y)
    {
        System.out.println("Addition of two numbers:"+(x+y));
    }
   static  void substraction(int x,int y)
    {
        System.out.println("Substraction of two numbers:"+(x-y));
    }
}
class Demo16
{
    {
        Demo15 o=new Demo15();
        o.addition(10,5);
        Demo15.substraction(15,5);
        System.out.println("it is instance");
    }
}
public class Test9 {
    public static void main(String[] args) {
        Demo16 o1=new Demo16();
    }
}

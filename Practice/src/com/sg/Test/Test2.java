package com.sg.Test;
class Demo1
{
    static void substraction(int x,int y)
    {
        System.out.println("Subtraction of two numbrs:"+(x-y));
    }
    static void multiplication(int a,int b)
    {
        System.out.println("Multiplication of two numbers:"+(a*b));
    }
}
class Demo2
{
    Demo2()
    {
        Demo1.multiplication(3,6);
        Demo1.substraction(6,3);
        System.out.println("Hi I am Rubeena");
    }
}
public class Test2 {
    public static void main(String[] args) {
        Demo2 o=new Demo2();

    }
}

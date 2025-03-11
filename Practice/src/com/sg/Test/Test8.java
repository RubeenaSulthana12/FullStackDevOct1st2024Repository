package com.sg.Test;
class Demo12
{
    void addition(int x,int y)
    {
        System.out.println("Addition of two numbers:" +(x+y));
    }
    {
        System.out.println("It is a Instance block");
    }
}
class Demo13
{
    static
    {
        Demo12 o=new Demo12();
        o.addition(5,3);


        System.out.println("it is a static block");
    }
}
public class Test8 {
    public static void main(String[] args) {
        Demo13 o1=new Demo13();

    }
}

package com.sg.Test;
class Demo
{
    void addition(int x,int y)
    {
        System.out.println("Addition of two Numbers:"+(x+y));
    }
    Demo()
    {
        System.out.println("It is a constructor statement");
    }
}
public class Test1 {
    public static void main(String[] args){
        Demo o=new Demo();
        o.addition(10,15);

    }

}

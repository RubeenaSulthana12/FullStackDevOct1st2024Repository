// if static members are available in an independent class
package com.gentech.Members;
class Maths
{
    static String mathsType;
    static void addition(int x,int y)
    {
        int res=(x+y);
        System.out.println("Addition:"+res);
    }
}
public class StaticDemo {
    public static void main(String[] args) {
        Maths.mathsType="Trignometry";
        System.out.println(Maths.mathsType);
        Maths.addition(100,50);
    }
}

// static methods can return a value
package com.gentech.Members;
class Maths1
{
    static int multiplication(int x,int y)
    {
        return (x * y);
    }
}
public class StaticDemo1 {
    public static void main(String[] args) {
        int val1=Maths1.multiplication(12,5);
        System.out.println(val1);
    }
}

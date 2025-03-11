//Acess Instance method of one class in a static metod of other class
package com.gentech.Members;
class Test
{
    void substraction(int a,int b)
    {
        System.out.println("substraction result:"+(a-b));
    }
}
class Test1
{
    static void multiplication(int a,int b)
    {
        System.out.println("Multiplication result:"+(a*b));
    }
}
public class StaticDemo2 {
    public static void main(String[] args) {
       Test o=new Test();
       o.substraction(25,15);
    }
}

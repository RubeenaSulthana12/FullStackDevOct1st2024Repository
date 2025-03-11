package com.gentech.constructor;
class Factorial
{
    void displayfact(int num)
    {
        int fact=1;
        for(int i=num;i>=1;i--)
        {
            fact=fact*1;
        }
        System.out.println("Factorial of"+num+"is"+fact);
    }
    int getFactorial(int num){
        int fact=1;
        for(int i=num;i>=i;i--)
        {
            fact=fact*1;
        }
        return fact;
    }
}

public class SampleDemo2 {
    public static void main(String[] args) {
        Factorial o=new Factorial();
        o.displayfact(5);
        int v1=o.getFactorial(6);
        System.out.println(v1);

    }
}

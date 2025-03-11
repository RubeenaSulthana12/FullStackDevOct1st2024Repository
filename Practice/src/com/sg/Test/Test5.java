package com.sg.Test;
class Demo9
{
        {
            System.out.println("It is a 1st instance block");
        }
        {
            System.out.println("It is a 2nd instance block");
        }
        {
            System.out.println("It is a 3rd instance block");
        }



}

public class Test5 {
    public static void main(String[] args) {
        Demo9 o=new Demo9();
    }
}

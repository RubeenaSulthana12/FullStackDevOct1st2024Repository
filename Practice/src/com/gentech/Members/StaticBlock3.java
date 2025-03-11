//if a class contain static block and constructor once after object execution
package com.gentech.Members;
class Demo2
{
    static
    {
        System.out.println("it is a static block");
    }
    Demo2()
    {
        System.out.println("it is no ars constructor");
    }
}

public class StaticBlock3 {
    public static void main(String[] args) {
        Demo2 o=new Demo2();
    }

}

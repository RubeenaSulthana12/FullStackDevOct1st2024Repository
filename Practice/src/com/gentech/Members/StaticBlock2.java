//if a class contain static block alone
package com.gentech.Members;
class Demo1
{
    static
    {
        System.out.println("it is a static block");
    }
}

public class StaticBlock2 {
    public static void main(String[] args) {
        Demo1 o=new Demo1();
    }
}

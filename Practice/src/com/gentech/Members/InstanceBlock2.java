//Multiple instance blocks in same class
package com.gentech.Members;
class Demo5
{
    {
        System.out.println("It is an instance");
    }
    {
        System.out.println("It is an instance");
    }
}
public class InstanceBlock2 {
    public static void main(String[] args) {
        Demo5 o=new Demo5();
    }
}




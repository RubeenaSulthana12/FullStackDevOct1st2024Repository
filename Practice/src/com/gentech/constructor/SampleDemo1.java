package com.gentech.constructor;
class Name
{
    void display(String name)
    {
        System.out.println("my name is:"+name);
    }
    String getName(String name)
    {
        return name;
    }
}

public class SampleDemo1 {
    public static void main(String[] args) {
        Name o=new Name();
        o.display("Ruby");
        String v1=o.getName("Sulthana");
        System.out.println(v1);

    }

}

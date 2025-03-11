package com.gentech.abstractdemo;

public class Assign2 {
    public static void main(String[] args)
    {
        String s = "ABCDEF";
        String rev = "";
        for (int i = s.length() - 1; i >= 0; i--)
        {
            rev += s.charAt(i);
        }
        System.out.println(rev);
    }
}

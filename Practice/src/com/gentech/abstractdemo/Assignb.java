package com.gentech.abstractdemo;

public class Assignb {
    public static void main(String[] args)
    {
        String s = "ABCDEF";
        char[] charArray = s.toCharArray();
        char[] rev = new char[charArray.length];
        int j = 0;
        for (int i = charArray.length - 1; i >= 0; i--)
        {
            rev[j++] = charArray[i];
        }
        System.out.println(new String(rev));
    }


}

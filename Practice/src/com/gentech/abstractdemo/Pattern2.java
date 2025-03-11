package com.gentech.abstractdemo;

public class Pattern2 {
    public static void main(String[] args)
    {
        String s = "PROGRAM";
        int a = s.length();
        for (int i = a - 1; i >= 0; i--)
        {
            for (int j = 0; j <= i; j++)
            {
                System.out.print(s.charAt(j));
            }
            System.out.println();
        }
    }

}

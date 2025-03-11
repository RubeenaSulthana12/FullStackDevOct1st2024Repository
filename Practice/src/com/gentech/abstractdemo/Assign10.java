package com.gentech.abstractdemo;

public class Assign10 {

    public static void main(String[] args)
    {
        String str = "Rubeena";
        int count = 0;
        for (char c : str.toCharArray())
        {
            count++;
        }
        System.out.println("Number of characters: " + count);
    }
}

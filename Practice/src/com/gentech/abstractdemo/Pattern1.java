package com.gentech.abstractdemo;
public class Pattern1
    {
        public static void main(String[] args)
        {
            String s="PROGRAM";
            int a=s.length();
            for(int i=0;i<a;i++)
            {
                for(int j=0;j<=i;j++)
                {
                    System.out.print(s.charAt(j));
                }
                System.out.println();
            }
        }

    }


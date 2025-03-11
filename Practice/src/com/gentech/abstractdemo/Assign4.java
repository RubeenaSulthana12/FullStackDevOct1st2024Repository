package com.gentech.abstractdemo;

public class Assign4 {
    public static void main(String[] args)
    {
        String s="SUNDAYMONDAYTUESDAYWEDNESDAYTHURSDAYFRIDAYSATURDAY";
        String result = "";
        for (int i = 0; i < s.length(); i++)
        {
            if (i <= s.length() - 3 && s.substring(i, i + 3).equals("DAY"))
            {
                i += 2;
            } else {
                result += s.charAt(i);
            }
        }
        System.out.println("Result: " + result);

    }

}

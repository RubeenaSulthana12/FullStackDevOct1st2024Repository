package com.gentech.abstractdemo;

public class Assign5 {
    public static void main(String[] args) {
        String days = "SUNDAYMONDAYTUESDAYWEDNESDAYTHURSDAYFRIDAYSATURDAY";
        String result = insertColon(days);
        System.out.println(result);
    }

    public static String insertColon(String days) {
        StringBuilder result = new StringBuilder();
        String[] dayNames = {"SUNDAY", "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY"};

        int i = 0;
        while (i < days.length()) {
            boolean found = false;
            for (String day : dayNames) {
                if (days.startsWith(day, i)) {
                    result.append(day).append(";");
                    i += day.length();
                    found = true;
                    break;
                }
            }
            if (!found) {
                i++;
            }
        }

        return result.toString();
    }
}

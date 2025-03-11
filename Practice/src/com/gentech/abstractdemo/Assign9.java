package com.gentech.abstractdemo;

public class Assign9 {
    public static void main(String[] args) {
        String sentence = "banglore is a garden city";
        String[] words = sentence.split(" ");
        String result = "";

        for (String word : words) {
            char[] chars = word.toCharArray();
            for (int i = 0, j = chars.length - 1; i < j; i++, j--) {

                char temp = chars[i];
                chars[i] = chars[j];
                chars[j] = temp;
            }
            result += new String(chars) + " ";
        }

        System.out.println("Original Sentence: " + sentence);
        System.out.println("Reversed Words Sentence: " + result.trim());




    }
}

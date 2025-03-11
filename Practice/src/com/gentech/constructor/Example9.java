package com.gentech.constructor;
class FibDemo
{
    void fib()
    {
        int fn=0;
        int sn=1;
        System.out.print(fn+" "+sn+" ");
        for(int i=2;i<=9;i++)
        {
            int tn=fn+sn;
            fn=sn;
            sn=tn;
            System.out.print(tn+" ");
        }
    }
}
public class Example9 {
    public static void main(String[] args) {
        FibDemo f1=new FibDemo();
        f1.fib();
    }

}



package com.gentech.abstractdemo;
abstract class Order
{
    int Id;
    Order(int Id)
    {
        System.out.println("Order Id :"+Id);
    }
    void displayOrderId(int Id)
    {
        System.out.println("Order Id :"+Id);
    }
}
class Shipping extends Order
{
    Shipping(int Id)
    {
        super(Id);
    }
    void showOrderId(int Id)
    {
        System.out.println("Order Id :"+Id);
    }
    void showShippingStatus(String status)
    {
        System.out.println("Shipping Status :"+status);
    }
}



public class InstanceMethod {
    public static void main(String[] args)
    {
        Shipping s1=new Shipping(98);
        s1.showOrderId(87);
        s1.showShippingStatus("Fast");
    }

}



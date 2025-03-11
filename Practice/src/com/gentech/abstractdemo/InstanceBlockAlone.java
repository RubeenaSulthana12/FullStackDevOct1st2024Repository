package com.gentech.abstractdemo;
abstract class Airport
{
    String aName;
    abstract void showAirportName();
    abstract void showAirportId();
    void showAirportCity(String city[])
    {
        for(int i=0;i<city.length;i++)
        {
            System.out.println("Airport City :"+city[i]);
        }
    }
}
class Pilot extends Airport
{
    int pId;
    void showAirportName()
    {
        System.out.println("Airport Name :"+aName);
    }
    void showAirportId()
    {
        System.out.println("Airport Id:"+pId);
    }
    void showPilotName(String pName)
    {
        System.out.println("Pilot Name :"+pName);
    }
}


public class InstanceBlockAlone {
    public static void main(String[] args)
    {
        Pilot p1=new Pilot();
        p1.pId=1001;
        p1.aName="Rubeena";

        p1.showAirportId();
        p1.showPilotName("Kiragi");
        p1.showAirportName();
        p1.showAirportCity(new String[]{"Banglore","Mysore"});
    }

}



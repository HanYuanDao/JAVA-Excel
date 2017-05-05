package main;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;

/**
 * Desciption:
 * Author: JasonHan.
 * Creation time: 2017/04/14 15:59:00.
 * © Copyright 2013-2017, Node Supply Chain Management.
 */
public class Calculate {

    public static void main(String[] myArgs) {
        float vol;
        float qty;
        String qtyAndVol;
        String[] strList;
        InputStreamReader is_reader = new InputStreamReader(System.in);
        while (true) {
            try {
                System.out.println("请输入重量（公斤）和体积（立方米），格式为：22 24");
                qtyAndVol = new BufferedReader(is_reader).readLine();
                strList = qtyAndVol.split(" ");
                if (strList.length < 2) {
                    System.out.println("格式错误");
                } else {
                    qty = Float.valueOf(strList[0]);
                    vol = Float.valueOf(strList[1]);
                    System.out.println("qty:" + qty + " vol:" + vol);
                    System.out.println("在方案一中为" + ((qty>=(vol*1000000/6000))?"重货":"泡货"));
                    System.out.println("在方案二中为" + ((qty>=(vol*210))?"重货":"泡货"));
                }

            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }
}

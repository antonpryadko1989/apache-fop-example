package com.visoft.example;

import org.apache.fop.fonts.apps.TTFReader;

/**
 * @author vlad
 * hello Anton!!!
 */
public class MetricsGen {
	public static void main(String[] args) {
//		for (int i =  0; i < args.length; i++) {
//			System.out.println("args: " + args[i]);
//		}
		
		TTFReader.main(new String[] {"C:\\Users\\vlad\\Desktop\\vi-soft-NG\\FOP-example\\apache-fop-example\\src\\main\\resources\\arial.ttf", "C:\\Users\\vlad\\Desktop\\vi-soft-NG\\FOP-example\\apache-fop-example\\src\\main\\resources\\arial.xml"});
	}
}

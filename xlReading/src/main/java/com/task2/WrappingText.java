package com.task2;

import org.davidmoten.text.utils.WordWrap;

public class WrappingText {

	public static void main(String[] args) {
		String text = "hello how are you going?";
		String wrapped = 
		  WordWrap.from(text)
		    .maxWidth(11)
		    .insertHyphens(true) // true is the default
		    .wrap();
		System.out.println(wrapped);
	}

}

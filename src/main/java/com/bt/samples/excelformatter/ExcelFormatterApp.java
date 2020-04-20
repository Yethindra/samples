package com.bt.samples.excelformatter;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class ExcelFormatterApp {

	public static void main(String[] args) {
		System.out.println("Enter the input file path :");
		Scanner in = new Scanner(System.in);
		String inputFilePath = in.nextLine();
		System.out.println("Enter the output file path :");
		String outputFilePath = in.nextLine();
		System.out.println("Input File :" + inputFilePath);
		System.out.println("Output File :" + outputFilePath);
		try {
			if (inputFilePath == null || outputFilePath == null) {
	             throw new FileNotFoundException("Input/Output file path is missing");
			}
			ExcelFormatter formatter = new ExcelFormatter(inputFilePath, outputFilePath);
			formatter.format();
			System.out.println("File is formatted");
		} catch (FileNotFoundException e) {
			System.out.println("File not found" + e.getCause());
		} catch (IOException e) {
			System.out.println("File IO error" + e.getCause());
		} catch (EncryptedDocumentException e) {
			System.out.println("File format error" + e.getCause());
		} catch (InvalidFormatException e) {
			System.out.println("File format error" + e.getCause());
		} finally {
			in.close();
		}
	}

}

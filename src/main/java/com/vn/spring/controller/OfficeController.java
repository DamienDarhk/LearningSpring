package com.vn.spring.controller;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class OfficeController {

	public static void main(String[] args) throws IOException {
		updateDocument("D:\\Work_project\\template_contract.docx", "D:\\Work_project\\contract_out.docx");
	}

	private static void updateDocument(String input, String output) throws IOException {

		try (XWPFDocument doc = new XWPFDocument(Files.newInputStream(Paths.get(input)))) {

			List<XWPFParagraph> xwpfParagraphList = doc.getParagraphs();
			// Iterate over paragraph list and check for the replaceable text in each
			// paragraph
			for (XWPFParagraph xwpfParagraph : xwpfParagraphList) {
				for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {
					String docText = xwpfRun.getText(0);
					System.out.println(docText);
					if (docText != null) {
						docText = replaceTextContract(docText);
					}
					xwpfRun.setText(docText, 0);
				}
			}
			// save the docs
			try (FileOutputStream out = new FileOutputStream(output)) {
				doc.write(out);
			}
			System.out.println("done");
		}
	}

	public static String replaceTextContract(String text) {
		if (text.contains("${cifNo1}")) {
			text = text.replace("${cifNo}", "1234567");
			return text;
		} else if (text.contains("${cifNo2}")) {
			text = text.replace("${cifNo2}", "1234567");
			text = text.replace("${sysdate}", "20241219");
			return text;
		} else if (text.contains("${sysdateDay}")) {
			text = text.replace("${sysdateDay}", "19");
			text = text.replace("${sysdateMonth}", "12");
			text = text.replace("${sysdateYear}", "2024");
			return text;
		} else if (text.contains("${limitAmount}")) {
			text = text.replace("${limitAmount}", "50.000.000");
			return text;
		} else if (text.contains("${t24.cusName}")) {
			text = text.replace("${t24.cusName}", "Bùi Thành Công");
			return text;
		} else if (text.contains("${t24.gender}")) {
			text = text.replace("${t24.gender}", "Nam");
			return text;
		} else if (text.contains("${t24.birthDay}")) {
			text = text.replace("${t24.birthDay}", "16/08/1998");
			return text;
		} else if (text.contains("${t24.nationality}")) {
			text = text.replace("${t24.nationality}", "Việt Nam");
			return text;
		} else if (text.contains("${t24.legalId}")) {
			text = text.replace("${t24.legalId}", "123456789");
			return text;
		} else if (text.contains("${t24.legalIssDate}")) {
			text = text.replace("${t24.legalIssDate}", "19/12/2020");
			return text;
		} else if (text.contains("${t24.legalIssAuth}")) {
			text = text.replace("${t24.legalIssAuth}", "Công An Bình Dương");
			return text;
		} else if (text.contains("${marriedStatus}")) {
			text = text.replace("${marriedStatus}", "OK");
			return text;
		} else if (text.contains("${t24.street}")) {
			text = text.replace("${t24.street}", "Hà Nội");
			return text;
		} else if (text.contains("${objUser.mobile}")) {
			text = text.replace("${objUser.mobile}", "0397133365");
			return text;
		} else if (text.contains("${email}")) {
			text = text.replace("${email}", "congbt@ncb-bank.vn");
			return text;
		} else if (text.contains("${company_type}")) {
			text = text.replace("${company_type}", "Finance");
			return text;
		} else if (text.contains("${company}")) {
			text = text.replace("${company}", "NCB");
			return text;
		} else if (text.contains("${company_address}")) {
			text = text.replace("${company_address}", "37 Ngô Quyền - Phan Chu Trinh - Hoàn Kiếm");
			return text;
		} else if (text.contains("${company_phoneNumber}")) {
			text = text.replace("${company_phoneNumber}", "0123456789");
			return text;
		} else if (text.contains("${position}")) {
			text = text.replace("${position}", "Chuyên viên");
			return text;
		} else if (text.contains("${working_month}")) {
			text = text.replace("${working_month}", "12");
			return text;
		} else if (text.contains("${total_amount}")) {
			text = text.replace("${total_amount}", "10.000.000");
			return text;
		} else if (text.contains("${month_payment_amount}")) {
			text = text.replace("${month_payment_amount}", "5.000.000");
			return text;
		} else if (text.contains("${contacts_name}")) {
			text = text.replace("${contacts_name}", "Mít");
			return text;
		} else if (text.contains("${contacts_relationship}")) {
			text = text.replace("${contacts_relationship}", "Con gái");
			return text;
		} else if (text.contains("${contacts_phoneNumber}")) {
			text = text.replace("${contacts_phoneNumber}", "0101010101");
			return text;
		} else if (text.contains("${Answer}")) {
			text = text.replace("${Answer}", "SH");
			return text;
		} else {
			return text;
		}
	}

}
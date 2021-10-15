package de.tgx03.csvconverter;

import ezvcard.VCard;
import ezvcard.property.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jetbrains.annotations.NotNull;

import java.io.File;
import java.io.IOException;

public class Converter {

	private static final int ID = 0;
	private static final int FIRSTNAME = 1;
	private static final int LASTNAME = 2;
	private static final int TITLE = 3;
	private static final int BIRTHDAY = 4;
	private static final int GENDER = 5;
	private static final int ADDRESS_EXTRA = 6;
	private static final int STREET_ADDRESS = 7;
	private static final int POSTCODE = 8;
	private static final int TOWN = 9;
	private static final int COUNTRY = 10;
	private static final int PHONE1 = 11;
	private static final int PHONE2 = 12;
	private static final int MAIL = 13;
	private static final int FAMILY_ID = 14;
	private static final int CONTACT = 15;
	private static final int PUBLICATION = 16;
	private static final int EXTRA1 = 17;
	private static final int EXTRA2 = 18;
	private static final int MEMO = 19;
	private static final int ACTIVE = 20;
	private static final int ENTRY = 21;
	private static final int EXIT = 22;
	private static final int MAIN = 23;
	private static final int DEPARTMENTS = 24;
	private static final int FEES = 25;
	private static final int HONORS = 26;
	private static final int FUNCTIONS = 27;

	private static File input;
	private static String outputDirectory;

	public static void main(@NotNull String[] args) throws IOException, InvalidFormatException {
		input = new File(args[0]);
		outputDirectory = args[1];

		Workbook workbook = WorkbookFactory.create(input);
		Sheet sheet = workbook.getSheetAt(0);
		int max = sheet.getPhysicalNumberOfRows();
		for (int i = 1; i < max; i++) {
			Row row = sheet.getRow(i);
			VCard card = new VCard();
			card.setFormattedName(row.getCell(FIRSTNAME).getStringCellValue() + " " + row.getCell(LASTNAME).getStringCellValue());
			card.setBirthday(new Birthday(row.getCell(BIRTHDAY).getDateCellValue()));
			String mail = row.getCell(MAIL).getStringCellValue();
			if (!mail.isBlank()) {
				card.addEmail(new Email(mail));
			}
			String number1 = row.getCell(PHONE1).getStringCellValue();
			if (!number1.isBlank()) {
				card.addTelephoneNumber(new Telephone(formatPhoneNumber(number1)));
			}
			String number2 = row.getCell(PHONE2).getStringCellValue();
			if (!number2.isBlank()) {
				card.addTelephoneNumber(new Telephone(formatPhoneNumber(number2)));
			}
			card.setGender(parseGender(row.getCell(GENDER).getStringCellValue()));

			Address address = new Address();
			address.setStreetAddress(row.getCell(STREET_ADDRESS).getStringCellValue());
			address.setCountry(row.getCell(COUNTRY).getStringCellValue());
			address.setLocality(row.getCell(TOWN).getStringCellValue());
			address.setPostalCode(row.getCell(POSTCODE).getStringCellValue());

			card.write(new File(outputDirectory + card.hashCode() + ".vcf"));
		}
	}

	@NotNull
	private static String formatPhoneNumber(@NotNull String original) {
		if (original.charAt(0) != '0') return "+497248" + original;
		else if (original.charAt(1) == '0') return original;
		else return "+49" + original;
	}

	private static Gender parseGender(@NotNull String value) {
		switch (value) {
			case "Männlich" -> {return Gender.male();}
			case "Weiblich" -> {return Gender.female();}
			case "Divers" -> {return Gender.other();}
			default -> {return Gender.unknown();}
		}
	}
}

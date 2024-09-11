///usr/bin/env jbang -C -Xlint:-options "$0" "$@" ; exit $?
//DEPS org.apache.poi:poi-ooxml:5.3.0
//DEPS org.apache.logging.log4j:log4j-core:2.23.1

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.security.GeneralSecurityException;
import java.util.Iterator;
import java.util.List;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

public class KonverterKlasseliste {
    private static final char CSV_SEPARATOR = ',';

    private String filnavn;
    private String passord;

    public KonverterKlasseliste(String filnavn, String passord) {
        this.filnavn = filnavn;
        this.passord = passord;
    }

    private void konverter() throws IOException, GeneralSecurityException {
        List<Oppføring> elever = lesElevOppføringerFraFil();
        skrivForesatteSomCsvForKontaktImport(elever);
    }

    private List<Oppføring> lesElevOppføringerFraFil() throws IOException, GeneralSecurityException {
        try (Workbook workbook = openWorkbook()) {
            return lesElevOppføringer(workbook);
        }
    }

    private Workbook openWorkbook() throws IOException, GeneralSecurityException {
        var filendelse = filnavn.substring(filnavn.indexOf('.'));
        if (filendelse.endsWith("xlsx")) {
            return readNewFormatEncrypted();
        } else if (filendelse.endsWith("xls")) {
            return readOldFormat();
        } else {
            System.err.println(String.format("Ikke støttet filformat: %s", filendelse));
            System.exit(1);
            throw new IllegalStateException(); // will never get here
        }
    }

    private Workbook readNewFormatEncrypted() throws IOException, GeneralSecurityException {
        try (InputStream fis = new FileInputStream(filnavn);
             POIFSFileSystem poifs = new POIFSFileSystem(fis)) {
            Decryptor decryptor = getDecryptor(poifs);
            return new XSSFWorkbook(decryptor.getDataStream(poifs));
        }
    }

    private Decryptor getDecryptor(POIFSFileSystem poifs) throws IOException, GeneralSecurityException {
        EncryptionInfo encryptionInfo = new EncryptionInfo(poifs);
        Decryptor decryptor = Decryptor.getInstance(encryptionInfo);

        if (!decryptor.verifyPassword(passord)) {
            System.err.println("Feil passord!");
            System.exit(1);
        }
        return decryptor;
    }

    private Workbook readOldFormat() throws IOException {
        try (InputStream fis = new FileInputStream(filnavn)) {
            return new HSSFWorkbook(fis);
        }
    }

    private List<Oppføring> lesElevOppføringer(Workbook workbook) {
        Sheet exportSheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = exportSheet.rowIterator();
        verifyHeaders(rowIterator.next());

        return iteratorToStream(rowIterator)
                .map(this::lesElevOppføring)
                .toList();
    }

    private void verifyHeaders(Row headerRow) {
        if (!"Klasse".equals(getNullableStringCellValue(headerRow, 0))) {
            throw new IllegalArgumentException("Filen mangler gyldig header");
        }
    }

    private static Stream<Row> iteratorToStream(Iterator<Row> rowIterator) {
        return StreamSupport.stream(
                Spliterators.spliteratorUnknownSize(rowIterator, Spliterator.ORDERED),
                false
        );
    }

    private Oppføring lesElevOppføring(Row elevRad) {
        return new Oppføring(
                getNullableStringCellValue(elevRad, 0),
                getNullableStringCellValue(elevRad, 1),
                getNullableStringCellValue(elevRad, 2),
                getNullableStringCellValue(elevRad, 3),
                getNullableStringCellValue(elevRad, 4),
                getNullableStringCellValue(elevRad, 5),
                getNullableStringCellValue(elevRad, 6),
                getNullableStringCellValue(elevRad, 7),
                getNullableStringCellValue(elevRad, 8)
        );
    }

    private static String getNullableStringCellValue(Row elevRad, int cellnum) {
        Cell cell = elevRad.getCell(cellnum);
        try {
            if (cell == null) {
                return null;
            } else {
                String tekst = fjernVrangSpace(cell.getStringCellValue()).trim();
                return tekst.isEmpty() ? null : tekst;
            }
        } catch (Exception e) {
            System.err.println(String.format("Ignorerer kolonne %d i rad %d: %s", cellnum, elevRad.getRowNum(), e.getMessage()));
            return null;
        }
    }

    private static String fjernVrangSpace(String tekst) {
        int vrangSpaceIdx = tekst.indexOf(160);
        if (vrangSpaceIdx > -1) {
            char vrangChar = tekst.charAt(vrangSpaceIdx);
            return tekst.replace(vrangChar, ' ');
        } else {
            return tekst;
        }
    }

    private void skrivForesatteSomCsvForKontaktImport(List<Oppføring> elever) {
        StringBuffer buffer = new StringBuffer();

        skrivHeader(buffer);
        elever.forEach(oppføring -> skrivElev(oppføring, buffer));

        System.out.print(buffer);
    }

    private void skrivHeader(StringBuffer buffer) {
        buffer.append("First Name").append(CSV_SEPARATOR);
        buffer.append("Last Name").append(CSV_SEPARATOR);
        buffer.append("Telephone").append(CSV_SEPARATOR);
        buffer.append("Email");
        buffer.append('\n');
    }

    private void skrivElev(Oppføring oppføring, StringBuffer buffer) {
        if (oppføring.navnForelder1 != null) {
            skrivOppføring(oppføring, buffer, oppføring.navnForelder1(), oppføring.telefonForelder1(), oppføring.epostForelder1());
        }
        if (oppføring.navnForelder2 != null) {
            skrivOppføring(oppføring, buffer, oppføring.navnForelder2(), oppføring.telefonForelder2(), oppføring.epostForelder2());
        }
    }

    private static void skrivOppføring(Oppføring oppføring, StringBuffer buffer, String forelderNavn, String telefonForelder, String epostForelder) {
        if (telefonForelder == null && epostForelder == null) {
            return;
        }

        buffer.append(muligTomTekst(forelderNavn)).append(CSV_SEPARATOR);
        buffer.append(String.format("(%s %s %s)", oppføring.fornavn, oppføring.etternavn(), oppføring.klasse())).append(CSV_SEPARATOR);

        if (telefonForelder != null) {
            buffer.append(telefonnummer(telefonForelder)).append(CSV_SEPARATOR);
        } else {
            buffer.append(CSV_SEPARATOR);
        }

        if (epostForelder != null) {
            buffer.append(muligTomTekst(epostForelder));
        }

        buffer.append('\n');
    }

    private static String muligTomTekst(String navn) {
        return navn == null ? "" : navn;
    }

    private static String telefonnummer(String telefonnummer) {
        if (telefonnummer == null) {
            return "";
        }

        var trimmetNummer = telefonnummer.replace(" ", "");
        if (trimmetNummer.length() > 8 && !trimmetNummer.startsWith("47")) {
            System.err.println("Støtter ikke utenlandsk telefonnummer: " + telefonnummer);
        }
        if (trimmetNummer.length() == 10 && trimmetNummer.startsWith("47")) {
            return trimmetNummer.substring(2);
        } else {
            return trimmetNummer;
        }
    }

    private record Oppføring(String klasse,
                             String etternavn,
                             String fornavn,
                             String navnForelder1,
                             String telefonForelder1,
                             String epostForelder1,
                             String navnForelder2,
                             String telefonForelder2,
                             String epostForelder2) {
    }

    ;

    public static void main(String[] args) throws Exception {
        if (args.length < 1 || args.length > 2) {
            printUsage();
            System.exit(1);
        }

        String filnavn = args[0];
        String passord = null;
        if (args.length == 2) {
            passord = args[1];
        }

        new KonverterKlasseliste(filnavn, passord).konverter();
    }

    private static void printUsage() {
        var usage = """
                Bruk slik: ./KonverterKlasseliste.java <filnavn> [<passord>]
                
                Sett <filnavn> i hermetegn, hvis det har mellomrom i seg: "Klasseliste 1A.xls"
                """;
        System.err.println(usage);
    }
}

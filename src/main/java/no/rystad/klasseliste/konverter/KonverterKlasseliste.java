package no.rystad.klasseliste.konverter;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.List;

public class KonverterKlasseliste {
    private static final char CSV_SEPARATOR = ',';

    private final ExcelRegneark excelRegneark;

    public KonverterKlasseliste(String filnavn, String passord) {
        this.excelRegneark = new ExcelRegneark(filnavn, passord);
    }

    private void konverter() throws IOException, GeneralSecurityException, ExcelRegneark.FeilPassord, ExcelRegneark.IkkeStøttetFilendelse {
        List<Oppføring> elever = lesElevOppføringerFraFil();
        skrivForesatteSomCsvForKontaktImport(elever);
    }

    private List<Oppføring> lesElevOppføringerFraFil() throws IOException, GeneralSecurityException, ExcelRegneark.FeilPassord, ExcelRegneark.IkkeStøttetFilendelse {
        try (Workbook workbook = excelRegneark.openWorkbook()) {
            return new Klasseliste(workbook).lesElevOppføringer();
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
        if (oppføring.navnForelder1() != null) {
            skrivOppføring(oppføring, buffer, oppføring.navnForelder1(), oppføring.telefonForelder1(), oppføring.epostForelder1());
        }
        if (oppføring.navnForelder2() != null) {
            skrivOppføring(oppføring, buffer, oppføring.navnForelder2(), oppføring.telefonForelder2(), oppføring.epostForelder2());
        }
    }

    private static void skrivOppføring(Oppføring oppføring, StringBuffer buffer, String forelderNavn, String telefonForelder, String epostForelder) {
        if (telefonForelder == null && epostForelder == null) {
            return;
        }

        buffer.append(muligTomTekst(forelderNavn)).append(CSV_SEPARATOR);
        buffer.append(String.format("(%s %s %s)", oppføring.fornavn(), oppføring.etternavn(), oppføring.klasse())).append(CSV_SEPARATOR);

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

        try {
            new KonverterKlasseliste(filnavn, passord).konverter();
        } catch (ExcelRegneark.FeilPassord e) {
            System.err.println("Feil passord!");
            System.exit(1);
        } catch (ExcelRegneark.IkkeStøttetFilendelse e) {
            System.err.println(String.format("Ikke støttet filformat: %s", e.filendelse()));
            System.exit(1);
        }
    }

    private static void printUsage() {
        var usage = """
                Bruk slik: ./KonverterKlasseliste.java <filnavn> [<passord>]
                
                Sett <filnavn> i hermetegn, hvis det har mellomrom i seg: "Klasseliste 1A.xls"
                """;
        System.err.println(usage);
    }
}

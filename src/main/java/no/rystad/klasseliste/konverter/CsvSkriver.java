package no.rystad.klasseliste.konverter;

import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.util.List;

class CsvSkriver {
    private static final char CSV_SEPARATOR = ',';

    private final PrintWriter writer;

    CsvSkriver(OutputStream outputStream) {
        this.writer = new PrintWriter(new OutputStreamWriter(outputStream));
    }

    void skrivForesatteSomCsvForKontaktImport(List<Oppføring> elever) {
        StringBuffer buffer = new StringBuffer();

        skrivHeader(buffer);
        elever.forEach(oppføring -> skrivElev(oppføring, buffer));

        writer.print(buffer);
        writer.flush();
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
}

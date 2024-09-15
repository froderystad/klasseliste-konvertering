package no.rystad.klasseliste.konverter;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.OutputStream;
import java.security.GeneralSecurityException;
import java.util.List;

public class KonverterKlasselisteCommand {
    private final ExcelRegneark excelRegneark;
    private final CsvSkriver csvSkriver;

    KonverterKlasselisteCommand(String filnavn, String passord, OutputStream outputStream) {
        this.excelRegneark = new ExcelRegneark(filnavn, passord);
        this.csvSkriver = new CsvSkriver(outputStream);
    }

    void konverter() throws IOException, GeneralSecurityException, ExcelRegneark.FeilPassord, ExcelRegneark.IkkeStøttetFilendelse {
        List<Oppføring> elever = lesElevOppføringerFraFil();
        csvSkriver.skrivForesatteSomCsvForKontaktImport(elever);
    }

    private List<Oppføring> lesElevOppføringerFraFil() throws IOException, GeneralSecurityException, ExcelRegneark.FeilPassord, ExcelRegneark.IkkeStøttetFilendelse {
        try (Workbook workbook = excelRegneark.openWorkbook()) {
            return new KlasselisteParser(workbook).lesElevOppføringer();
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
            new KonverterKlasselisteCommand(filnavn, passord, System.out).konverter();
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

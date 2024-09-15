package no.rystad.klasseliste.konverter;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.io.FileNotFoundException;
import java.net.URL;

import static org.junit.jupiter.api.Assertions.assertThrows;

class ExcelRegnearkTest {

    private static final String TESTFIL_PASSORD = "test123";

    @Test
    void åpneFil_nårKryptertOgRiktigPassord_lykkes() throws Exception {
        åpneOgLukk(new ExcelRegneark(testfilnavn("/test-klasseliste-encrypted.xlsx"), TESTFIL_PASSORD));
    }

    @Test
    void åpneFil_nårKryptertOgFeilPassord_feiler() {
        assertThrows(ExcelRegneark.FeilPassord.class, () ->
                new ExcelRegneark(testfilnavn("/test-klasseliste-encrypted.xlsx"), "feil").openWorkbook());
    }

    @Test
    void åpneFil_nårKryptertOgUtenPassord_feiler() {
        assertThrows(ExcelRegneark.FeilPassord.class, () ->
                new ExcelRegneark(testfilnavn("/test-klasseliste-encrypted.xlsx"), null).openWorkbook());
    }

    @Test
    void åpneFil_nårUkryptertOgUtenPassord_lykkes() throws Exception {
        åpneOgLukk(new ExcelRegneark(testfilnavn("/test-klasseliste-open.xlsx"), null));
    }

    @Test
    void åpneLegacyFil_nårKryptertOgMedPassord_lykkes() throws Exception {
        åpneOgLukk(new ExcelRegneark(testfilnavn("/test-klasseliste-legacy-encrypted.xls"), TESTFIL_PASSORD));
    }

    @Test
    void åpneLegacyFil_nårUkryptertOgUtenPassord_lykkes() throws Exception {
        åpneOgLukk(new ExcelRegneark(testfilnavn("/test-klasseliste-legacy-open.xls"), null));
    }

    @Test
    void åpneLegacyFil_nårUkryptertOgMedPassord_feiler() throws Exception {
        åpneOgLukk(new ExcelRegneark(testfilnavn("/test-klasseliste-legacy-open.xls"), "feil"));
    }

    @Test
    void openWorkbook_medIkkeStøttetFilendelse_feiler() {
        assertThrows(ExcelRegneark.IkkeStøttetFilendelse.class, () ->
                new ExcelRegneark("/test-klasseliste-feiler.txt", null).openWorkbook());
    }

    private String testfilnavn(String filnavn) throws FileNotFoundException {
        URL resource = this.getClass().getResource(filnavn);
        if (resource == null) {
            throw new FileNotFoundException(filnavn);
        }
        return resource.getFile();
    }

    private static void åpneOgLukk(ExcelRegneark regneark) throws Exception {
        Workbook workbook = regneark.openWorkbook();
        workbook.close();
    }
}
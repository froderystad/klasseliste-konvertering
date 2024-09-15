package no.rystad.klasseliste.konverter;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.io.FileNotFoundException;
import java.net.URL;

import static org.junit.jupiter.api.Assertions.assertThrows;

class KonverterKlasselisteTest {
    @Test
    void åpneFil_nårKryptertOgRiktigPassord_lykkes() throws Exception {
        KonverterKlasseliste konverterer = new KonverterKlasseliste(testfilnavn("/test-klasseliste-encrypted.xlsx"), "test123");
        Workbook workbook = konverterer.openWorkbook();
        workbook.close();
    }

    @Test
    void åpneFil_nårKryptertOgFeilPassord_feiler() {
        assertThrows(KonverterKlasseliste.FeilPassord.class, () ->
                new KonverterKlasseliste(testfilnavn("/test-klasseliste-encrypted.xlsx"), "feil").openWorkbook());
    }

    @Test
    void åpneFil_nårKryptertOgUtenPassord_feiler() {
        assertThrows(KonverterKlasseliste.FeilPassord.class, () ->
                new KonverterKlasseliste(testfilnavn("/test-klasseliste-encrypted.xlsx"), null).openWorkbook());
    }

    @Test
    void åpneLegacyFil_nårUkryptertOgUtenPassord_lykkes() throws Exception {
        KonverterKlasseliste konverterer = new KonverterKlasseliste(testfilnavn("/test-klasseliste-legacy-open.xls"), null);
        Workbook workbook = konverterer.openWorkbook();
        workbook.close();
    }

    @Test
    void åpneLegacyFil_nårUkryptertOgMedPassord_feiler() throws Exception {
        KonverterKlasseliste konverterer = new KonverterKlasseliste(testfilnavn("/test-klasseliste-legacy-open.xls"), "feil");
        Workbook workbook = konverterer.openWorkbook();
        workbook.close();
    }

    @Test
    void openWorkbook_medIkkeStøttetFilendelse_feiler() {
        assertThrows(KonverterKlasseliste.IkkeStøttetFilendelse.class, () ->
                new KonverterKlasseliste("/test-klasseliste-feiler.txt", null).openWorkbook());
    }

    private String testfilnavn(String filnavn) throws FileNotFoundException {
        URL resource = this.getClass().getResource(filnavn);
        if (resource == null) {
            throw new FileNotFoundException(filnavn);
        }
        return resource.getFile();
    }
}
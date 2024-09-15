package no.rystad.klasseliste.konverter;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.io.FileNotFoundException;
import java.net.URL;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

class KlasselisteParserTest {
    @ParameterizedTest
    @CsvSource({
            "/test-klasseliste-encrypted.xlsx,test123",
            "/test-klasseliste-legacy-open.xls,null"
    })
    void lesElevoppføringer_leserRiktig(String filnavn, String passord) throws Exception {
        if ("null".equals(passord)) {
            passord = null;
        }

        try (Workbook workbook = new ExcelRegneark(testfilnavn(filnavn), passord).openWorkbook()) {
            var klasselisteParser = new KlasselisteParser(workbook);
            List<Oppføring> oppføringer = klasselisteParser.lesElevOppføringer();

            assertThat(oppføringer).hasSize(1)
                    .satisfiesExactly(
                            forelder -> Testdata.klasseliste.getFirst()
                    );
        }
    }

    private String testfilnavn(String filnavn) throws FileNotFoundException {
        URL resource = this.getClass().getResource(filnavn);
        if (resource == null) {
            throw new FileNotFoundException(filnavn);
        }
        return resource.getFile();
    }
}
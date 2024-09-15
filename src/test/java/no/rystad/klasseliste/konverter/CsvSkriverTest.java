package no.rystad.klasseliste.konverter;

import org.junit.jupiter.api.Test;

import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStreamReader;
import java.util.List;

import static no.rystad.klasseliste.konverter.CsvSkriver.formaterTelefonnummer;
import static org.assertj.core.api.Assertions.assertThat;

class CsvSkriverTest {
    @Test
    void skrivForesatteSomCsvForKontaktImport_medTestdata_skrivesRiktig() {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        CsvSkriver csvSkriver = new CsvSkriver(os);

        csvSkriver.skrivForesatteSomCsvForKontaktImport(Testdata.klasseliste);
        BufferedReader reader = new BufferedReader(new InputStreamReader(new ByteArrayInputStream(os.toByteArray())));
        List<String> linjer = reader.lines().toList();

        assertThat(linjer).satisfiesExactly(
                linje -> assertThat(linje).isEqualTo("First Name,Last Name,Telephone,Email"),
                linje -> assertThat(linje).isEqualTo("Mamma Testesen,(Test Testesen 1A),+4790807060,mamma@test.com"),
                linje -> assertThat(linje).isEqualTo("Pappa Testesen,(Test Testesen 1A),98765432,pappa@test.com")
        );
    }

    @Test
    void formaterTelefonnummer_formatererRiktig() {
        assertThat(formaterTelefonnummer("12345678")).isEqualTo("12345678");
        assertThat(formaterTelefonnummer("123 45 678")).isEqualTo("12345678");
        assertThat(formaterTelefonnummer("4712345678")).isEqualTo("+4712345678");
        assertThat(formaterTelefonnummer("+4712345678")).isEqualTo("+4712345678");
        assertThat(formaterTelefonnummer("46111222333")).isEqualTo("+46111222333");
        assertThat(formaterTelefonnummer("+46111222333")).isEqualTo("+46111222333");
        assertThat(formaterTelefonnummer(null)).isEqualTo("");
    }
}

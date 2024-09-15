package no.rystad.klasseliste.konverter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Iterator;
import java.util.List;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

class KlasselisteParser {
    private final Workbook workbook;

    KlasselisteParser(Workbook workbook) {
        this.workbook = workbook;
    }

    List<Oppføring> lesElevOppføringer() {
        Sheet exportSheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = exportSheet.rowIterator();
        verifyHeaders(rowIterator.next());

        return iteratorToStream(rowIterator)
                .filter(row -> row.getCell(0) != null)
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

}

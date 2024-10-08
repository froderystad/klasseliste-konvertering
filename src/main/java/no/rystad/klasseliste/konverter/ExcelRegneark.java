package no.rystad.klasseliste.konverter;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.security.GeneralSecurityException;

class ExcelRegneark {
    private String filnavn;
    private String passord;

    ExcelRegneark(String filnavn, String passord) {

        this.filnavn = filnavn;
        this.passord = passord;
    }

    Workbook openWorkbook() throws IOException, GeneralSecurityException, FeilPassord, IkkeStøttetFilendelse {
        var filendelse = filnavn.substring(filnavn.indexOf('.'));
        if (filendelse.endsWith("xlsx")) {
            return readNewFormat();
        } else if (filendelse.endsWith("xls")) {
            return readOldFormat();
        } else {
            throw new IkkeStøttetFilendelse(filendelse);
        }
    }

    private Workbook readNewFormat() throws IOException, GeneralSecurityException, FeilPassord {
        try (InputStream fis = new FileInputStream(filnavn);
             POIFSFileSystem poifs = new POIFSFileSystem(fis)) {
            Decryptor decryptor = getDecryptor(poifs);
            return new XSSFWorkbook(decryptor.getDataStream(poifs));
        } catch (OfficeXmlFileException e) {
            // ikke kryptert
            try (InputStream fis = new FileInputStream(filnavn)) {
                return new XSSFWorkbook(fis);
            }
        }
    }

    private Decryptor getDecryptor(POIFSFileSystem poifs) throws IOException, GeneralSecurityException, FeilPassord {
        EncryptionInfo encryptionInfo = new EncryptionInfo(poifs);
        Decryptor decryptor = Decryptor.getInstance(encryptionInfo);

        if (!decryptor.verifyPassword(passord)) {
            throw new FeilPassord();
        }
        return decryptor;
    }

    private Workbook readOldFormat() throws IOException {
        try (InputStream fis = new FileInputStream(filnavn);
             POIFSFileSystem poifs = new POIFSFileSystem(fis)) {
            Biff8EncryptionKey.setCurrentUserPassword(passord);
            return new HSSFWorkbook(poifs);
        } finally {
            Biff8EncryptionKey.setCurrentUserPassword(null);
        }

    }

    static class FeilPassord extends Exception {
    }

    static class IkkeStøttetFilendelse extends Exception {
        private final String filendelse;

        IkkeStøttetFilendelse(String filendelse) {
            this.filendelse = filendelse;
        }

        String filendelse() {
            return filendelse;
        }
    }
}

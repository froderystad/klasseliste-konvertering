# Klasseliste-konverterer

Dette programmet konverterer klasselister fra Excel-fil til et CSV-format som kan importeres i Google Kontakter.

"Nytt" XLSX-format og gammelt XLS-format støttes, både med og uten passord/kryptering.

### Inn-format

Inn-formatet har én linje per elev.

| Klasse | Etternavn | Fornavn | Navn forelder 1 | Telefon forelder 1 | E-post forelder 1 | Navn forelder 2 | Telefon forelder 2 | E-post forelder 2 |
|--------|-----------|---------|-----------------|--------------------|-------------------|-----------------|--------------------|-------------------|

### Ut-format

Ut-formatet har én linje per foresatt.

```text
First Name,Last Name,Telephone,Email
```

## Bruk

Du må ha installert [Java](https://adoptium.net/installation/) 21 eller nyere
og [Maven](https://maven.apache.org/install.html).

Kjør programmet og skriv til CSV-fil med:

```shell
mvn package
./konverter-klasseliste.sh "<full-sti-til-fil>" [<passord>] > klasseliste.csv
```

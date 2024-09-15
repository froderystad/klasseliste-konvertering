# Klasseliste-konverterer

Dette programmet konverterer klasseliste eksportert fra iSkole til et format som kan importeres i Google Kontakter.

"Nytt" XLSX-format og gammelt XLS-format støttes, både med og uten passord/kryptering.

Kun norske telefonnummer støttes.

## Bruk

Du må ha installert [Java](https://adoptium.net/installation/) 21 eller nyere
og [Maven](https://maven.apache.org/install.html).

Kjør programmet og skriv til CSV-fil med:
```shell
mvn package
./konverter-klasseliste.sh "<full-sti-til-fil>" [<passord>] > klasseliste.csv
```

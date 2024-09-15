# Klasseliste-konverterer

Dette programmet konverterer klasseliste eksportert fra iSkole til et format som kan importeres i Google Kontakter.

Kryptert XLSX-format og ukryptert XLS-format støttes. Dette kan utvides senere.

Kun norske telefonnummer støttes.

## Bruk

Du må ha installert [Java](https://adoptium.net/installation/) 21 eller nyere
og [Maven](https://maven.apache.org/install.html).

Kjør programmet med:
```shell
mvn package
./konverter-klasseliste.sh "<full-sti-til-fil>" [<passord>]
```

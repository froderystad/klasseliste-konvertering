#!/usr/bin/env sh

jar_base_name="konverter-klasseliste.jar"

if [ -f "./${jar_base_name}" ]; then
  jar_name="./${jar_base_name}"
else
  jar_name="./target/${jar_base_name}"
fi

if [ ! -f "${jar_name}" ]; then
  echo "Finner ikke ${jar_base_name}. Kjør 'mvn package' først."
  exit 1
fi

java -jar ${jar_name} "$@"

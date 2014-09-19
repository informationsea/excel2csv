#!/bin/sh

if [ -d /Library/Java/JavaVirtualMachines/jdk1.7.0_55.jdk/Contents/Home ];then
    export JAVA_HOME=/Library/Java/JavaVirtualMachines/jdk1.7.0_55.jdk/Contents/Home
fi

mvn3 clean
mvn3 install

rm -r excel2csv
mkdir -p excel2csv/bin excel2csv/lib
echo Copying jar
cp target/excel2csv-*-jar-with-dependencies.jar excel2csv/lib/excel2csv.jar||exit 1
echo Copying command
cp bin/excel2csv excel2csv/bin/excel2csv||exit 1
cp README.md excel2csv/README.txt
cp LICENSE excel2csv/LICENSE.txt
echo Making tar.gz
tar czf excel2csv.tar.gz excel2csv||exit 2



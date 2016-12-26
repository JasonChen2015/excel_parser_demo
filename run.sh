PROG=`pwd`
CODE="$PROG/code"
LIB="$CODE/lib"
XLS="$PROG/xls"
OUTPUT="$PROG/output"

# for generated class file
CP="$CODE"
# the poi
CP="$CP:$LIB/poi-3.15.jar:$LIB/poi-ooxml-3.15.jar:$LIB/poi-ooxml-schemas-3.15.jar"
CP="$CP:$LIB/xmlbeans-2.6.0.jar"
# xerces
CP="$CP:$LIB/xercesImpl.jar:$LIB/xml-apis.jar"

# run java
java -cp $CP J_Cvt "$XLS/test.xls" $OUTPUT "a,b,c,d,e" 5 1 HAHAHA

#java -cp $CP J_Value "$XLS/test.xlsx" b1 1


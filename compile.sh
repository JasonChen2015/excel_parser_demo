PROG=`pwd`
CODE="$PROG/code"
LIB="$PROG/lib"

# the poi
CP="$LIB/poi-3.15.jar:$LIB/poi-ooxml-3.15.jar:$LIB/poi-ooxml-schemas-3.15.jar"

echo "[`date`]prepare to compile J_Cvt.java..."
javac -Xlint -cp $CP "$CODE/J_Cvt.java"
if [ $? == 0 ]; then
    echo -e "[`date`]Done!\n\n"
else
    echo -e "[`date`]Compile error!\n\n"
    exit 1
fi

echo "[`date`]prepare to compile J_Value.java..."
javac -Xlint -cp $CP "$CODE/J_Value.java"
if [ $? == 0 ]; then
    echo -e "[`date`]Done!\n\n"
else
    echo -e "[`date`]Compile error!\n\n"
    exit 1
fi


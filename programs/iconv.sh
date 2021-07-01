#!/bin/env bash
sjis="charset=unknown-8bit"
filename="common.bas"
filepath="./"
echo ${filepath}${filename}
temp=$(file -i ${filepath}${filename} |awk '{print $3}')
echo ${temp}
if [ $temp = $sjis ]; then
  iconv -f SHIFT-JIS -t UTF-8 ${filepath}${filename} > ${filepath}temp.bas
  echo "SHIFT-JIS -> UTF-8"
else
  iconv -f UTF-8 -t SHIFT-JIS ${filepath}${filename} > ${filepath}temp.bas
  echo "UTF-8 -> SHIFT-JIS"
fi
cp ${filepath}temp.bas ${filepath}${filename}
rm ${filepath}temp.bas
filename="vpn_management.bas"
filepath="./"
echo ${filepath}${filename}
temp=$(file -i ${filepath}${filename} |awk '{print $3}')
echo ${temp}
if [ $temp = $sjis ]; then
  iconv -f SHIFT-JIS -t UTF-8 ${filepath}${filename} > ${filepath}temp.bas
  echo "SHIFT-JIS -> UTF-8"
else
  iconv -f UTF-8 -t SHIFT-JIS ${filepath}${filename} > ${filepath}temp.bas
  echo "UTF-8 -> SHIFT-JIS"
fi
cp ${filepath}temp.bas ${filepath}${filename}
rm ${filepath}temp.bas
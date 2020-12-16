#!/bin/sh

#rm ./gantt.png;
rm ./gantt.csv;
rm ~/*.log;
rm ./Report.xlsx;
rm ./Report.pdf;

# gantt.png
#/usr/bin/ganttproject -export png gantt.gan
#echo "\n"
#sleep 5

# gantt.csv
/usr/bin/ganttproject -export csv gantt.gan
echo "\n"
sleep 5

# Report.xlsx
/home/osboxes/.pyenv/shims/python ./gantt.py &&

# Report.pdf
/usr/bin/libreoffice --headless --convert-to pdf:calc_pdf_Export --outdir ./ Report.xlsx

rm ./gantt.csv;
rm ~/*.log;

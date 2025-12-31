# text_to_pdf
A simple python script that converts text files into pdf using Microsoft word

#usage

python text_to_pdf.py

#merge all files

python text_to_pdf.py --merge --output final.pdf

#merge with a specific order

python text_to_pdf.py --merge --order intro.txt chapter1.docx conclusion.txt --output final.pdf

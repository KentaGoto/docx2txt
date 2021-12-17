# coding: utf-8

import os
import pathlib
from docx import Document


def all_files(directory):
    for root, dirs, files in os.walk(directory):
        for file in files:
            yield os.path.join(root, file)

if __name__ == '__main__':
    s = input("Dir: ")
    root_dir = s.strip('\"')

    print("Processing...")
    print()

    for i in all_files(root_dir):
        print(i)
        p = pathlib.Path(i)

        if p.suffix == ".docx":
            t = str(p) + '.txt'
            docx = Document(p)
            f = open(t, 'w', encoding="utf-8")

            for j in docx.paragraphs:
                f.write(j.text + "\n")

            f.close()
    
    print()
    print("Done!")
    os.system("pause > nul")

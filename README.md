Client request: xlookup from source files to consolidated worksheet

Challenges: source file and destination file are heavily formatted and use data validation/drop-downs

Issues: slow (each xlookup value takes about a second, running for all 16 source/destination files takes about 90 mintues). Investigated doing this proecess with Python instead of VBA,
but Python is not good for heavily formatted Excel files (does not have Excel data model).

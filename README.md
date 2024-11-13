This set of tools enables easy creation of spell cards (63.5x88mm) for Pathfinder 1e. Since only I used it, no command line interface has been developed, but the application can be easily extended. I do not own any part of the game, this is just a tool to make playing easier and more fun

Workflow:

1. Select class and spells at the beginning of generate.py and QR.py
2. Generate spell cards and QR codes (generate.py and create_qr.py)
3. Merge spell cards (concat_output.py)
4. Check spell cards, and adjust layout  and text if necessary (e.g.: text overflow). QR code placeholder is not populated by default (explained later)
5. From PowerPoint, export the merged file as image (File/Export/Change file type/Export for print)
6. Open `card_layout.svg` in Inkscape, and populate card and QR places. This results in much cleaner image, than exporting from powerpoint
7. Print file to pdf (A4 size)
8. Concatenate pdf files

The workflow is divided like this, to make costumization easier. Many spells have too long descriptions, so it would require too much effort to automate the process completely. This way, after each step, the result can be easily adjusted manually.

Spell spreadsheet can be downloaded [here](https://www.d20pfsrd.com/magic/tools/spells-db/)

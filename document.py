import aspose.words as aw
# create document object
doc = aw.Document()
# create a document builder
builder = aw.DocumentBuilder(doc)
#add text
builder.write("Welcome home")
builder.write('\n')
builder.write('\n')
# create font
font = builder.font
font.size = 22
font.italic = True
font.name = "Calibri"


# paragraph formatting
paragraphFormat = builder.paragraph_format
paragraphFormat.first_line_indent = 9
paragraphFormat.alignment = aw.ParagraphAlignment.CENTER
paragraphFormat.keep_together = True

#add text
builder.writeln("Python is used for server-side web development, software development, mathematics, and system scripting, and is popular for Rapid Application Development and as a scripting or glue language to tie existing components because of its high-level, built-in data structures, dynamic typing, and dynamic binding.")
builder.write('\n\n')
builder.write("--------------------------------------------------------------------")
builder.write('\n')

# start table
table = aw.DocumentBuilder(doc)

#insert cell
builder.start_table();
builder.row_format.height = 30.0;
builder.cell_format.width = 100.0;
builder.insert_cell();
builder.write("R1C1");

builder.insert_cell();

builder.writeln("R1C2");

builder.end_row();

builder.insert_cell();
builder.write("R2C1");

builder.insert_cell();
builder.writeln("R2C2");
builder.end_row();

builder.insert_cell();
builder.write("R3C1");

builder.insert_cell();
builder.writeln("R3C2");
builder.end_row();

builder.end_table();

#insert image
builder.insert_image("Delivery 1.png")

#save document
doc.save("abcd.docx")
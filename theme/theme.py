import zipfile
import os

src = "theme_extracted"   # ← folder you unzipped into
out = "mytheme.thmx"

if not os.path.exists(src):
    print(f"ERROR: folder '{src}' not found. Current dir: {os.getcwd()}")
    print("Files here:", os.listdir("."))
else:
    files_found = []
    with zipfile.ZipFile(out, "w") as zout:
        for root, dirs, files in os.walk(src):
            for file in files:
                abs_path = os.path.join(root, file)
                arc_path = os.path.relpath(abs_path, src)
                compress = zipfile.ZIP_STORED if file == "[Content_Types].xml" else zipfile.ZIP_DEFLATED
                zout.write(abs_path, arc_path, compress_type=compress)
                files_found.append(arc_path)
    
    print(f"Written {len(files_found)} files to {out}:")
    for f in files_found:
        print(" ", f)
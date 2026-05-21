import zipfile
import os

path = r"C:\Users\熊凯鹏\Desktop\功能点展示.pptx"

print(f"文件大小: {os.path.getsize(path)} bytes")
print()

with zipfile.ZipFile(path, 'r') as z:
    files = z.namelist()
    print("PPT 内部文件结构:")
    for f in sorted(files):
        info = z.getinfo(f)
        print(f"  {f:50s}  {info.file_size:>8d} bytes")

    # 检查关键文件是否存在
    key_files = [
        '[Content_Types].xml',
        'ppt/presentation.xml',
        'ppt/slides/slide1.xml',
        'ppt/slideLayouts/slideLayout1.xml',
        'ppt/slideMasters/slideMaster1.xml',
    ]
    print()
    print("关键文件检查:")
    for kf in key_files:
        status = "OK" if kf in files else "MISSING"
        print(f"  {kf:45s}  {status}")

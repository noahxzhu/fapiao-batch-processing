import pymupdf

file_path = "/Users/noah/Downloads/invoice/60701018+463.66.pdf"

with pymupdf.open(file_path) as doc:
    page = doc.load_page(0)
    blocks = page.get_textpage().extractBLOCKS()
    print(blocks)
    for block in blocks:
        block_text = block[4]
        print(block_text)

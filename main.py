import sys, subprocess
subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'PyMuPDF'])
# subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'fitz'])
# subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'io'])
# subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'PIL'])
# import fitz
import PyMuPDF
import io
from PIL import Image

file = "/content/pdf_file.pdf"

# open the file
pdf_file = PyMuPDF.fitz.open(file)

# STEP 3
# iterate over PDF pages
for page_index in range(len(pdf_file)):

	# get the page itself
	page = pdf_file[page_index]
	image_list = page.getImageList()

	# printing number of images found in this page
	if image_list:
		print(
			f"[+] Found a total of {len(image_list)} images in page {page_index}")
	else:
		print("[!] No images found on page", page_index)
	for image_index, img in enumerate(page.getImageList(), start=1):

		# get the XREF of the image
		xref = img[0]

		# extract the image bytes
		base_image = pdf_file.extractImage(xref)
		image_bytes = base_image["image"]

		# get the image extension
		image_ext = base_image["ext"]

import sys, os, re, datetime, os.path, traceback, shutil, time, subprocess	# builtin
import cv2, docx, win32com.client, numpy as np, pdfrw						# 3rd-party
# pip3 install opencv-python python-docx pywin32 numpy reportlab pdfrw

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.colors import white, transparent, black
from pdfrw import PdfReader, PdfWriter, PageMerge

########################################################################################

def find_ghostscript_executable():
	gs = shutil.which("gswin64")
	if gs:
		return gs
	print('\nCould not find ghostscript executable.')
	print('Either install ghostscript or check the PATH environment variable.')
	sys.exit()

def create_pdf(input, output):
	util = win32com.client.Dispatch("Bullzip.PDFUtil")
	settings = win32com.client.Dispatch("Bullzip.PDFSettings")
	settings.PrinterName = util.DefaultPrinterName		# make sure we're controlling the right PDF printer
	
	statusFile = re.sub("\.[^.]+$", ".status", input)
	settings.SetValue("Output", output)
	settings.SetValue("ConfirmOverwrite", "no")
	settings.SetValue("AutoRotatePages", "none")
	settings.SetValue("ShowSaveAS", "never")
	settings.SetValue("ShowSettings", "never")
	settings.SetValue("ShowPDF", "no")
	settings.SetValue("ShowProgress", "no")
	settings.SetValue("ShowProgressFinished", "no")		# disable balloon tip
	settings.SetValue("StatusFile", statusFile)			# created after print job
	settings.WriteSettings(True)						# write settings to the runonce.ini
	util.PrintFile(input, util.DefaultPrinterName)   	# send to Bullzip virtual printer
	
	# wait until print job completes before continuing
	# otherwise settings file for next job may not be used
	timestamp = datetime.datetime.now()
	while( (datetime.datetime.now() - timestamp).seconds < 10):
		if os.path.exists(statusFile) and os.path.isfile(statusFile):
			error = util.ReadIniString(statusFile, "Status", "Errors", '')
			if error != "0":
				raise IOError("PDF was created with errors")
			time.sleep(0.3)
			os.remove(statusFile)
			return output
		time.sleep(0.1)
	raise IOError("PDF creation timed out")


def create_pngs(inputfile, dst):
	if not os.path.exists(dst):
		os.makedirs(dst)
	for i in os.listdir(dst):
		os.remove(os.path.join(dst, i))
	gs = find_ghostscript_executable()
	cmd = f'"{gs}" -dSAFER -dBATCH -dNOPAUSE -sDEVICE=png16m -r600 -sOutputFile="{os.path.join(dst, "out")}%d.png" "{inputfile}"'
	process = subprocess.run(cmd, capture_output=True, text=True, shell=True)
	if process.returncode == 0:
		print("\ncreated folder '%s', with %d images" % (dst, len(os.listdir(dst))))
	else:
		print("\nfailed to create PNGs")
	
	
def clean_docx(file, output):
	rgx = r'(?<=w:fill=")FF0000(?=")'
	doc = docx.Document(file)
	for e in doc.element.iter():	# iterate through all descendants of the entire xml tree
		if "shd" in e.tag:			# find any elements with shd in the tag name
			attrib = [k for k in e.keys() if "fill" in k][0]	# get the attribute key with "fill" in its name
			e.set(attrib, "auto")	# use the shd element with the target attribute to change the attribute value
		if "rect" in e.tag:
			attrib = [k for k in e.keys() if "fillcolor" in k][0]
			if "lime" in e.get(attrib):
				e.set(attrib, "auto")
	doc.save(output)
	return output


def getSizedWindow(title, target_height, img):
	cv2.namedWindow(title, cv2.WINDOW_NORMAL)
	aspect_ratio = (1.0 * img.shape[1] / img.shape[0])
	target_width = int(aspect_ratio * target_height)
	cv2.resizeWindow(title, target_width, target_height)
	

def remap(cnt, img_width, img_height, pdf_width, pdf_height):
	# contour example...   [ [[ 484  258]], [[ 484  297]], [[1510  297]], [[1510  258]] ]
	# a contour is a list of nested numpy arrays; each coord is a corner since our input has well-defined boxes
	newlist = [j.tolist() for i in cnt for j in i]
	# above example is now  [[484, 258], [484, 297], [1510, 297], [1510, 258]]
	
	# we need to remap the corners from their cv2 coordinates to the coordinates that reportlab uses
	min_x = min([x for x,y in newlist])
	max_x = max([x for x,y in newlist])
	min_y = min([y for x,y in newlist])
	max_y = max([y for x,y in newlist])
	
	a = (1.0*min_x/img_width)*pdf_width
	b = pdf_height - ((1.0*min_y/img_height)*pdf_height)
	w = (1.0*(max_x - min_x)/img_width)*pdf_width
	h = (1.0*(max_y - min_y)/img_height)*pdf_height
	# reportlab textfields require a top-left point and width and height
	return (a,b,w,-h)


def perform(inputfile):

	temp_folder = inputfile.replace(".docx", "_temp")
	if not os.path.exists(temp_folder):
		os.makedirs(temp_folder)
	
	inputpdf = create_pdf(inputfile, os.path.join(temp_folder, "input.pdf")) # pdf (with red boxes) to feed to poppler
	png_folder = os.path.join(temp_folder, "pngs")							 # temp output folder to store pngs
	create_pngs(inputpdf, png_folder)										 # poppler produces one png per page 
	images = [os.path.join(png_folder, i) for i in os.listdir(png_folder) if i.endswith(".png")]
	
	c = canvas.Canvas(
		filename = os.path.join(temp_folder, "canvas.pdf"),
		pagesize = A4)
	pdf_width, pdf_height = A4
	c.setFont("Courier", 14)
	form = c.acroForm
	
	labelcounter = 1
	for image in images:
		print(os.path.basename(image), ": ", end="")
		img = cv2.imread(image, -1)
		img_width, img_height = img.shape[1], img.shape[0]
		#print(img.shape, img.dtype)
		#print(pdf_width, pdf_height)		# 595.275590551 841.88976378
		#print(img_width, img_height)		# 1653 2339		
		cntimg = np.zeros(img.shape, np.uint8)
		
		lower = np.array([0, 0, 200])	# red channel is last because opencv uses BGR by default
		upper = np.array([0, 0, 255])
		boxes = cv2.inRange(img, lower, upper)
		
		contours, hierarchy = cv2.findContours(boxes.copy(), cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
		print(f"found {len(contours):2} red text boxes and ", end="")
		if len(contours) > 0:
			for i, cnt in enumerate(contours[::-1]):
				cv2.drawContours(cntimg, [cnt], 0, (0, 0, 255), -1)
				a, b, w, h = remap(cnt, img_width, img_height, pdf_width, pdf_height)
				#print(cnt, a, b, w, h)
				form.textfield(	# page 61 in the reportlab userguide
					name=str(labelcounter),
					tooltip=str(labelcounter),
					x=a+2,	# add a tiny bit of left padding
					y=b,
					width=w-2,
					height=h,
					fontSize=11,
					maxlen=int(w/6),
					borderWidth=0,
					borderColor=transparent,
					fillColor=transparent,
					textColor=black,
					forceBorder=False)
				labelcounter+=1	# we need to ensure every field on every page has a unique id
		
		lower = np.array([0, 200, 0])
		upper = np.array([0, 255, 0])
		chkboxes = cv2.inRange(img, lower, upper)
		
		contours, hierarchy = cv2.findContours(chkboxes.copy(), cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
		print(f"{len(contours):2} green check boxes")
		if len(contours) > 0:
			for i, cnt in enumerate(contours[::-1]):
				cv2.drawContours(cntimg, [cnt], 0, (0, 255, 0), -1)
				a, b, w, h = remap(cnt, img_width, img_height, pdf_width, pdf_height)
				#print(cnt, a, b, w, h)
				form.checkbox(	# page 61 in the reportlab userguide
					name=str(labelcounter),
					tooltip=str(labelcounter),
					x=a,
					y=b+h,
					size=w,
					borderWidth=0,
					borderColor=white,
					fillColor=white,	# BUG: for some reason checkboxes don't work properly when using transparent
					textColor=black,
					forceBorder=False)
				labelcounter+=1	# we need to ensure every field on every page has a unique id
		
		c.showPage() # starts new canvas page
#		getSizedWindow("cntimg", 1000, cntimg)
#		cv2.imshow("cntimg", cntimg)
#		cv2.waitKey(0)
#		cv2.destroyAllWindows()
		
	c.save()
	cleandocx = clean_docx(inputfile, os.path.join(temp_folder, "clean.docx"))
	cleanpdf = create_pdf(cleandocx, os.path.join(temp_folder, "clean.pdf"))
	
	interactive = PdfReader(c._filename)			# reads the blank pdf with the interactive fields
	visible = PdfReader(cleanpdf)					# reads the compact pdf with the visible non-interactive form
	for i, p in enumerate(interactive.pages):
		PageMerge(p).add(visible.pages[i]).render() # merge them together
	
	output = inputfile.replace(".docx", ".pdf")
	PdfWriter(output, trailer=interactive).write()
	print("\ncreated file '%s'" % output)
	os.startfile(output)
	time.sleep(5)
	shutil.rmtree(temp_folder)
	print("\nremoved temp folder)

########################################################################################

if __name__ == '__main__':
	if  len(sys.argv) != 2:
		print("To use this script you need to pass it a single docx file.")
		print("You can do that on that command line or drag and drop")
		print("your docx file on to this script's icon.")
	else:
		try:
			find_ghostscript_executable()
			if sys.argv[1].endswith(".docx"):
				perform(sys.argv[1])
			else:
				print("\nThis script is for docx files only.")
		except SystemExit:
			pass
		except PermissionError:
			print("\nFailed to delete the temp folder.")
		except:
			print("\nScript failed.\n")
			traceback.print_exc()
	input("\nPress any key to exit...")

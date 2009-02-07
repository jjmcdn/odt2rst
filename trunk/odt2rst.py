import sys
import re
import os
import shutil
import zipfile
import hashlib
import math
import getopt
import xml.etree.ElementTree

office_prefix = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"
text_prefix = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"
table_prefix = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
drawing_prefix = "{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}"
xlink_prefix = "{http://www.w3.org/1999/xlink}"

debug_flag = False


def unpackOdt(input_path, temp_folder = "."):
	"Unpack the odt file into the temp folder and return a dictionary translating .png file path into they hashes."
	odtfile = zipfile.ZipFile(input_path)

	try:
		os.mkdir(temp_folder)
	except:
		pass

	try:
		os.mkdir(os.path.join(temp_folder, "Pictures"))
	except:
		pass

	odt_pictures_hashes = {}
	for path in odtfile.namelist():
		if path.lower() == "content.xml".lower():
			g = open(os.path.join(temp_folder, path), "wb")
			bytes = odtfile.read(path)
			g.write(bytes)
			g.close()

		folder, name = os.path.split(path)
		name, ext = os.path.splitext(name)
		if folder.lower() == "Pictures".lower() and ext in [".png", ".jpg"]:
			g = open(os.path.join(temp_folder, path), "wb")
			bytes = odtfile.read(path)
			g.write(bytes)
			g.close()

			h = hashlib.md5()
			h.update(bytes)
			h = h.digest()
			odt_pictures_hashes[path] = h

	return odt_pictures_hashes


def cleanPack(temp_folder = "."):
	"Delete the files and folder created by unpackOdt apart from the temp_folder itself."
	os.remove(os.path.join(temp_folder, "content.xml"))
	shutil.rmtree(os.path.join(temp_folder, "Pictures"))


def getHashesRstImages(output_folder, images_relative_folder):
	"Return a dictonary translating hash into the its .png file path."
	hashes_rst_images = {}

	image_folder = os.path.join(output_folder, images_relative_folder)

	if not os.path.isdir(image_folder):
		return {}

	for path in os.listdir(image_folder):
		name, ext = os.path.splitext(path)
		if ext not in [".png", ".jpg"]:
			continue

		path = os.path.join(images_relative_folder, path)

		f = open(path, "rb")
		bytes = f.read()
		f.close()

		h = hashlib.md5()
		h.update(bytes)
		h = h.digest()
		hashes_rst_images[h] = path

	return hashes_rst_images


def synchronizeImagesFolders(temp_folder, output_path, images_relative_folder, odt_pictures_hashes):
	output_folder, output_name = os.path.split(output_path)
	image_folder = os.path.join(output_folder, images_relative_folder)

	hashes_rst_images = getHashesRstImages(output_folder, images_relative_folder)

	# Build the picture_dict that convert odt image path into rst image path (when possible)
	picture_prefix = "picture_"
	existing_picture_names = []
	if os.path.isdir(image_folder):
		for path in os.listdir(image_folder):
			path = path.lower()
			name, ext = os.path.splitext(path)
			if ext not in [".png", ".jpg"]:
				continue

			if not name.startswith(picture_prefix):
				continue

			existing_picture_names.append(name)

	picture_dict = {}
	picture_index = 0
	for path in odt_pictures_hashes:
		h = odt_pictures_hashes[path]
		if h in hashes_rst_images:
			picture_dict[path] = hashes_rst_images[h]
		else:
			# Find an available picture name:
			while picture_prefix + str(picture_index) in existing_picture_names:
				picture_index += 1

			picture_name = picture_prefix + str(picture_index)
			existing_picture_names.append(picture_name)

			name, ext = os.path.splitext(path)
			picture_relative_path = os.path.join(images_relative_folder, picture_name) + ext

			if not os.path.isdir(image_folder):
				os.mkdir(image_folder)

			shutil.copyfile(os.path.join(temp_folder, path), os.path.join(output_folder, picture_relative_path))

			picture_dict[path] = picture_relative_path

	return picture_dict


def splitIntoLines(text):
	text = re.sub(r"([a-zA-Z\"']{2})\. ", r"\1.\n", text)

	return text

def getRawText(node):
	text = ""
	if node.text:
		text += node.text
	for child in node:
		if child.tag in [text_prefix + "p", text_prefix + "span"]:
			text += getRawText(child)

		if child.tail:
			text += child.tail

	text = text.replace("\n", " ")
	return text


def getCodeText(node):
	text = ""
	if node.text:
		text += node.text
	for child in node:
		if child.tag in [text_prefix + "p", text_prefix + "span"]:
			text += getCodeText(child)

		if child.tag == text_prefix + "line-break":
			text += "\n"

		if child.tag == text_prefix + "s":
			identation = int(child.attrib[text_prefix + "c"])
			identation = " " * identation

			text += identation

		if child.tail:
			text += child.tail

	return text

	
def escapeCellText(text):
	"Return a rst version of the text that is suitable for rst cell content."
	text = text.replace("+", "\\+")
	text = text.replace("-", "\\-")
	text = text.replace("|", "\\|")
	return text
	
	
class Table:
	def __init__(self):
		self.rows = []

	def __str__(self):
		ret = "Table(\n"
		ret += "  rows : [\n"
		for row in self.rows:
			ret += str(row) + ",\n"
		ret += "  ]\n"
		ret += ")\n"

		return ret

	def addCoveredCells(self):
		num_columns = 0
		row  = self.rows[0]
		for cell in row.cells:
			num_columns += cell.h_span

		grid = [[None for i in range(num_columns)] for j in range(len(self.rows))]

		for row_index in range(len(self.rows)):
			row = self.rows[row_index]

			column_index = 0
			for cell in row.cells:
				while column_index < num_columns and grid[row_index][column_index]:
					column_index += 1

				grid[row_index][column_index] = cell
				for extra_column_index in range(cell.h_span):
					for extra_row_index in range(cell.v_span):
						if extra_column_index == 0 and extra_row_index == 0:
							continue
						covered_cell = TableCell()
						grid[row_index + extra_row_index][column_index + extra_column_index] = covered_cell
						covered_cell.covered = True

						if extra_column_index:
							covered_cell.left_wall = False

						if extra_row_index:
							covered_cell.top_wall = False

		for row_index in range(len(self.rows)):
			row = self.rows[row_index]

			row.cells = grid[row_index]

	def getColumnWidths(self):
		column_widths = []

		for row in self.rows:
			column_index = 0
			while column_index < len(row.cells):
				cell = row.cells[column_index]

				if cell.covered:
					column_index += 1
					continue

				while len(column_widths) < column_index + cell.h_span:
					column_widths.append(0)

				actual_width = sum(column_widths[column_index : column_index + cell.h_span]) + cell.h_span - 1
				width = len(cell.text) + 2
				if width > actual_width:
					addition = int(math.ceil((width - actual_width) / float(cell.h_span)))
					for index in range(column_index, column_index + cell.h_span):
						column_widths[index] += addition

				column_index += cell.h_span

		return column_widths


class TableRow:
	def __init__(self):
		self.header = False
		self.cells = []

	def __str__(self):
		ret = "Row(\n"
		ret += "  rows : [\n"
		for cell in self.cells:
			ret += str(cell) + ",\n"
		ret += "  ]\n"
		ret += ")\n"

		return ret


class TableCell:
	def __init__(self):
		self.h_span = 1
		self.v_span = 1

		self.text = ""

		self.covered = False
		self.top_wall = True
		self.left_wall = True

	def __str__(self):
		ret = "Cell(\n"
		ret += "  h_span : %d,\n" % self.h_span
		ret += "  v_span : %d,\n" % self.v_span
		ret += '  text : "%s",\n' % self.text
		ret += '  covered : %d,\n' % self.covered
		ret += ")\n"

		return ret
	

class RstDocument:
	# Set here the char that should be used to underline the titles according to they levels.
	# The default is the Python convention for documentation.
	levelchars = ["#", "*", "=", "-", "^", '"']
	identation_string = "   "

	def __init__(self, path = ""):
		self.path = path
		self.file = None

		self.list_levels = []
		self.list_indexes = []

		self.picture_dict = {}

		self.inline_images = {}
		self.paragraphs = []

	def flush(self):
		for paragraph in self.paragraphs:
			text = paragraph
			if debug_flag:
				text += "endof para"
			text += "\n"
			self.file.write(text)

		self.paragraphs = []

	def open(self, path = ""):
		if path:
			self.path = path
		self.file = open(self.path, "w")

	def close(self):
		self.flush()

		for path in self.inline_images:
			name = self.inline_images[path]
			self.file.write("\n.. |%s| image:: %s\n" % (name, path))

		self.file.close()

	def write(self, s):
		self.flush()

		s = s.encode("utf8")
#		if s == "\n":
#			raise Exception("hidden return")
		self.file.write(s)

	def writeTitle(self, text, level):
		paragraph = ""
		if debug_flag:
			paragraph += "pre-title"
		paragraph += "\n"
		char = self.levelchars[level]
		paragraph += "\n" + text + "\n" + char * len(text) + "\n"
		self.write(paragraph)

	def writeParagraph(self, text):
		if not text:
			return

		if text.startswith("Unknown interpreted text role"):
			return

		paragraph = ""
		if not self.list_levels:
			if debug_flag:
				paragraph += "pre-para"
			paragraph += "\n"

		elif not self.list_levels[-1]:
			if debug_flag:
				paragraph += "pre-item"
			paragraph += "\n"

		identation = ""
		bullet = ""
		non_bullet = ""
		if self.list_levels:
			if self.list_indexes[-1] >= 0:
				bullet = "    "
				non_bullet = "    "
				if self.list_levels[-1]:
					bullet = " %d. " % (self.list_indexes[-1] % 10)
			else:
				bullet = "   "
				non_bullet = "   "
				if self.list_levels[-1]:
					bullet = " - "

			self.list_levels[-1] = False

			for i in self.list_indexes[:-1]:
				if i >= 0:
					identation += "    "
				else:
					identation += "   "
			#identation = "   " * (len(self.list_levels) - 1)

		text = splitIntoLines(text)
		text = text.split("\n")

		paragraph += identation + bullet + ("\n" + identation + non_bullet).join(text)

		self.paragraphs.append(paragraph)

	def writeDefinitionBody(self, text):
		if self.paragraphs:
			paragraph = self.paragraphs.pop()

			paragraph = paragraph.replace("**", "")
			self.write(paragraph)

		text = splitIntoLines(text)
		text = text.split("\n")

		self.write("\n")

		identation = self.identation_string
		text = identation + ("\n" + identation).join(text) + "\n"
		self.write(text)

	def writeCodeBlock(self, text):
		if self.paragraphs:
			paragraph = self.paragraphs[-1]
			paragraph += ":\n"
			self.paragraphs[-1] = paragraph

		identation = self.identation_string
		text = text.split("\n")
		self.write(identation + ("\n" + identation).join(text) + "\n")

	def writeNoteHeader(self):
		self.write("\n.. note::\n")

	def appendToNote(self, text):
		text = splitIntoLines(text)
		text = text.split("\n")
		self.write("   " + "\n   ".join(text) + "\n\n")

	def writeWarningHeader(self):
		self.write("\n.. warning::\n")

	def appendToWarning(self, text):
		self.write("   " + text + "\n\n")

	def writeImage(self, path):
		self.write("\n")
		if path in self.picture_dict:
			path = self.picture_dict[path]
		path = path.replace('\\', '/')
		self.write(".. image:: %s\n" % path)

	def writeFigure(self, path, legend):
		self.write("\n")
		if path in self.picture_dict:
			path = self.picture_dict[path]
		path = path.replace('\\', '/')
		self.write(".. figure:: %s\n\n" % path)
		if legend:
			self.write("   %s\n" % legend)

	def writeComment(self, text):
		text = text.split("\n")
		self.write("\n.. " +  text.pop(0)+ "\n")
		while text:
			self.write("   " + text.pop(0) + "\n")
		self.write("\n")
		
	def writeTable(self, table):
		self.write("\n")

		table.addCoveredCells()
		column_widths = table.getColumnWidths()

		bottom = ""
		previous_header = False
		for row_index in range(len(table.rows)):
			row = table.rows[row_index]

			top = ""
			body = ""
			column_index = 0
			while column_index < len(row.cells):
				cell = row.cells[column_index]

				if cell.covered:
					while column_index < len(row.cells):
						cursor_cell = row.cells[column_index]
						if not cursor_cell.covered:
							break

						top_char = " "
						if cursor_cell.top_wall:
							if previous_header:
								top_char = "="
							else:
								top_char = "-"

						cross_char = "+"
						if not cursor_cell.top_wall and not cursor_cell.left_wall:
							cross_char = " "

						top +=  cross_char + top_char * column_widths[column_index]

						wall_char = " "
						if cursor_cell.left_wall:
							wall_char = "|"
						body += wall_char + " " * column_widths[column_index]

						column_index += 1

				else:
					for cursor_column_index in range(column_index, column_index + cell.h_span):
						cursor_cell = row.cells[cursor_column_index]
						top_char = " "
						if cursor_cell.top_wall:
							if previous_header:
								top_char = "="
							else:
								top_char = "-"

						top +=  "+" + top_char * column_widths[cursor_column_index]

					width = sum(column_widths[column_index : column_index + cell.h_span]) + cell.h_span - 1
					body += "|" + " " + cell.text + " " * (width - len(cell.text) - 2) + " "

					column_index += cell.h_span


			top += "+\n"
			if row_index == 0:
				bottom = top

			body += "|\n"

			self.write(top)
			self.write(body)
			
			previous_header = row.header

		self.write(bottom)		

	def getElementText(self, node):
		text = ""
		if node.text:
			text += node.text
		for child in node:
			if child.tag == text_prefix + "span":
				if child.attrib[text_prefix + "style-name"] in ["rststyle-emphasis"]:
					text += "*%s*" % child.text

				elif child.attrib[text_prefix + "style-name"] in ["rststyle-strong"]:
					text += "**%s**" % child.text
					
				elif child.attrib[text_prefix + "style-name"] in ["rststyle-inlineliteral"]:
					text += "``%s``" % child.text

				else:
					text += child.text

			elif child.tag == drawing_prefix + "frame":
				if child[0].tag == drawing_prefix + "image":
					image = child[0]
					path = image.attrib[xlink_prefix + "href"]
					if path in self.picture_dict:
						path = self.picture_dict[path]
					path = path.replace('\\', '/')

					folder, name = os.path.split(path)
					name, ext = os.path.splitext(name)

					text += "|%s|" % name

					self.inline_images[path] = name

			elif child.tag == text_prefix + "p":
				text += self.getElementText(child)

			else:
				print child.tag

			if child.tail:
				text += child.tail

		text = text.replace("\n", " ")
		return text

	def transformTableNode(self, table_node):
		table = Table()
		column_sizes = []

		for child in table_node:
			header = False
			row_node = child
			if child.tag == table_prefix + "table-header-rows":
				header = True
				row_node = child[0]

			if row_node.tag != table_prefix + "table-row":
				continue

			row = TableRow()
			row.header = header
			table.rows.append(row)

			for cell_node in row_node:
				cell = TableCell()
				if cell_node.tag != table_prefix + "table-cell":
					continue
				row.cells.append(cell)

				cell.h_span = int(cell_node.attrib.get(table_prefix + "number-columns-spanned", 1))
				cell.v_span = int(cell_node.attrib.get(table_prefix + "number-rows-spanned", 1))
				cell.text = self.getElementText(cell_node[0])
				cell.text = escapeCellText(cell.text)

		self.writeTable(table)

	def transformNode(self, node):
		for child in node:
			if child.tag == text_prefix + "p":
				style = child.attrib[text_prefix + "style-name"]
				frame = child.find(drawing_prefix + "frame")
				comment = child.find(office_prefix + "annotation")

				if style == "rststyle-title":
					self.writeTitle(child.text, 0)

				elif style == "rststyle-admon-note-hdr":
					self.writeNoteHeader()

				elif style == "rststyle-admon-note-body":
					self.appendToNote(self.getElementText(child))

				elif style == "rststyle-admon-warning-hdr":
					self.writeWarningHeader()

				elif style == "rststyle-admon-warning-body":
					self.appendToWarning(self.getElementText(child))

				elif style == "rststyle-blockindent":
					self.writeDefinitionBody(self.getElementText(child))

				elif style == "rststyle-codeblock":
					self.writeCodeBlock(getCodeText(child))

				elif frame and frame.attrib[text_prefix + "anchor-type"] == "paragraph":
					if frame[0].tag == drawing_prefix + "image":
						image = frame[0]
						path = image.attrib[xlink_prefix + "href"]
						self.writeImage(path)

					elif frame[0].tag == drawing_prefix + "text-box":
						try:
							text_box = frame[0]
							paragraph = text_box[0]
							frame = paragraph[0]
							image = frame[0]
							path = image.attrib[xlink_prefix + "href"]
							legend = frame.tail

							self.writeFigure(path, legend)
						except:
							print "fail to convert the figure"

				elif comment:
					try:
						text = getRawText(comment)

						self.writeComment(text)
					except:
						print "fail to find the comment"

				else:
					#self.write("\n")
					self.writeParagraph(self.getElementText(child))

			if child.tag == text_prefix + "h":
				level = int(child.attrib[text_prefix + "outline-level"])
				self.writeTitle(self.getElementText(child), level)

			if child.tag == text_prefix + "section":
				self.transformNode(child)

			if child.tag == text_prefix + "list":
				if child.attrib.get(text_prefix + "style-name", "") == "Outline":
					item = child[0]
					while item._children:
						if item[0].tag == text_prefix + "h":
							self.transformNode(item)
							break;
						item = item[0]

				else:
					if child.attrib[text_prefix + "style-name"] == "rststyle-blockquote-enumlist":
						self.list_indexes.append(0)
					else:
						self.list_indexes.append(-1)

					self.write("\n")
					self.list_levels.append(True)
					self.transformNode(child)
					#self.write("\n")
					self.list_levels.pop()
					self.list_indexes.pop()

			if child.tag == text_prefix + "list-item":
				self.list_levels[-1] = True # Make sure the first paragraph get its bullet mark.

				# Update the item index of the item:
				if self.list_indexes[-1] >= 0:
					self.list_indexes[-1] += 1

				self.transformNode(child)
				#self.writeListItem(self.getElementText(child))

			if child.tag == table_prefix + "table":
				self.transformTableNode(child)
				
	def transform(self, content_path, styles_path, picture_dict):
		self.picture_dict = picture_dict
		
		parser = xml.etree.ElementTree.XMLTreeBuilder()
		doc = xml.etree.ElementTree.parse(content_path, parser)
	
		body = doc.find(office_prefix + "body")
		text = body.find(office_prefix + "text")

		self.open()
		self.transformNode(text)
		self.close()


def odt2rst(input_path, output_path, images_relative_folder = "images", temp_folder = ".", clean = True):
	odt_pictures_hashes = unpackOdt(input_path, temp_folder)

	picture_dict = synchronizeImagesFolders(temp_folder, output_path, images_relative_folder, odt_pictures_hashes)

	content_path = os.path.join(temp_folder, "content.xml")
	styles_path = os.path.join(temp_folder, "styles.xml")

	rst_document = RstDocument(output_path)
	rst_document.transform(content_path, styles_path, picture_dict)

	if clean:
		cleanPack(temp_folder)


def version():
	print "1.0"


def help():
	print "odt2rst.py [--images images-folder] [--temp temp-folder] odtfile [rstfile]"


def main():
	opts, args = getopt.getopt(sys.argv[1:], "vh", ["version", "help", "do-not-clean", "images=", "temp="])
	images_relative_folder = "images"
	temp_folder = "."
	clean = True
	for o, v in opts:
		if o in ["-v", "--version"]:
			version()
			return

		if o in ["-h", "--help"]:
			help()
			return

		if o in ["--images"]:
			images_relative_folder = v

		if o in ["--temp"]:
			temp_folder = v

		if o in ["--do-not-clean"]:
			clean = False

	input_file = ""
	if len(args) >= 1:
		input_file = args[0]

	if input_file == "":
		usage()

	name, ext = os.path.splitext(input_file)
	output_file = name + ".rst"
	if len(args) >= 2:
		output_file = args[1]

#	print "input", input_file
#	print "output:", output_file
#	print "temp:", temp_folder
#	print "images:", images_relative_folder

	odt2rst(input_file, output_file, images_relative_folder = images_relative_folder, temp_folder = temp_folder, clean = clean)


if __name__ == "__main__":
	main()
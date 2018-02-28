#!/usr/bin/env python

"""
makeenv.py: Makes envelopes from Excel spreadsheet.

Assumes there is an .xlsx file in the same directory as makeenv.py
and maybe a file called env.conf for configuration

.conf:
1. Path to the excel spreadsheet with the envelope data. String: path.
2. The sheet within the spreadsheet where the pertinent data is located. Int: sheetNo.
3. The range of rows with the data. List of 2 ints: [rowStart, rowEnd].
4. The columns of the rows with the data. List of 4 chars: [name, street, city, country].
5. The dimensions of the physical envelope in inches: List of 3 floats: [height, width, margin].
6. The output filename. String: filename.

usage:
python makeenv.py [-out] [[pth=]input.xlsx] [env.conf] [sht=2] [row=[5,129]] [col=['B','E','F','G']] [dim=[5.25,7.25,1]] [output.tex | output.pdf]
"""

import xlrd

class Envelopes(object):
	def __init__(self, source, sheet, rows, cols, dims):
		self.source = source
		self.sheet = int(sheet) - 1
		self.rows = (self.start, self.end) = rows
		self.cols = (self.name, self.street, self.city, self.country) = self.convertCols(cols)
		self.dims = (self.height, self.width, self.margin) = dims
		try:	
			self.tex = self.excel_to_TeX(True, True)
		except:	
			print("Fatal Error, exiting...")
			exit()
		
	def convertCols(self, columnTuple):
		"""	Converts tuple of chars ('A','B','D') into tuple of array indexes (0,1,3). 
		only supports columns A-Z. 	"""
		
		o = []
		for col in columnTuple:
			o.append(ord(col.upper())-65)
		return tuple(o)
		
	def sanitize(self, unwashed_string):
		"""	Converts raw text to TeX friendly formatting. """
		s = unwashed_string
		ss = unwashed_string.split(' ')
		
		# add escape to special characters
		for c in ['#','&']:
			if c in s:
				s = s.replace(c, '\\'+c)
				
		# detect numbers within address and format correctly
		if ss[0].isdigit() or (ss[0].split('-'))[0].isdigit():
			
			for w in ss:
				if w !='' and w[0].isdigit():				
					if w[-2:] in ['st', 'nd', 'rd', 'th']:
						s = s.replace(w, '$'+w[:-2]+'^{'+w[-2:]+'}$')
			
		return s.lstrip()
	
	def excel_to_TeX(self, returnaddress=False, stamp=False):
		""" Uses data from an excel spreadsheet to generate LaTeX. """

		head =\
		"\\documentclass[12pt]{article}\n\n" +\
		"\\usepackage[paperheight="+str(self.height)+"in,paperwidth="+str(self.width)+"in,margin="+str(self.margin)+"in,nofoot,nohead]{geometry}\n\n" +\
		"\\begin{document}\n\n" +\
		"\\pagestyle{empty}\n" +\
		"\\setlength{\\unitlength}{1in}\n\n"
		
		body = ""
				
		tail = "\\end{document}"
		
		try:
			book = xlrd.open_workbook(self.source)
		except:	
			print("Failed to open Excel spreadsheet. Exiting...")
			exit()
		
		sh = book.sheet_by_index(self.sheet)

		for i in range(self.start-1, self.end):

			r = sh.row_values(i)
			
			if (r[self.street] != '?'):
				name = self.sanitize(r[self.name]) + "\\\\"
				street = self.sanitize(r[self.street]) + "\\\\"
				city = self.sanitize(r[self.city]) + "\\\\"
				country = self.sanitize(r[self.country])
				body += "\\begin{minipage}{.5\\linewidth} \\noindent\n"
				if returnaddress:	
					body += "Sean Evans \& Lauren Demell\\\\\ \n" +\
					"134 Crescent Lane\\\\ \n" +\
					"Roslyn Heights NY 11577\n"
				body += "\\end{minipage}\n" +\
				"\\begin{minipage}{.5\\linewidth \\hspace{-.2in} \\vspace{-.3in}}\n"
				if stamp:
					body += "\\begin{flushright}\n" +\
					"\\framebox(1,1){STAMP}\n" +\
					"\\end{flushright}\n"
				body += "\\end{minipage}\n\n"
				body += "\\begin{center} \\begin{Huge} \\vspace*{\\fill}\n" +\
				name + "\n" +\
				street + "\n" +\
				city + "\n" 
				#if country != "United States":	body += country + "\n"
				body += "\\vspace{\\fill} \\end{Huge} \\end{center}\n\n"
				body += "\\clearpage\n\n"
			
		tex = head + body + tail
		return(tex)
		
	def generate_TeX(self, output_filename):
		""" Generates a .tex file with the contents of self.tex """
		try:
			with open(output_filename,'w') as g:			
				g.write(self.tex)
			print(output_filename + " created successfully.")
			return True
		except:
			print(output_filename + " NOT created.")
			return False

	def generate_PDF(self, pdf_filename):
		""" Generates a .tex and a .pdf.
		note: only works if texify.exe is present. """
		if(self.generate_TeX(pdf_filename[:-4]+".tex")):
			try:
				command = "texify " + pdf_filename[:-4] + ".tex -p -c -q"	# Windows 10 with MikTeX
				os.system(command)
				print(pdf_filename + " created successfully.")
				return True
			except:
				print(pdf_filename + " NOT created.")
				return False
				
def unpack(configuration_file):
	""" Unpacks .conf file into a dict """
	
	print("using " + configuration_file + "...")
	try:
		with open(configuration_file, 'r') as conf:
			s = conf.read().splitlines()
	except:
		print("Configuration file could not be opened, exiting...")
		exit()
	
	try:
		s[1] = int(s[1])										# sheet
		s[2] = list(map(int, s[2][1:-1].split(',')))			# rows
		s[3] = list([b[1:-1] for b in s[3][1:-1].split(',')])	# cols
		s[4] = list(map(float, s[4][1:-1].split(',')))			# dims
	except:
		print("Configuration file is invalid, exiting...")
		exit()
	
	t = {"path":s[0], "sheet":s[1], "rows":s[2], "cols":s[3], "dims":s[4], "out_file":s[5]}
	return t
	
def createDefaultConf():
		""" Re-writes env.conf to working default values. """
		with open("env.conf","w") as e:
				e.write("marry.xlsx\n")			# path
				e.write("2\n")					# sheet
				e.write("[5, 129]\n")			# rows
				e.write("['B','E','F','G']\n")	# cols
				e.write("[5.25, 7.25, 1]\n")	# dims
				e.write("envelopes.pdf")		# out_file

if __name__ == "__main__":
	
	import os
	import sys
	from hashlib import md5

	
	dconf = "conf/env.conf"
	valid_hash = "b8e54d1dfd6816e139160478553ae209"	#	md5 of default env.conf
	config = {"path":'', "sheet":0, "rows":[], "cols":[], "dims":[], "out_file":''}
	conf = ''
			
	#	Check whether any env.conf exists.
	#	If so: hash for later comparison.
	if os.path.exists(dconf):
		with open(dconf,"r") as def_conf:
			contents = def_conf.readlines()
			cont_enc = ''.join(contents).encode("utf-8")
			def_hash = md5(cont_enc).hexdigest()
	
	#	Figure out which, if any, of the arguments are a .conf file.
	for arg in sys.argv:
		if ".conf" in arg:						
			conf = "conf/" + arg
			
	#	If no .conf file exists or env.conf is corrupted then remake it
	def_hash = valid_hash
	if conf == '':
		if not os.path.exists(dconf) or def_hash != valid_hash:
			print("env.conf does not exist or is invalid, creating...")
			createDefaultConf()
		conf = dconf
	
	#	Unpack configuration file into dict
	config = unpack(conf)
	
	#	Grab optional arguments from command line
	for arg in sys.argv:
		if ".xlsx" in arg or ".xls" in arg:	config["path"] = "in/" + arg
		if "pth=" in arg: 	config["path"] = "conf/" + arg[4:]
		if "sht=" in arg: 	config["sheet"] = int(arg[4:])
		if "row=" in arg:	config["rows"] = [int(i) for i in arg[5:-1].split(',')]
		if "col=" in arg:	config["cols"] = arg[5:-1].split(',')
		if "dim=" in arg:	config["dims"] = [int(i) for i in arg[5:-1].split(',')]
	
	#	Actually make the envelopes
	env = Envelopes(config["path"], config["sheet"], config["rows"], config["cols"], config["dims"])
	
	#	Output files if necessary
	if len(sys.argv) == 1:	sys.argv.append("-out")
	
	for arg in sys.argv:
		if ".tex" in arg:	env.generate_TeX("out/"+arg)			# generates just .tex
		if ".pdf" in arg:	env.generate_PDF("out/"+arg)			# generates .tex and .pdf
		if arg == "-out":	print(env.excel_to_TeX())		# prints .tex source
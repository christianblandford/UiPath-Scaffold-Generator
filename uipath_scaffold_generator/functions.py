import urllib, os, shutil, zipfile, json, urllib.request
from shutil import copyfile
from openpyxl import Workbook, load_workbook
from .sequence import Sequence_To_Scaffold

class Functions:
	def __init__(cls):
		pass

	#strips special chars from text to be used when creating a filename
	def make_file_name(text) :
		return ''.join(e for e in text if e.isalnum())

	#Converts to Title Case
	def make_title_case(text) :
		return text.title()

	#Creates a string that is suitable for a UiPath Project Name
	@classmethod
	def make_project_name(cls, text):
		return cls.make_file_name(cls.make_title_case(text))

	#Create a directory and any parent directories needed
	def create_dir(path):
		os.makedirs(path, 777, True)
		return path

	#Downloads a file from a web location (url) and saves it to dst + filename
	def download_file(url, dst, name):
		#determine filetype of download
		path = urllib.parse.urlparse(url).path
		ext = os.path.splitext(path)[1]
		#Create full filename for download
		file_name = name + ext
		file_path = os.path.join(dst, file_name)

		# Download the file from `url` and save it locally under `file_name`:
		with urllib.request.urlopen(url) as response, open(file_path, 'wb') as out_file:
			shutil.copyfileobj(response, out_file)
			return file_path

	#Unzips a file and returns an array of the files created
	def unzip_file(src, dst):
		z_file = zipfile.ZipFile(src, 'r')

		#get the name of the output dir from upzip
		output = os.path.join(os.path.dirname(src), z_file.namelist()[0])
		#unzip
		z_file.extractall(os.path.dirname(src))
		z_file.close()
		#Delete the zip
		os.remove(src)

		return output

	#Creates a new file if one does not exist. Otherwise it will update the file with "data" if it is passed in
	def create_file(dst, data=None):
		file = open(dst,"w+")
		if not data is None: file.write(data)
		file.close()
		return dst

	#Reads contents of file
	def read_file(file):
		infile = open(file, "r")
		data = infile.read()
		infile.close()
		return data

	#Copies a file from src to dst
	def copy_file(src, dst):
		copyfile(src, dst)
		return dst

	#Moves a file from src to dst, recursively
	def move_file(src, dst):
		os.move(src, dst)
		return dst

	#Deletes a file
	def delete_file(file):
		os.remove(file)

	#Updates a file with the data passed in
	@classmethod
	def save_file(cls, file, data):
		cls.create_file(file, data) # Create file with new data
		return file

	#Deletes a directory. Will throw error if dir is not empty.
	def delete_dir(dir):
		os.rmdir(dir)

	#Renames a single file
	def rename_file(src, dst):
		os.raname(src, dst)
		return dst

	#Renames a directory
	def rename_dir(src, dst):
		for retry in range(10):
			try:
				os.rename(src,dst)
				break
			except:
				if retry < 9: print('rename failed, retrying...')
				else: raise
		return dst

	#reads an excel file sheet and returns it
	def read_excel_sheet(file, sheet):
		wb = load_workbook(file, read_only=True) # Load the worksheet as read only
		ws = wb[sheet] # Grab the proper worksheet
		return ws

	#Gets a list of sequences to scaffold by reading excel file
	@classmethod
	def get_sequences_to_create(cls, file):
		ws = cls.read_excel_sheet(file, "Sequences")

		#Build an array of sequences to scaffold out
		sequences_to_scaffold = []

		#Loop through cells in the worksheet to build the array
		for i,row in enumerate(ws.rows):
			#check if this is the first run of the loop, ignore the data if it is. Due to the excel headers.
			if i > 0:
				name = row[0].value
				location = row[1].value
				parent = row[2].value
				description = row[3].value
				sequences_to_scaffold.append(Sequence_To_Scaffold(name, location, parent, description))

		return sequences_to_scaffold

	#Returns contents of JSON file
	def read_json(file):
		json_file = open(file, "r") # Open the JSON file for reading
		data = json.loads(json_file.read()) # Read the JSON into the buffer
		json_file.close() # Close the JSON file
		return data

	#Opens JSON file, replaces contents with "data" variable
	def update_json(file, data):
		json_file = open(file, "w+")
		json_file.write(json.dumps(data, indent=4, sort_keys=True)) # PrettyPrint and write the JSON
		json_file.close()
		return file

	#Updates a single value in JSON file
	@classmethod
	def update_json_file_value(cls,file, key, new_val):
		json_data = read_json(file) # Read the file
		json_data[key] = new_val # Update value at key
		cls.update_json(file, data) # Save the JSON
		return file


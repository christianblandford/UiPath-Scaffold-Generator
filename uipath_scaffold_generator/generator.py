# A uipath-scaffold generator class
import urllib, os, shutil, zipfile, json, tempfile, xml.etree.ElementTree, urllib.request, console_functions as terminal
from openpyxl import Workbook, load_workbook
from bs4 import BeautifulSoup
from pathlib import Path
from shutil import copyfile

import uipath as UiPath


class Functions:

	#strips special chars from text to be used when creating a filename
	@classmethod
	def make_file_name (cls, text) :
		return ''.join(e for e in text if e.isalnum())

	#Converts to Title Case
	@classmethod
	def make_title_case(cls, text) :
		return text.title()

	#Creates a string that is suitable for a UiPath Project Name
	@classmethod
	def make_project_name(cls, text):
		return cls.make_file_name(cls.make_title_case(text))

	#Create a directory and any parent directories needed
	@classmethod
	def create_dir(cls, path):
		os.makedirs(path, 777, True)
		return path

	#Downloads a file from a web location (url) and saves it to dst + filename
	@classmethod
	def download_file(cls, url, dst, name):
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
	@classmethod
	def unzip_file(cls, src, dst):
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
	@classmethod
	def create_file(cls, dst, data=None):
		file = open(dst,"w+")
		if not data is None: file.write(data)
		file.close()
		return dst

	#Reads contents of file
	@classmethod
	def read_file(cls, file):
		infile = open(file, "r")
		data = infile.read()
		infile.close()
		return data

	#Copies a file from src to dst
	@classmethod
	def copy_file(cls, src, dst):
		copyfile(src, dst)
		return dst

	#Moves a file from src to dst, recursively
	@classmethod
	def move_file(cls, src, dst):
		os.move(src, dst)
		return dst

	#Deletes a file
	@classmethod
	def delete_file(cls, file):
		os.remove(file)

	#Updates a file with the data passed in
	@classmethod
	def save_file(cls, file, data):
		cls.create_file(file, data) # Create file with new data
		return file

	#Deletes a directory. Will throw error if dir is not empty.
	@classmethod
	def delete_dir(cls, dir):
		os.rmdir(dir)

	#Renames a single file
	@classmethod
	def rename_file(cls, src, dst):
		os.raname(src, dst)
		return dst

	#Renames a directory
	@classmethod
	def rename_dir(cls, src, dst):
		os.rename(src, dst)
		return dst

	#reads an excel file sheet and returns it
	@classmethod
	def read_excel_sheet(cls, file, sheet):
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
	@classmethod
	def read_json(cls, file):
		json_file = open(file, "r") # Open the JSON file for reading
		data = json.loads(json_file.read()) # Read the JSON into the buffer
		json_file.close() # Close the JSON file
		return data

	#Opens JSON file, replaces contents with "data" variable
	@classmethod
	def update_json(cls, file, data):
		json_file = open(file, "w+")
		json_file.write(json.dumps(data, indent=4, sort_keys=True)) # PrettyPrint and write the JSON
		json_file.close()
		return file

	#Updates a single value in JSON file
	@classmethod
	def update_json_file_value(cls, file, key, new_val):
		json_data = read_json(file) # Read the file
		json_data[key] = new_val # Update value at key
		cls.update_json(file, data) # Save the JSON
		return file


class Generator:
	def __init__(self, name, description, zip_url=None, default_sequence=None, scaffold_type=None):
		self.name = name
		self.description = description
		self.zip_url = zip_url
		self.default_sequence = default_sequence

	
	#Scaffold a project
	def scaffold_project(self):
		#Create new project
		self.project = Project()
		self.project.framework = self.name
		terminal.variable("Using: ", self.project.framework)
		self.project.name = Functions.make_project_name(terminal.input("Enter a name for the project:"))
		terminal.variable("Project name set to: ", self.project.name)
		self.project.description = terminal.input("Enter a description for the project:", self.project.name + " UiPath project.")
		terminal.variable("Project description set to: ", self.project.description)
		self.project.path = os.path.join(self.get_working_dir(), self.project.name)

		#Attempt to get the parent of the path input by the user
		parent_dir = Path(self.project.path).parent

		#Check if parent dir exists. Error and exit if it does not
		if not os.path.isdir(parent_dir) :
			terminal.warn("Parent directory does not exist.")
			create_dirs = terminal.input("Would you like to create the directories now?", "Y")
			if create_dirs != "Y":
				console.error("Parent directory does not exist. Please double check and try again.")
				exit(1)

		#Create directories
		terminal.variable("Creating parent directory: ", parent_dir)
		Functions.create_dir(parent_dir)

		#Output download progress
		terminal.variable("Downloading framework from: ", self.zip_url)

		#Download and unzip files
		zip_location = Functions.download_file(self.zip_url, parent_dir, self.project.name)
		expanded_zip_path = Functions.unzip_file(zip_location, self.project.path)
		Functions.rename_dir(expanded_zip_path, self.project.path)

		#Call project.ready() now that it has been created
		self.project.ready()

		#Update project.json
		self.project.edit_json_value("name", self.project.name)
		self.project.edit_json_value("description", self.project.description)

		#Copy SequencesToScaffold.xlsx to the project directory
		Functions.copy_file("SequencesToScaffold.xlsx", os.path.join(self.project.path, "SequencesToScaffold.xlsx"))

		terminal.header(["Project created successfully.", "Location: " + self.project.path], "white", "white", "*")
	
	#Prompts the user for directory to create project in
	def get_working_dir(self):
		user_path = os.getcwd()

		#Ask user if they want to use the current working path
		scaffold_here = terminal.input("Your current directory is: " + user_path + " Is this where you would like to scaffold your project?", "Y")

		#If user chooses to use a different path
		if not scaffold_here == "Y":
			user_path = terminal.input("Okay please enter the path to the directory you want to use: ")

		return user_path

	#Scaffolds files
	def scaffold_seqeuences(self):
		terminal.input("I have created a SequencesToScaffold.xlsx for you in your project directory. Please add the sequences to scaffold to this file. When you are done, hit any key.", allow_empty=True)

		excel_file_location = os.path.join(self.project.path, "SequencesToScaffold.xlsx") 
		sequences = Functions.get_sequences_to_create(excel_file_location) #get the list of all sequences we want to scaffold out. This comes from the excel file.

		self.files_created = []

		for item in sequences:
			#Create parent dirs
			Functions.create_dir(os.path.join(self.project.path, item.location))
			
			item.path = os.path.join(self.project.path, item.location, item.name + ".xaml")
			Functions.copy_file(os.path.join(self.project.path, self.default_sequence), item.path) # Copy the default sequence over to the specified location
			self.files_created.append(item)

		terminal.header("Files created successfully.", "white", "white", "*")

class Project():
	def __init__(self):
		self.name = ""
		self.description = ""
		self.path = ""
		self.framework = ""

	#run this when done download,extracting,renaming files
	def ready(self):
		self.project_json_location = str(os.path.join(self.path, "project.json"))
		self.json = Functions.read_json(self.project_json_location)

	#Edits a single json value
	def edit_json_value(self, key, value):
		self.json[key] = value
		Functions.update_json(self.project_json_location, self.json)
		return self.json

class Sequence_To_Scaffold():
	def __init__(self, name, location, parent, description):
		self.name = name
		self.location = location
		self.parent = parent
		self.description = description
		self.path = ""
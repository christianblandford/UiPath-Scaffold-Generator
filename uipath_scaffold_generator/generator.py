# A uipath-scaffold generator class
import os, console_functions as console
from pathlib import Path
from .functions import * #import generator_functions.py
from .project import * #import uipath_project.py
from .sequence import *

class Generator:
	def __init__(self, name, description, zip_url=None, default_sequence=None, scaffold_type=None):
		self.name = name
		self.description = description
		self.zip_url = zip_url
		self.default_sequence = default_sequence

		self.self_path = os.path.dirname(os.path.abspath(__file__)) # Path to this package, so we can copy the excel file over

	
	#Scaffold a project
	def scaffold_project(self):
		#Create new project
		self.project = Project()
		self.project.framework = self.name
		console.variable("Using: ", self.project.framework)
		self.project.name = Functions.make_project_name(console.input("Enter a name for the project:"))
		console.variable("Project name set to: ", self.project.name)
		self.project.description = console.input("Enter a description for the project:", self.project.name + " UiPath project.")
		console.variable("Project description set to: ", self.project.description)
		self.project.path = os.path.join(self.get_working_dir(), self.project.name)

		#Attempt to get the parent of the path input by the user
		parent_dir = Path(self.project.path).parent

		#Check if parent dir exists. Error and exit if it does not
		if not os.path.isdir(parent_dir) :
			console.warn("Parent directory does not exist.")
			create_dirs = console.input("Would you like to create the directories now?", "Y")
			if create_dirs != "Y":
				console.error("Parent directory does not exist. Please double check and try again.")
				exit(1)

		#Create directories
		console.variable("Creating parent directory: ", parent_dir)
		Functions.create_dir(parent_dir)

		#Output download progress
		console.variable("Downloading framework from: ", self.zip_url)

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
		Functions.copy_file(os.path.join(self.self_path, "SequencesToScaffold.xlsx"), os.path.join(self.project.path, "SequencesToScaffold.xlsx"))

		console.header(["Project created successfully.", "Location: " + self.project.path], "white", "white", "*")
	
	#Prompts the user for directory to create project in
	def get_working_dir(self):
		user_path = os.getcwd()

		#Ask user if they want to use the current working path
		scaffold_here = console.input("Your current directory is: " + user_path + " Is this where you would like to scaffold your project?", "Y")

		#If user chooses to use a different path
		if not scaffold_here == "Y":
			user_path = console.input("Okay please enter the path to the directory you want to use: ")

		return user_path

	#Scaffolds files
	def scaffold_seqeuences(self):
		console.input("I have created a SequencesToScaffold.xlsx for you in your project directory. Please add the sequences to scaffold to this file. When you are done, hit any key.", allow_empty=True)

		excel_file_location = os.path.join(self.project.path, "SequencesToScaffold.xlsx") 
		sequences = Functions.get_sequences_to_create(excel_file_location) #get the list of all sequences we want to scaffold out. This comes from the excel file.

		self.files_created = []

		for item in sequences:
			#Create parent dirs
			Functions.create_dir(os.path.join(self.project.path, item.location))
			
			item.path = os.path.join(self.project.path, item.location, item.name + ".xaml")
			Functions.copy_file(os.path.join(self.project.path, self.default_sequence), item.path) # Copy the default sequence over to the specified location
			self.files_created.append(item)

		console.header("Files created successfully.", "white", "white", "*")
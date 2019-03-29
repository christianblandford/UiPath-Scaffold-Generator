import os
from .functions import *

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
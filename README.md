# Table of Contents
- [What is it](#what_is_it)
- [Main Features](#main-features)  
- [How to run](#how_to_run)
- [Project Structure](#project_structure)
- [Where to use the project](#where-to-use)   
- [How to Run this example](#how-to-run)
- [Discussion and Development](#dev)

# Word-Add-In-QuickStart: a simple Microsoft Word Add-in example


# What is it? <a name="what_is_it"></a>

This repo is an illustration of a Microsoft Word Add-in developed with reference to the default API and the  

# Main Features <a name="main-features"></a>

- PerfFlow.js is built on Tensorflow.js, with D3.js and Leaflet.js among other packages. 
- It provides neat-looking correlograms to compare the correlation of the features that it captures.  
- It provides functions to train Linear and Non-Linear Regression neural nets to determine the optimum load a device can and should be able to handle.
- It is built for [*Leaflet.js*](https://github.com/Leaflet/Leaflet) and Uses Leaflet EasyButton to allow you to create your own custom buttons to extend functionality.

# Project Structure <a name="project_structure"></a>
    Home.html     # main page of the client
    Home.js       # where the core logic of the applications lies, from map initalization to custom button creation.
	Home.css      # made to make the sample app look acceptable.
    Web.config    # Configuration file for the solution
    /Scripts	  # included jQuery, FabricUI and Office libraries.
    /Images       # where images for the client are stored
    /Content      # CSS resources related to FabricUI
	/Functions    # folder for future dev of other JS files and custom panel views/logic

# Where to use it <a name="where-to-use"></a>
This project, or parts of it can be helpful samples for you if you are looking to do any the following:
1. Develop a Microsoft Word add-in
	- Capture text in a word document text selection
	- Highlight text in a word document text selection
	- Use the Word JS API to start working with Add-In development.

# How to run <a name="how-to-run"></a>
1. Clone the repository
2. Navigate to the word-add-in-quickstart directory
3. Open the solution in Visual Studio (2017 and after for best compatibility)
4. Perform a `build->clean solution` and a  `build->build solution` of the project.
6. The sample Word application can now be started with a `debug->start debugging` command.

# Discussion and Development <a name="dev"></a>
All development and discussion was done by Sanket and Smriti Mehrotra. This Project was developed as an addon tool to aid teams working on common business-wide
documentation to lookup definitions of common business terms from a custom prepared Data Source(e.g. an Excel Sheet or Sharepoint File)

#### Licence: [MIT](https://github.com/opensource9ja/danfojs/blob/master/LICENCE)

#### Created by [Smriti Mehrotra](https://github.com/SSMehrotra)

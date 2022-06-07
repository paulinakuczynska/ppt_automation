# XML customization for PowerPoint
Python scripts run from the command line for automating tasks related to PowerPoint
## Table of contents
* [General info](#general-info)
* [Technologies](#technologies)
* [Setup](#setup)
* [References](#references)
## General info
These scripts enable efficient modification of xml code in .pptx files. Due to not all PowerPoint options are available from the GUI level, adjusting presentation templates to your needs often requires the preparation of custom xml parts. These tasks tend to be time-consuming and monotonous, so their automation is important for people who design presentation templates.
#### Custom colors
PowerPoint allows to set only 10 colors constituting the main palette named Theme Colors. By adding a custom xml part, the palette can be increased by another 51 colors. The script allows to set the values and names of the colors.
#### Custom margins
By default, PowerPoint sets frames around text inside placeholders, which makes it difficult to align text and graphics elements. Although it is possible to change the margins manually, it is necessary to do this on each placeholder separately. Additionally, copying placeholders to the slide restores the default value. Changing the default value is possible by modifying the xml code. The script allows to set the left, right, bottom and top margins to zero, and the change will be applied to all layouts.
## Technologies
* Python 3.8.10
* The Python Standard Library only
## Setup
You need to [install Python](https://www.python.org/downloads/) to run the scripts. Paths used in scripts are readable both in Linux and Windows.
Scripts have a shebang line defined, so can be run as ```./<filename>.py```. The presentation file for customization should be named "todo.pptx" and located on the desktop.
* The "Custom colors" script requires two parameters, a color and a name, each can contain from one to 51 values. Please, use ```./pptAddCustomColors.py -h``` to see the usage. Example: ```./pptAddCustomColors.py -v 000000 ffffff -n black white```.
* The "Custom margins" script doesn't require any parameter. Please, use ```./pptSetMargins.py```.
## References
* [OOXML Hacking](https://www.brandwares.com/bestpractices/category/xml-hacks/) by John Korchok
* Al Sweigart. Automate the Boring Stuff with Python, 2nd Edition: Practical Programming for Total Beginners, 2021.
* Preserve namespaces when parsing xml in [Stack Overflow](https://stackoverflow.com/questions/54439309/how-to-preserve-namespaces-when-parsing-xml-via-elementtree-in-python)

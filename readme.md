This is the guide to the folder tax writen in july 2026.
Result.py - main code to start the taxes. 




Cnb_converter.py just serves as a function in result.py to convert currencies (either major or minor currencies). Its indended only for result.py not independently run

merge.py - merge all xml in the "xml" folder. If gaps between dates of the xmls it will alert. The merged xml is called "merged.xml"

czk.py - "merge.xml" -> "final.xml" - adds czk values into the xml


# plantuml-word-add-in-vsto

PlantUML Gizmo (add-in) for Microsoft Word

Software to render PlantUML diagrams within Word. This uses the PlantUML server in the cloud to generate the images.

Developed as a VSTO add-in for Word. It installs with a setup.exe

## Installation / Updates

The solution now is a VSTO (Visual Studio Tools for Office) Document-level customization (in Word). 
It requires .NET 4.0 as well as some VSTO 2010 files to run.

To install this on your Windows machine, follow these steps: 

 * In Windows, add http://profs.etsmtl.ca/ to the "Trusted Sites" zone (do not require https:). 
   For more information, see how to add sites to zones. 
 * In Internet Explorer (I'm not sure it works with other browsers), download and run [setup.exe](http://profs.etsmtl.ca/cfuhrman/PlantUMLGizmoWord/setup.exe).
 * You should see a warning that the Publisher cannot be verified. This is normal because currently there is no certificate for this tool. However, you can verify that it's indeed<br/> 
From: `http://profs.etsmtl.ca/cfuhrman/PlantUMLGizmoWord/PlantUML Gizmo Word Add-in.vsto`

If you don't trust this tool, then do not install it.

Confidentiality warning: Since PlantUML.com creates the images from the text descriptions you specify, any UML drawings done in your documents could be exposed to the cloud.

![](http://www.plantuml.com:80/plantuml/png/hSw_IWH130RmVfuYNEM1kmzWSOq78dXWeTXS4sV3CDnECcHohFhqPbPFiB2pcJ_VzoEadJ9NL2pmYl6KLCuytSei2gR8pIjY2-r7DNkVoK_DqSvb3fvQZhaY6snUH2SGFVATI7AsbzWsW0sMt-vnzklvyE0mVnHPIVgBZ57AAcf8mwx23IHYKZJQPqo-qv6lZDvi6-emx9BtkM2YZfV-uKPgzprok5RNyEN7SRoeHFRasKLim_8zgykE-fkclAf_tkCJ.png)

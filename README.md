# d365EmailTemplate

A d365 CE package that would allow users to create visually rich email templates.

## Installation

1. Create 3 web resources
    * prefix_email.ribbon.js - js file that will launch the html web resource from a ribbon button
    * prefix_d365EmailTemplate.js - the main js file of the application
    * prefix_d365EmailTemplate.html - the main html file that hosts prefix_d365EmailTemplate.js
    
2. Modify these files to use your **prefix_**
    * prefix_d365EmailTemplate.pug
      - script(src="prefix_d365EmailTemplate.js" type="text/javascript")
    * prefix_email.ribbon.js
      - const wrName = "prefix_d365EmailTemplate.html"
      
3. Use the Ribbon Workbench to customize the Insert Template Button Command
    * Delete the default action
    * Add a new Custom JS action
      - Library: prefix_email.ribbon.js
      - Function Name: openD365EmailTemplate
      - Add Crm Parameter = PrimaryControl

    
  
  

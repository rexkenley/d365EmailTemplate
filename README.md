# d365EmailTemplate

A d365 CE package that would allow users to create visually rich email templates.

## Installation

1. Install [NodeJS](https://nodejs.org/en/)

2. run `npm i`

3. Modify these files to use your **prefix_**
    * prefix_d365EmailTemplate.pug
      - script(src="prefix_d365EmailTemplate.js" type="text/javascript")
    * prefix_email.ribbon.js
      - const wrName = "prefix_d365EmailTemplate.html"

4. run `npm run build:dev` for development build

5. run `npm run build:d365EmailTemplate:prd` and `npm run build:email:prd` for production build

> the files will be in the **dist** folder

6. Create 3 web resources
    * prefix_email.ribbon.js - js file that will launch the html web resource from a ribbon button. This will also pass the regarding object if available.
    * prefix_d365EmailTemplate.js - the main js file of the application
    * prefix_d365EmailTemplate.html - the main html file that hosts prefix_d365EmailTemplate.js
      
7. Use the Ribbon Workbench to customize the Insert Template Button Command
    * Delete the default action
    * Add a new Custom JS action
      - Library: prefix_email.ribbon.js
      - Function Name: openD365EmailTemplate
      - Add Crm Parameter = PrimaryControl

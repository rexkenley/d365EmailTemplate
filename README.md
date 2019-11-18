# d365EmailTemplate (v9)

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

## Buttons

- **Entity** - this lists all the entities that can be merge with templates. The check mark will indicate the current selected entity.
- **Template** - this lists all the templates that are available for the selected entity. The check mark will indicate the current selected template. Clicking on a template will load it in the tinyMCE viewer below.
- **Attribute** - this lists all the available attributes of the selected entity. When you click on an attribute it will add the attribute to the selected template. 
- **Format** - if you select a DateTime attribute, this button will appear to provie additional formatting options.
- **Save**
- **Merge** - this will merge the selected template and place the resulting html to the "parent" email window.
- **Preview** - this will show the result the merge.

## Addtional Documentation
- [tinyMCE](https://www.tiny.cloud/docs/) - the editor used in this project
- [Handlebars](https://handlebarsjs.com/) - the templating engine used in this project
- [Handlebars Intl](https://formatjs.io/handlebars/) - Handlebars "helper" that formats the date

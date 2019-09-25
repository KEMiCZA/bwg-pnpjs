# Creating a Provisioning Solution Using Site Designs and PnPjs

Contains code for the demo given at BIWUG 2019.

In this solution a site design is used to provision the site structure. This site design contains several actions (feel free to extend this):
 * create a list
 * install a custom *Hello Biwug 2019* SPFx solution.
 * flowtrigger

The flowtrigger is our glue between the site design- and the back-end- provisioning in an Azure Function. This is allows custom code to do pretty much anything with our site after it has been created. In this case we will manipulate the homepage and add a webpart with properties.

## Solution structure
* **slides**: contains the slides that were presented at BIWUG 2019
* **spfx-pnpjs-biwug**: SPFx solution that contains two webparts that were used during the demo to showcase PnPjs & Managing Site Designs/Scripts using PnPjs
* **spfx-hello-biwug**: SPFx solution that contains a simple 'Hello Biwug' webpart and a [Feature Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/provision-sp-assets-from-package) asset to provisiong a simple list
* **azure-function**: Contains the Azure Function in nodejs that will manipulate the homepage
* **ms-flow**: Exported Microsoft Flow to trigger the Azure Function
* **site-scripts**: several site scripts used in this solution

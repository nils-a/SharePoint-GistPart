# GistPart - A WebPart to show GitHub gists in SharePoint modern pages

This is where you include your WebPart documentation.

### Building the code

#### get & prepare

    bash
    git clone <the repo>
    cd src
    npm i
    npm i -g gulp
    gulp

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* temp/* - some temp files
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.


#### Run local server

    gulp serve  

#### Ship

##### Ship js/css/whatever to your CDN

	gulp --ship

copy ./temp/dist/* to your CDN

##### Ship *.sppkg to SharePoint

	gulp package-solution --ship
	
copy ./sharepoint/solution/*.sppkg to your SharePoints app-catalog
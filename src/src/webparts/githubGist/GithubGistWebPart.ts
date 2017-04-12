import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
//import { escape } from '@microsoft/sp-lodash-subset';
//import { SPComponentLoader } from '@microsoft/sp-loader';

import styles from './GithubGist.module.scss';
import * as strings from 'githubGistStrings';
import { IGithubGistWebPartProps } from './IGithubGistWebPartProps';
import * as superagent from 'superagent';
//import { jsonp } from 'superagent-jsonp';
declare var superagentJSONP: any; import 'superagent-jsonp'; var jsonp = superagentJSONP;// remove and switch with the line above, when https://github.com/lamp/superagent-jsonp/issues/25 is fixed

export default class GithubGistWebPart extends BaseClientSideWebPart<IGithubGistWebPartProps> {

  public render(): void {
    var gistContent = `<div class="${styles.warning}">${strings.WarningGistPropertyNotSet}</div>`;
    var title = ``;
    if (!!this.properties.gistId) {
      gistContent = `<a class="${styles.loading}" href="">Loading gist-id:${this.properties.gistId} ...</a>`;
    }
    if (!!this.properties.title) {
      title = `<div class="${styles.title}">${this.properties.title}</div>`;
    }

    this.domElement.innerHTML = `
      <div class="${styles.gist}">
        ${title}
        <div class="gist-content">
          ${gistContent}
        </div>
      </div>`;

    if (!!this.properties.gistId) {
      // now load the stuff...
      superagent.get(`https://gist.github.com/${this.properties.gistId}.json`).use(jsonp).end((err, res) => {
        var content = this.domElement.querySelector(".gist-content");
        if (err) {
          content.classList.add(`${styles.error}`);
          content.innerHTML = `Sadly there was an error. (${err})`; 
        } else {
          content.classList.remove(`${styles.error}`);
          content.innerHTML = res.body.div;
          if(!!this.properties.gistFile) {
  /*
  * TODO: What if multiple WebParts all point to the same gist, sowing only different files?
  * we'd go fetch the full gist every time... this is not so good..
  */
            var files = content.querySelectorAll(".gist-file");
            var toKeep = res.body.files.indexOf(this.properties.gistFile);
            if(toKeep > -1) {
              for(var i=0; i<files.length; i+=1){
                if(i==toKeep) continue;
                files[i].remove();
              }
            }

          }
          if(!document.head.querySelector("link[href='"+res.body.stylesheet+"']")){
            var link = document.createElement("link");
            link.setAttribute("type", "text/css");
            link.setAttribute("rel", "stylesheet");
            link.setAttribute("href", res.body.stylesheet);
            document.getElementsByTagName("head")[0].appendChild(link);
          }
        }
      });
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('gistId', {
                  label: strings.GistIdFieldLabel
                }),
                PropertyPaneTextField('gistFile', {
                  label: strings.GistFileFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

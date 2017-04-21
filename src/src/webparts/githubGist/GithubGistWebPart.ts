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
  private readonly gistEventName = "gistload";
  private gistContent = "<div>Initializing...</div>";

  public extractGist(store: Element) {
    var content = this.domElement.querySelector(".gist-content");
    var err = store.querySelector("div.error");
    if (!!err) {
      content.classList.add(`${styles.error}`);
      content.innerHTML = `Sadly there was an error. (${err.innerHTML})`;
      return;
    }

    content.classList.remove(`${styles.error}`);
    content.innerHTML = store.innerHTML;
    if (!!this.properties.gistFile) {
      var attribQuery = "[data-file-name=\"" + this.properties.gistFile + "\"]";
      var files = content.querySelectorAll(".gist-file:not("+ attribQuery+")");
      for (var i = 0; i < files.length; i += 1) {
        files[i].remove();
      }
    }
  }

  public handleGistLoaded(e: CustomEvent): void {
    if (e.detail !== this.properties.gistId) {
      return;
    }
    var domGistId = "gsitpart-loaded-" + this.properties.gistId;
    var gistStore = document.querySelector("[data-gist-id=\"" + domGistId + "\"]");
    if (!gistStore) {
      //impossible...
      throw "Gist-Store not found!";
    }
    if (gistStore.getAttribute("data-load-state") !== "loaded") {
      //also impossible
      throw ("Gist-Store-loaded event was fired, but store reported " + gistStore.getAttribute("data-load-state"));
    }
    this.extractGist(gistStore);
  }

  public render(): void {
    this.gistContent = `<div class="${styles.warning}">${strings.WarningGistPropertyNotSet}</div>`;
    var title = ``;
    if (!!this.properties.gistId) {
      this.gistContent = `<a class="${styles.loading}" href="">Loading gist-id:${this.properties.gistId} ...</a>`;
    }
    if (!!this.properties.title) {
      title = `<div class="${styles.title}">${this.properties.title}</div>`;
    }

    this.domElement.innerHTML = `
      <div class="${styles.gist}">
        ${title}
        <div class="gist-content">
          ${this.gistContent}
        </div>
      </div>`;


    if (!this.properties.gistId) {
      return;
    }

    var domGistId = "gsitpart-loaded-" + this.properties.gistId;
    var gistStore = document.querySelector("[data-gist-id=\"" + domGistId + "\"]");
    if (!!gistStore) {
      if (gistStore.getAttribute("data-load-state") === "loaded") {
        // TODO: copy the stuff directly here and we're done!
        this.extractGist(gistStore);
        return;
      }
      gistStore.addEventListener(this.gistEventName, (e:CustomEvent) => {this.handleGistLoaded(e)}); //use () => {} to avoid "this" re-scoping
      return;
    } else {
      // add a nice store for this gist then load it
      var storeContainer = document.querySelector("[data-gist-id=\"storecontainer\"]");
      if (!storeContainer) {
        storeContainer = document.createElement("div");
        storeContainer.setAttribute("data-gist-id", "storecontainer");
        storeContainer.setAttribute("style", "display:none;");
        document.body.appendChild(storeContainer);
      }

      gistStore = document.createElement("div");
      gistStore.setAttribute("data-gist-id", domGistId);
      gistStore.setAttribute("data-load-state", "loading");
      storeContainer.appendChild(gistStore);
      gistStore.addEventListener(this.gistEventName,(e:CustomEvent) => {this.handleGistLoaded(e)}); //use () => {} to avoid "this" re-scoping
      superagent.get(`https://gist.github.com/${this.properties.gistId}.json`).use(jsonp).end((err, res) => {
        if (!!err) {
          var errContainer = document.createElement("div");
          errContainer.classList.add("error");
          errContainer.innerHTML = `Sadly there was an error. (${err})`;
          gistStore.appendChild(errContainer);
        } else {
          // add content
          gistStore.innerHTML = res.body.div;
          // mark files, so they are easy to find
          var fileDefs = res.body.files;
          var files = gistStore.querySelectorAll(".gist-file");
          for (var i = 0; i < fileDefs.length; i += 1) {
            files[i].setAttribute("data-file-name", fileDefs[i]);
          }
          // add css
          if (!document.head.querySelector("link[href='" + res.body.stylesheet + "']")) {
            var link = document.createElement("link");
            link.setAttribute("type", "text/css");
            link.setAttribute("rel", "stylesheet");
            link.setAttribute("href", res.body.stylesheet);
            document.getElementsByTagName("head")[0].appendChild(link);
          }
          // notify all waiting listeners
          gistStore.setAttribute("data-load-state", "loaded");
          gistStore.dispatchEvent(new CustomEvent(this.gistEventName, { bubbles: true, cancelable: false, detail: this.properties.gistId }));
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

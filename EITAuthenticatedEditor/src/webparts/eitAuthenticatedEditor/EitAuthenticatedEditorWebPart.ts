import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { AadTokenProvider, HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import * as microsoftTeams from '@microsoft/teams-js';

import styles from './EitAuthenticatedEditorWebPart.module.scss';
import * as strings from 'EitAuthenticatedEditorWebPartStrings';

import EitAuthenticatedEditorWebPart from './components/EitAuthenticatedEditorWebPart';
import { IEitAuthenticatedEditorWebPartProps } from './components/IEitAuthenticatedEditorWebPartProps';

export interface IEitAuthenticatedEditorWebPartWebpartProps {
  WebPartTitle: string;
  AadTokenProvider: AadTokenProvider;
  TemplateUrl: string;
  TemplateHtml: string;
  Resources: string;
  BearerToken: string;
}

export default class EitAuthenticatedEditorWebPartWebPart extends BaseClientSideWebPart<IEitAuthenticatedEditorWebPartProps> {

  private _teamsContext: microsoftTeams.Context;
  private _unqiueId;
  public _scriptEditorPropertyPane;
  public _disposeScriptEditorPropertyPane;
  public _bearerToken;
  public _currentUser;

  constructor() {
    super();
    this.scriptUpdate = this.scriptUpdate.bind(this);
    this.disposeScriptUpdate = this.disposeScriptUpdate.bind(this);
    this.executeScript = this.executeScript.bind(this);
    this.evalScript = this.evalScript.bind(this);
    this.cleanUp = this.cleanUp.bind(this);
  }


  public scriptUpdate(_property: string, _oldVal: string, newVal: string) {
    this.properties.script = newVal;
    this._scriptEditorPropertyPane.initialValue = newVal;
  }

  public disposeScriptUpdate(_property: string, _oldVal: string, newVal: string) {
    this.properties.disposeScript = newVal;
    this._disposeScriptEditorPropertyPane.initialValue = newVal;
  }

  protected onInit(): Promise<any> {
    let retVal: Promise<any> = Promise.resolve();
    if (this.context.microsoftTeams) {
      retVal = new Promise((resolve, reject) => {
        this.context.microsoftTeams.getContext(context => {
          this._teamsContext = context;
          resolve();
        });
      });
    }
    return retVal;
  }

  public async render(): Promise<void> {

    this._unqiueId = this.context.instanceId;
    this._currentUser = this.context.pageContext.user;

    if (!(Environment.type === EnvironmentType.Local)) {
      this.properties.AadTokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
    }

    ReactDom.unmountComponentAtNode(this.domElement);

    var output;
    // Create the webpart container
    output = `<div class="${styles.eitAuthenticatedEditor}">`;
    // Set the title, if there is one provided
    output += this.properties.WebPartTitle ? `<div class="${styles.webpartTitle}">${this.properties.WebPartTitle.replace("{SiteCollectionTitle}", this.GetCurrentWebTitle())}</div>` : '';
    if (this.properties.TemplateUrl) {
      this._bearerToken = await this.properties.AadTokenProvider.getToken(this.properties.Audience);
      // Set the HTML template
      this.properties.TemplateHtml = await this.GetTemplateHtml();
      // Replace the User token in the HTML
      this.properties.TemplateHtml = this.replaceTokens("CurrentUser_displayName", this._currentUser.displayName, this.properties.TemplateHtml);
      this.properties.TemplateHtml = this.replaceTokens("CurrentUser_email", this._currentUser.email, this.properties.TemplateHtml);
      // Replace the Bearer token in the HTML
      this.properties.TemplateHtml = this.replaceTokens("BearerToken", `Bearer ${this._bearerToken}`, this.properties.TemplateHtml);
      output += this.properties.TemplateHtml;
    } else {
      output += `<div class="${styles.templateError}">${strings.NoTemplateProvided}</div>`;
    }
    // Close the webpart container
    output += `</div>`;

    this.domElement.innerHTML = output;
    this.executeScript(this.domElement);

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
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
              groupFields: [
                PropertyPaneTextField('WebPartTitle', {
                  label: strings.WebPartTitle,
                  description: strings.WebPartTitleDescription
                }),
                PropertyPaneTextField('TemplateUrl', {
                  label: strings.TemplateUrl,
                  description: strings.TemplateUrlDescription
                }),
                PropertyPaneTextField('Audience', {
                  label: strings.Audience,
                  description: strings.AudienceDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private replaceTokens(name, value, str): string {
    return str.replaceAll("{" + name + "}", value);
  }

  private GetTemplateHtml(): Promise<string> {
    return new Promise<string>((resolve, reject) => {
      this.context.httpClient.get(`${this.properties.TemplateUrl}`, HttpClient.configurations.v1, {
        headers: {
          Accept: "text/html; charset=utf-8",
          "Content-Type": "text/html; charset=utf-8"
        }
      })
        .then((response: HttpClientResponse): any => {
          if (response.ok && response.status == 200) {
            resolve(response.text());
          } else if (response.status == 404) {
            resolve(`<div class="${styles.templateError}">${strings.TemplateError}</div>`);
          }
        });
    });
  }

  private GetCurrentWebAbsoluteUrl(): string {
    if (this._teamsContext) {
      return this._teamsContext.teamSiteUrl;
    } else {
      return this.context.pageContext.web.absoluteUrl;
    }
  }

  private GetCurrentWebTitle(): string {
    if (this._teamsContext) {
      return this._teamsContext.teamName;
    } else {
      return this.context.pageContext.web.title;
    }
  }

  private evalScript(elem) {
    const data = (elem.text || elem.textContent || elem.innerHTML || "");
    const headTag = document.getElementsByTagName("head")[0] || document.documentElement;
    const scriptTag = document.createElement("script");

    for (let i = 0; i < elem.attributes.length; i++) {
      const attr = elem.attributes[i];
      // Copies all attributes in case of loaded script relies on the tag attributes
      if (attr.name.toLowerCase() === "onload") continue; // onload handled after loading with SPComponentLoader
      scriptTag.setAttribute(attr.name, attr.value);
    }

    // set a bogus type to avoid browser loading the script, as it's loaded with SPComponentLoader
    scriptTag.type = (scriptTag.src && scriptTag.src.length) > 0 ? "pnp" : "text/javascript";
    // Ensure proper setting and adding id used in cleanup on reload
    scriptTag.setAttribute("pnpname", this._unqiueId);

    try {
      // doesn't work on ie...
      scriptTag.appendChild(document.createTextNode(data));
    } catch (e) {
      // IE has funky script nodes
      scriptTag.text = data;
    }

    headTag.insertBefore(scriptTag, headTag.firstChild);
  }

  // Clean up added script tags in case of smart re-load        
  private cleanUp(): void {
    if (this.domElement) {
      this.domElement.innerHTML = "";
    }
    const headTag = document.getElementsByTagName("head")[0] || document.documentElement;
    let scriptTags = headTag.getElementsByTagName("script");
    for (let i = 0; i < scriptTags.length; i++) {
      const scriptTag = scriptTags[i];
      if (scriptTag.hasAttribute("pnpname") && scriptTag.attributes["pnpname"].value == this._unqiueId) {
        headTag.removeChild(scriptTag);
      }
    }
  }

  // Finds and executes scripts in a newly added element's body.
  // Needed since innerHTML does not run scripts.
  //
  // Argument element is an element in the dom.
  private async executeScript(element: HTMLElement) {
    if (this.properties.spPageContextInfo && !window["_spPageContextInfo"]) {
      window["_spPageContextInfo"] = this.context.pageContext.legacyPageContext;
    }

    if (this.properties.teamsContext && !window["_teamsContexInfo"]) {
      window["_teamsContexInfo"] = this.context.sdks.microsoftTeams.context;
    }

    // Define global name to tack scripts on in case script to be loaded is not AMD/UMD
    (<any>window).ScriptGlobal = {};

    // main section of function
    const scripts = [];
    const children_nodes = element.getElementsByTagName("script");

    for (let i = 0; children_nodes[i]; i++) {
      const child: any = children_nodes[i];
      if (!child.type || child.type.toLowerCase() === "text/javascript") {
        scripts.push(child);
      }
    }

    const urls = [];
    const onLoads = [];
    for (let i = 0; scripts[i]; i++) {
      const scriptTag = scripts[i];
      if (scriptTag.src && scriptTag.src.length > 0) {
        urls.push(scriptTag.src);
      }
      if (scriptTag.onload && scriptTag.onload.length > 0) {
        onLoads.push(scriptTag.onload);
      }
    }

    let oldamd = null;
    if (window["define"] && window["define"].amd) {
      oldamd = window["define"].amd;
      window["define"].amd = null;
    }

    for (let i = 0; i < urls.length; i++) {
      try {
        let scriptUrl = urls[i];
        // Add unique param to force load on each run to overcome smart navigation in the browser as needed
        const prefix = scriptUrl.indexOf('?') === -1 ? '?' : '&';
        scriptUrl += prefix + 'pnp=' + new Date().getTime();
        await SPComponentLoader.loadScript(scriptUrl, { globalExportsName: "ScriptGlobal" });
      } catch (error) {
        if (console.error) {
          console.error(error);
        }
      }
    }
    if (oldamd) {
      window["define"].amd = oldamd;
    }

    for (let i = 0; scripts[i]; i++) {
      const scriptTag = scripts[i];
      if (scriptTag.parentNode) { scriptTag.parentNode.removeChild(scriptTag); }
      this.evalScript(scripts[i]);
    }
    // execute any onload people have added
    for (let i = 0; onLoads[i]; i++) {
      onLoads[i]();
    }
  }

}

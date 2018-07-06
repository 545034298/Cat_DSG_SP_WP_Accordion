import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IWebPartContext,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CatDsgSpWp1006AccordionWebPart.module.scss';
import * as strings from 'CatDsgSpWp1006AccordionWebPartStrings';
import CatDsgSpWp1006ScriptLoader from './CatDsgSpWp1006ScriptLoader';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
require('./CatDsgSpWp1006AccordionWebPart.scss');

export interface ICatDsgSpWp1006AccordionWebPartProps {
  description: string;
  listName: string;
  headerColumnName: string;
  contentColumnName: string;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Id: string;
  Title: string;
}

export interface IAccordion {
  Header: string;
  Content: string;
}

export default class CatDsgSpWp1006AccordionWebPart extends BaseClientSideWebPart<ICatDsgSpWp1006AccordionWebPartProps> {
  private spListDropDownOption: IPropertyPaneDropdownOption[] = [];

  public constructor(context: IWebPartContext) {
    super();
    SPComponentLoader.loadCss("https://code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css");
  }

  public onInit<T>(): Promise<T> {
    this.getSPLists()
      .then((response) => {
        this.spListDropDownOption = response.value.map((list: ISPList) => {
          return {
            key: list.Title,
            text: list.Title
          };
        });

      });
    return Promise.resolve();
  }
  public render(): void {
    this.context.statusRenderer.clearError(this.domElement);
    this.properties.description = strings.catDsgSpWp1006AccordionDescription;
    this.domElement.innerHTML = `
      <div class="${ styles.catDsgSpWp1006Accordion}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <div class="${styles.catDsgSpWp1006AccordionContainer}">
              </div>
            </div>
          </div>
        </div>
      </div>`;
    let script: CatDsgSpWp1006ScriptLoader.IScript = {
      Url: "https://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.11.1.min.js",
      GlobalExportsName: "jQuery",
      WindowPropertiesChain: "jQuery"
    };

    let dependencies: CatDsgSpWp1006ScriptLoader.IScript[] = [
      {
        Url: "https://ajax.aspnetcdn.com/ajax/jquery.ui/1.12.1/jquery-ui.min.js",
        GlobalExportsName: "jqueryui",
        WindowPropertiesChain: "jQuery.ui"
      }
    ];

    CatDsgSpWp1006ScriptLoader.LoadScript(script, dependencies).then((object) => {
      if ((this.properties.listName == null || (this.properties.listName != null && this.properties.listName.trim() == '')) ||
        (this.properties.headerColumnName == null || this.properties.headerColumnName != null && this.properties.headerColumnName.trim() == '') ||
        (this.properties.contentColumnName == null || this.properties.contentColumnName != null && this.properties.contentColumnName.trim() == '')) {
        this.context.statusRenderer.renderError(this.domElement, strings.catDsgSpWp1006AccordionConfigurationSettingsRequiredMessage);
      }
      else {
        this.getAccordions().then((accordions: IAccordion[]) => {
          let accordionsHtml = ``;
          if (accordions) {
            if (accordions.length > 0) {
              for (var j = 0; j < accordions.length; j++) {
                let accodionHtml = `
               <h3>${accordions[j].Header}</h3>
               <div>
                 <p>${accordions[j].Content}</p>
               </div>`;
                accordionsHtml += accodionHtml;
              }
            }
            else {
              accordionsHtml += strings.catDsgSpWp1006AccordionNoDataMessage;
            }
          }
          else {
            accordionsHtml += strings.catDsgSpWp1006AccordionNoDataMessage;
          }
          this.domElement.querySelector('.' + styles.catDsgSpWp1006AccordionContainer).innerHTML = accordionsHtml;
          ($(this.domElement).find('.' + styles.catDsgSpWp1006AccordionContainer) as any).accordion();

        }, (error: any) => {
          this.context.statusRenderer.renderError(this.domElement, error);
        });
      }
    }, error => {
      alert(error);
    });
  }

  private validateRequiredProperty(value: string): string {
    if (value === null || (value != null && value.trim().length === 0)) {
      return strings.catDsgSpWp1006AccordionRequiredPropertyMessage;
    }
    return "";
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneDropdown('listName', {
                  label: strings.catDsgSpWp1006AccordionDropdownLabelListName,
                  options: this.spListDropDownOption
                }),
                PropertyPaneTextField('headerColumnName', {
                  label: strings.catDsgSpWp1006AccordionFieldLabeHeaderColumnName,
                  onGetErrorMessage: this.validateRequiredProperty.bind(this)
                }),
                PropertyPaneTextField('contentColumnName', {
                  label: strings.catDsgSpWp1006AccordionFieldLabeContentColumnName,
                  onGetErrorMessage: this.validateRequiredProperty.bind(this)
                }),
              ]
            }
          ]
        }
      ]
    };
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected getAccordions(): Promise<IAccordion[]> {
    return new Promise<IAccordion[]>((resolve, reject) => {

      this.getListItems().then((listitems) => {
        let accordions: IAccordion[] = [];
        if (listitems && listitems.value) {
          if (listitems.value.length > 0) {
            for (var i = 0; i < listitems.value.length; i++) {
              var accordinItem: IAccordion = { Header: '', Content: '' };
              for (var columnName in listitems.value[i]) {
                if (typeof (columnName) == 'string') {
                  if ((columnName as string).trim().toLocaleLowerCase() == this.properties.headerColumnName.trim().toLocaleLowerCase()) {
                    accordinItem.Header = listitems.value[i][columnName];
                  }
                  if ((columnName as string).trim().toLocaleLowerCase() == this.properties.contentColumnName.trim().toLocaleLowerCase()) {
                    accordinItem.Content = listitems.value[i][columnName];
                  }
                }
              }
              accordions.push(accordinItem);
            }
          }
        }
        resolve(accordions);
      }, (error: any) => {
        reject(error);
      });
    });
  }

  private getSPLists(): Promise<ISPLists> {
    const queryString: string = '$select=Id,Title,RootFolder/ServerRelativeUrl&$expand=RootFolder';
    const sortString: string = '$orderby=Title asc';
    return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists?${queryString}&${sortString}&$filter=Hidden eq false`,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private getListItems(): Promise<any> {
    const queryString: string = '$select=*';
    const sortingString: string = '$orderby=Modified desc';
    let url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.properties.listName}')/items?${queryString}&${sortingString}`;
    return this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 404) {
          this.context.statusRenderer.renderError(this.domElement, strings.catDsgSpWp1006AccordionNotFoundMessage + `:'${this.properties.listName}'`);
          return [];
        }
        else if (response.status === 400) {
          this.context.statusRenderer.renderError(this.domElement, `${strings.catDsgSpWp1006AccordionBadRequestMessagePrefix}${url}`);
          return [];
        }
        else {
          return response.json();
        }
      }, (error: any) => {
        this.context.statusRenderer.renderError(this.domElement, error);
      });
  }
}

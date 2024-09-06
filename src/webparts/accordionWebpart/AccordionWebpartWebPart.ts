import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'AccordionWebpartWebPartStrings';
import MyAccordionTemplate from './MyAccordionTemplate';
import * as jquery from 'jquery';
import 'jqueryui'; // Ensure jqueryui is properly imported and configured
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IAccordionWebpartWebPartProps {
  description: string;
}

export default class AccordionWebpartWebPart extends BaseClientSideWebPart<IAccordionWebpartWebPartProps> {

  public constructor() {
    super();
    // CSS URL
    SPComponentLoader.loadCss('https://code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');
  }

  public render(): void {
    // Inject the accordion HTML template into the web part's DOM element
    this.domElement.innerHTML = MyAccordionTemplate.templateHtml;

    // Define the options for the jQuery UI accordion
    const accordionOptions: JQueryUI.AccordionOptions = {
      animate: true, // Enables animation for the accordion
      collapsible: true, // Allows all sections to be collapsible
      active: false, // Keeps all sections collapsed initially
      icons: {
        header: 'ui-icon-circle-arrow-e', // Icon for the header when collapsed
        activeHeader: 'ui-icon-circle-arrow-s' // Icon for the header when expanded
      }
    };

    // Apply the jQuery UI accordion functionality to the specified element
    jquery(this.domElement).find('.accordion').accordion(accordionOptions);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      console.log('Environment Message:', message); // Usage of environment message
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}



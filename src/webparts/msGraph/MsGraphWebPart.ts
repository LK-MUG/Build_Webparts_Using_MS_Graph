import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneTextField,
} from '@microsoft/sp-webpart-base';

import * as strings from 'MsGraphWebPartStrings';
import CandidateInternalInfo from './components/CandidateInternalInfo';

import { graph } from "@pnp/graph";
import { sp } from "@pnp/sp";
import { AadTokenProvider } from '@microsoft/sp-http';
import { MSGraphRestService } from './services/MSGraphRestService';
import { CandidateListService } from './services/CandidateListService';
import { CandidateGraphService } from './services/CandidateGraphServices';
import { ICandidateInternalInfoProps } from './components/ICandidateInternalInfoProps';

export interface IMsGraphWebPartProps {
  listName: string;
}

export default class MsGraphWebPart extends BaseClientSideWebPart<IMsGraphWebPartProps> {

  // need to instantiate the @pnp/graph api, requires your o-auth token
  // need to instantiate the @pnp/sp api, requires the context of the webpart
  public onInit(): Promise<void> {
    return new Promise((resolve, reject) => {
      sp.setup({
        spfxContext: this.context
      });
      
      this.context.aadTokenProviderFactory
        .getTokenProvider()
        .then((tokenProvider: AadTokenProvider) => {

          graph.setup({
            graph: {
              fetchClientFactory: () => {
                return new MSGraphRestService(tokenProvider);
              }
            }
          });


          resolve();
        })
        .catch(reject);
    });
  }
  
  // the rendering life-cycle phase of the webpart.
  public render(): void {
    // instantiates the Candidate internal information class.
    const element: React.ReactElement<ICandidateInternalInfoProps>  = React.createElement(
      // We use concrete classes here because this is the integration point with 
      // SharePoint and with our services.
      CandidateInternalInfo,
      {
        graphClient: new CandidateGraphService(graph),
        spListClient: new CandidateListService(sp, this.properties.listName)
      }
    );

    ReactDom.render(element, this.domElement);
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
                PropertyPaneTextField('listName', {
                  label: "List Name"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'AadTokenHandlerWebPartStrings';
import AadTokenHandler from './components/AadTokenHandler';
import { IAadTokenHandlerProps } from './components/IAadTokenHandlerProps';
import { HttpClient } from '@microsoft/sp-http';

export interface IAadTokenHandlerWebPartProps {
  serviceId: string;
  serviceUrl: string;
  cancelPopupWithError: boolean;
}

export default class AadTokenHandlerWebPart extends BaseClientSideWebPart<IAadTokenHandlerWebPartProps> {

  private _redirectUrl = "";
  private _redirectRequired = false;
  private _popup: () => void;
  private _popupRequired = false;

  private _log: string[] = [];

  public render(): void {
    const element: React.ReactElement<IAadTokenHandlerProps> = React.createElement(
      AadTokenHandler,
      {
        context: this.context,
        redirectionRequired: this._redirectRequired,
        redirectionUrl: this._redirectUrl,
        popupRequired: this._popupRequired,
        popup: this._popup,
        invoke: this._invoke,
        log: this._log
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    const aad = this.context.aadTokenProviderFactory;

    return super.onInit()
      .then(_ => aad.getTokenProvider())
      .then(provider => {

        provider.onBeforeRedirectEvent.add(this, event => {
          event.cancel();
          this._log.unshift("redirection is required for authentication");
          this._redirectUrl = event.redirectUrl;
          this._redirectRequired = true;
          this.render();
        });

        provider.popupEvent.add(this, event => {
          if (this.properties.cancelPopupWithError)
            event.cancel(new Error("cancel popup with error"));
          else
            event.cancel();

          this._log.unshift("popup is required for authentication");
          this._popup = () => { event.showPopup(); };
          this._popupRequired = true;
          this.render();
        });

        provider.tokenAcquisitionEvent.add(this, event => {
          this._log.unshift(`token acquisition event | message: ${event.message} | redirect url: ${event.redirectUrl}`);
        });
      });
  }

  private _invoke = (): Promise<void> => {

    this._log.unshift("initiating service call");
    this.render();

    const aad = this.context.aadTokenProviderFactory;

    return aad.getTokenProvider().then(provider => {
      this._log.unshift("retrieving access token");
      this.render();

      return provider.getToken(this.properties.serviceId);

    }).then(accessToken => {
      this._log.unshift("sending service request");
      this.render();

      const headers = new Headers({
        "authorization": `Bearer ${accessToken}`,
        "accept": "application/json"
      });

      const url = this.properties.serviceUrl;
      const client = this.context.httpClient;
      return client.get(url, HttpClient.configurations.v1, {
        headers: headers,
        credentials: "omit",
        mode: 'cors',
        cache: 'no-cache',
      });

    }).then(response => {
      this._log.unshift("processing service response");
      this.render();

      return response.json();

    }).then(result => {
      this._log.unshift(JSON.stringify(result, null, 3));
      this.render();

    }).catch(error => {
      this._log.unshift(`service call failed: ${error.message}`);
      this.render();

    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('serviceId', {
                  label: strings.ServiceIdLabel
                }),
                PropertyPaneTextField('serviceUrl', {
                  label: strings.ServiceUrlLabel
                }),
                PropertyPaneToggle('cancelPopupWithError', {
                  label: "Cancel popup with error"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

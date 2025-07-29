import * as React from 'react';
import * as ReactDom from 'react-dom';
// importa o objeto Version para controle de versão da Web Part
import { Version } from '@microsoft/sp-core-library';
// importa o tipo IPropertyPaneConfiguration e o componente PropertyPaneTextField que permitem configurar propriedades no painel de edição de Web Part (aquela caixinha lateral do SharePoint quando você edita uma Web Part)
import {
  type IPropertyPaneConfiguration,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
// importa a classe base da WebPart de cliente da qual você está herdando para criar sua própria Web Part
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
// interface usada para manipular o tema ativo (como mdo claro ou escuro) e as cores semânticas do tema
import { IReadonlyTheme } from '@microsoft/sp-component-base';
// importa os recursos de localização (strings de texto definidas para multiplos idiomas )
import * as strings from 'TreinamentoSpFxWebPartStrings';
// importa  componente React da Web Part e sua interface de props
import TreinamentoSpFx from './components/TreinamentoSpFx';
import { ITreinamentoSpFxProps } from './components/ITreinamentoSpFxProps';

// define quais propriedades a WebPart aceita
export interface ITreinamentoSpFxWebPartProps {
  description: string;
  sourceList: string; // lista de fontes de dados
  qtdItens: number; // quantidade de itens a serem buscados na lista
  contador: number; // contador para controle interno
}

// define a classe da Web Part, que herda de BaseClientSideWebPart
export default class TreinamentoSpFxWebPart extends BaseClientSideWebPart<ITreinamentoSpFxWebPartProps> {

  // define variáveis de estado da Web Part, como tema ativo, mensagem de ambiente e outros
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  // método chamado quando a Web Part é renderizada na página. Ele cria o componente React passando várias props úteis (como descrição, tema, nome do usuario, etc.) e renderiza o componente dentro do elemento e insere no DOM da Web Part.
  public render(): void {
    const element: React.ReactElement<ITreinamentoSpFxProps> = React.createElement(
      TreinamentoSpFx,
      {
        description: this.properties.description,
        sourceList: this.properties.sourceList,
        qtdItens: this.properties.qtdItens,
        contador: this.properties.contador,
        context: this.context,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }
// executado quando a WebPart é carregada. Aqui está pegando a mensagem do ambiente (ex: se está rodando no SharePoint, Teams, Outlook, etc.) e armazenando na variável _environmentMessage.
  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }


// este método retorna uma mensagem que identifica em qual ambiente a WebPart está rodando: Office, Outlook, Teams ou Sharepoint, e se está em localhost.
  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
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

  // detecta quando o tema do site muda e atualiza as variáveis CSS correspondentes
  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  // remove o componente React do DOM quando a Web Part é removida da página
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // define a versão da WebPart para controle interno do SharePoint
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
// configurações da webpart
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
                PropertyPaneTextField('sourceList', {
                  label: 'Lista de fontes de dados'
                }),
                PropertyPaneSlider('qtdItens', {
                  label: 'Quantidade de itens',
                  min: 1,
                  max: 20,
                  value: 10,
                  showValue: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

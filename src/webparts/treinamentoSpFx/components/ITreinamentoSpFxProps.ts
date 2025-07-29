import { WebPartContext } from "@microsoft/sp-webpart-base";

// quais as propriedades que o componente espera receber
export interface ITreinamentoSpFxProps {
  description: string;
  sourceList: string; // lista de fontes de dados
  qtdItens: number; // quantidade de itens a serem buscados na lista
  contador: number; // contador para controle interno
  context: WebPartContext; // contexto da Web Part, necessário para interações com o SharePoint
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}

import * as React from 'react';
import styles from './TreinamentoSpFx.module.scss';
import type { ITreinamentoSpFxProps } from './ITreinamentoSpFxProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ITreinamentoState } from './ITreinamentoState';
import { spfi, SPFI, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users';

export default class TreinamentoSpFx extends React.Component<ITreinamentoSpFxProps, ITreinamentoState> {
  constructor(props: ITreinamentoSpFxProps) {
    super(props);

    const sp: SPFI = spfi().using(SPFx(this.props.context));
    this.state = {
      items: [
      ],
      contador: 0,
      sp: sp
    };
  }

  public componentDidMount(): void {
    const items: any = this.state.sp.web.lists.getByTitle("My List").items();
    console.log(items);
    let arrayTemp = this.state.items;
    arrayTemp.push({ title: 'Item 1', id: 1 });

    setTimeout(() => {
      const contTemp = arrayTemp.length + this.state.contador

      this.setState({
        contador: contTemp,
      });
    }, 2000);
  }

  public render(): React.ReactElement<ITreinamentoSpFxProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.treinamentoSpFx} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}, {this.props.sourceList}, número de itens a serem buscados na lista: {this.props.qtdItens}, contador: {this.state.contador} </strong></div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
          </ul>
        </div>
      </section>
    );
  }
}

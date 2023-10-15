/* eslint-disable prefer-const */
/* eslint-disable dot-notation */
/* eslint-disable no-unused-expressions */
import * as React from "react";
import styles from "./HelloWorld.module.scss";
import type { IHelloWorldProps } from "./IHelloWorldProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { getSP } from "../PnPconfig";
import { SPFI } from "@pnp/sp";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

interface IHelloWorldState {
  items: string[];
}

export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldState> {
  public _sp: SPFI;
  private listName = "TestList";

  constructor(props: IHelloWorldProps) {
    super(props);
    this.state = {items:[]};
    this._sp = getSP();
    console.log(this._sp);
  }

  componentDidMount(): void {
    let items:string[]=[];
    this._sp.web.lists
      .getByTitle(this.listName)
      .items()
      .then((res) => {
        res && res.length > 0 && res.forEach((item) => {
          // eslint-disable-next-line no-unused-expressions
          item && items.push(item['Title']);
        });
        this.setState({
          items:items
        })
      }).
      catch(e=>{
        console.log(e);
      })
  }

  public render(): React.ReactElement<IHelloWorldProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <section
        className={`${styles.helloWorld} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        {
          this.state.items && this.state.items.length > 0 && this.state.items.map((item,i)=>{
            return (
              <div key={i}>
                <h5>{item}</h5>

              </div>
            );
          })
        }
        <div className={styles.welcome}>
          <img
            alt=""
            src={
              isDarkTheme
                ? require("../assets/welcome-dark.png")
                : require("../assets/welcome-light.png")
            }
            className={styles.welcomeImage}
          />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>
            Web part property value: <strong>{escape(description)}</strong>
          </div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is a extensibility model for
            Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest
            way to extend Microsoft 365 with automatic Single Sign On, automatic
            hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li>
              <a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">
                SharePoint Framework Overview
              </a>
            </li>
            <li>
              <a
                href="https://aka.ms/spfx-yeoman-graph"
                target="_blank"
                rel="noreferrer"
              >
                Use Microsoft Graph in your solution
              </a>
            </li>
            <li>
              <a
                href="https://aka.ms/spfx-yeoman-teams"
                target="_blank"
                rel="noreferrer"
              >
                Build for Microsoft Teams using SharePoint Framework
              </a>
            </li>
            <li>
              <a
                href="https://aka.ms/spfx-yeoman-viva"
                target="_blank"
                rel="noreferrer"
              >
                Build for Microsoft Viva Connections using SharePoint Framework
              </a>
            </li>
            <li>
              <a
                href="https://aka.ms/spfx-yeoman-store"
                target="_blank"
                rel="noreferrer"
              >
                Publish SharePoint Framework applications to the marketplace
              </a>
            </li>
            <li>
              <a
                href="https://aka.ms/spfx-yeoman-api"
                target="_blank"
                rel="noreferrer"
              >
                SharePoint Framework API reference
              </a>
            </li>
            <li>
              <a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">
                Microsoft 365 Developer Community
              </a>
            </li>
          </ul>
        </div>
      </section>
    );
  }
}

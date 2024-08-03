import * as React from "react";
import styles from "./GetListItemsPnp.module.scss";
import type { IGetListItemsPnpProps } from "./IGetListItemsPnpProps";

import { SPFI } from "@pnp/sp";
import { getSP } from "../pnpjsConfig";
import { Logger } from "@pnp/logging";
import { IStackTokens, PrimaryButton, Stack, TextField } from "@fluentui/react";

export interface IReponseItem {
  ID: number;
  Title: string;
}

export interface IpnpState {
  items: IReponseItem[];
}

export class PnpState implements IpnpState {
  constructor(public items: IReponseItem[] = []) {}
}

export default class GetListItemsPnp extends React.Component<
  IGetListItemsPnpProps,
  PnpState
> {
  private LOG_SOURCE = "pnp-webpart";
  private LIST_NAME = "Clients";
  private _sp: SPFI;

  private customSpacingStackTokens: IStackTokens = {
    childrenGap: "10%",
    padding: "s1 15%",
  };

  constructor(props: IGetListItemsPnpProps) {
    super(props);
    this.state = new PnpState();
    this._sp = getSP();
  }

  public componentDidMount(): void {
    /*eslint no-void: ["error", { "allowAsStatement": true }]*/    
    //    Logger.write(`${this.LOG_SOURCE} componentDidMount - ${this.state.items}`)
  }

  private getListItems = async (listName: string): Promise<void> => {
    const items: IReponseItem[] = await this._sp.web.lists
      .getByTitle(listName)
      .items.select("Id", "Title")();
    this.setState({ items });
    console.log(items);
    Logger.write(`${this.LOG_SOURCE} componentDidMount - ${items}`);
  };

  public render(): React.ReactElement<IGetListItemsPnpProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <section
        className={`${styles.getListItemsPnp} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <div className={styles.welcome}>
          <div className={""}>
            <div className={""}>
              <div className={""}>
                {this.state.items.map((row: IReponseItem, index: number) => (
                  <Stack
                    key={row.ID}
                    horizontal
                    tokens={this.customSpacingStackTokens}
                  >
                    <TextField
                      label={row.Title}
                      underlined
                      value={row.Title}
                      onChange={() => {
                        row.Title = "test";
                      }}
                      key={"btn" + row.ID}
                    />
                  </Stack>
                ))}
              </div>
            </div>
          </div>
        </div>
        <div>
          <PrimaryButton text="Get List Items" onClick={() => this.getListItems(this.LIST_NAME)} />
        </div>
        <div>
          {description} {isDarkTheme} {environmentMessage} {userDisplayName}
        </div>
      </section>
    );
  }
}

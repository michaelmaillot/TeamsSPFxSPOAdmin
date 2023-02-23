import * as React from 'react';
import {
  Button,
  ButtonProps,
  SearchIcon,
  Tooltip,
} from "@fluentui/react-northstar";
import { getSP } from 'PnPJsConfig';
import { SPFI } from '@pnp/sp';
import { IWeb, Web } from '@pnp/sp/webs';

import styles from "styles/Common.module.scss";

export interface ISearchBarProps {
  siteUrl: string;
  disabled: boolean;
}

enum SearchBoxInNavBar {
  Inherit,
  AllPages,
  ModernOnly,
  Hidden
}

interface ISearchBarState {
  isLoading: boolean;
  searchBarStatus: SearchBoxInNavBar;
}

export default class SearchBar extends React.Component<ISearchBarProps, ISearchBarState> {
  private _sp: SPFI;
  private _web: IWeb;

  public constructor(props: ISearchBarProps) {
    super(props);

    this.state = {
      isLoading: false,
      searchBarStatus: SearchBoxInNavBar.Inherit
    };

    this._sp = getSP();
  }

  // eslint-disable-next-line @microsoft/spfx/no-async-await
  public async componentDidMount(): Promise<void> {
    this._web = Web([this._sp.web, this.props.siteUrl]);

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const searchBar: number = (await this._web.select("SearchBoxInNavBar")() as any).SearchBoxInNavBar;
    this.setState({
      searchBarStatus: searchBar
    });
  }

  public render(): React.ReactElement<ISearchBarProps> {
    return (
      <Tooltip
        trigger={
          <Button
            disabled={this.props.disabled}
            icon={this.state.searchBarStatus === SearchBoxInNavBar.Inherit ? <div className={styles.disabled}><SearchIcon /></div> : <SearchIcon />}
            loading={this.state.isLoading}
            iconOnly
            onClick={this._setSiteSearchBar} />}
        content={this.state.searchBarStatus === SearchBoxInNavBar.Inherit ? "Hide the search bar" : "Display the search bar"} />
    );
  }

  // eslint-disable-next-line @microsoft/spfx/no-async-await
  private _setSiteSearchBar = async (event: React.SyntheticEvent<HTMLElement, Event>, _data: ButtonProps): Promise<void> => {

    event.stopPropagation();

    this.setState({
      isLoading: true,
    });
    try {
      const newStatus: SearchBoxInNavBar = this.state.searchBarStatus === SearchBoxInNavBar.Inherit ? SearchBoxInNavBar.Hidden : SearchBoxInNavBar.Inherit;
      await this._web.update({
        SearchBoxInNavBar: newStatus
      });

      this.setState({
        searchBarStatus: newStatus,
      });
    } catch (error) {
      console.log(error);
    }
    finally {
      this.setState({
        isLoading: false,
      });
    }
  }
}
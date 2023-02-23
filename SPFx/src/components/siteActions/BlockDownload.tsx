import * as React from 'react';
import {
  Button,
  ButtonProps,
  DownloadIcon,
  Tooltip,
} from "@fluentui/react-northstar";

import styles from "styles/Common.module.scss";
import AdminServices from 'services/AdminServices';

export interface IBlockDownloadProps {
  blockDownloadEnabled: boolean;
  siteUrl: string;
  disabled: boolean;
}

interface IBlockDownloadState {
  isLoading: boolean;
  newBlockDownloadState: boolean;
}

export default class BlockDownload extends React.Component<IBlockDownloadProps, IBlockDownloadState> {
  public constructor(props: IBlockDownloadProps) {
    super(props);

    this.state = {
      isLoading: false,
      newBlockDownloadState: props.blockDownloadEnabled
    };
  }

  public render(): React.ReactElement<IBlockDownloadProps> {
    return (
      <Tooltip
        trigger={
          <Button
            disabled={this.props.disabled}
            icon={this.state.newBlockDownloadState ? <div className={styles.disabled}><DownloadIcon /></div> : <DownloadIcon />}
            loading={this.state.isLoading}
            iconOnly
            onClick={this._setBlockDownloadState} />}
        content={this.state.newBlockDownloadState ? "Allow file downloads" : "Block file downloads"} />
    );
  }

  // eslint-disable-next-line @microsoft/spfx/no-async-await
  private _setBlockDownloadState = async (event: React.SyntheticEvent<HTMLElement, Event>, _data: ButtonProps): Promise<void> => {

    event.stopPropagation();

    this.setState({
      isLoading: true,
    });
    try {
      const newState: boolean = !this.state.newBlockDownloadState;
      await AdminServices.UpdateSiteProperties(this.props.siteUrl, { "BlockDownloadPolicy": newState });

      this.setState({
        newBlockDownloadState: newState
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
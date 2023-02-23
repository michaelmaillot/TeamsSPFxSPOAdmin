/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @microsoft/spfx/no-async-await */
/* eslint-disable react/jsx-no-bind */
import * as React from 'react';
import {
  Header,
  Label,
  Checkbox,
  Dropdown,
  Button,
  Flex,
  Loader,
  Divider,
} from "@fluentui/react-northstar";
import { Caching } from "@pnp/queryable";

import { SPHttpClient } from '@microsoft/sp-http';
import AdminServices from 'services/AdminServices';
import { FlowsPolicy, ITenantSitePropertiesInfo, SharingCapabilities } from '@pnp/sp-admin';
import { Panel } from '@fluentui/react/lib/components/Panel/Panel';
import { getSP } from 'PnPJsConfig';

export interface IMoreActionsProps {
  spHttpCtx: SPHttpClient;
  isPanelOpen: boolean;
  site: Partial<ITenantSitePropertiesInfo>;
  isHomeSite: boolean;
  isCommSite: boolean;
  closePanel: () => void;
  closePanelSuccess: () => void;
  refreshHomeSite: (newSiteUrl: string) => void;
}

interface IMoreActionsState {
  isLoading: boolean;
  isNewHomeSite: boolean;
}

export default class MoreActions extends React.Component<IMoreActionsProps, IMoreActionsState> {
  private readonly flowPolicies = (Object.keys(FlowsPolicy).filter(value => isNaN(Number(value)) === false).map((key: any) => FlowsPolicy[key]));
  private readonly sharingCapabilities = (Object.keys(SharingCapabilities).filter(value => isNaN(Number(value)) === false).map((key: any) => SharingCapabilities[key]));

  public constructor(props: IMoreActionsProps) {
    super(props);

    this.state = {
      isLoading: false,
      isNewHomeSite: false,
    };
  }

  public render(): React.ReactElement<IMoreActionsProps> {
    return (
      <Panel
        styles={{ main: { width: "500px !important" } }} // TODO: get styles from Northstar theme to apply
        closeButtonAriaLabel="Close"
        isOpen={this.props.isPanelOpen}
        onDismiss={(_ev) => { this.setState({ isNewHomeSite: false }); this.props.closePanel() }}>
        <Header>Site properties</Header>
        <b>Site: </b><Label circular color="brand">{this.props.site?.Title}</Label>
        <Divider styles={{ paddingTop: "15px", paddingBottom: "15px" }} />
        {this.state.isLoading
          ?
          <Loader label="Applying changes..." />
          :
          <Flex gap="gap.small" column>
            {this.props.isCommSite &&
              <Checkbox labelPosition='start' styles={{ maxWidth: "300px", paddingLeft: "0px" }} label="Comments on site pages disabled" toggle checked={this.props.site?.CommentsOnSitePagesDisabled} onChange={(_ev, data) => this._updateSiteProperty({ "CommentsOnSitePagesDisabled": data.checked })} />
            }
            {!this.props.isHomeSite && this.props.isCommSite &&
              <Checkbox labelPosition='start' disabled={this.state.isNewHomeSite} styles={{ maxWidth: "300px", paddingLeft: "0px" }} label="Set this site as Home Site" checked={this.state.isNewHomeSite} onChange={this._setHomeSite} />
            }
            <span>
              Flow policy:{' '}
              <Dropdown inline items={this.flowPolicies} defaultValue={FlowsPolicy[this.props.site?.DisableFlows]} onChange={(_ev, data) => this._updateSiteProperty({ "DisableFlows": data.highlightedIndex })} />
            </span>
            <span>
              Sharing Capabilities:{' '}
              <Dropdown inline items={this.sharingCapabilities} defaultValue={SharingCapabilities[this.props.site?.SharingCapability]} onChange={(_ev, data) => this._updateSiteProperty({ "SharingCapability": data.highlightedIndex })} />
            </span>
          </Flex>
        }
        <Flex gap='gap.small' styles={{ position: "absolute", bottom: "0", paddingBottom: "15px" }}>
          <Button content="Close" type='button' secondary onClick={(_ev, _data) => this.props.closePanel()} disabled={this.state.isLoading} />
        </Flex>
      </Panel>
    );
  }

  private _updateSiteProperty = async (updatedProperty: Record<string, any>): Promise<void> => {
    this.setState({
      isLoading: true,
    });

    try {
      await AdminServices.UpdateSiteProperties(this.props.site.Url, updatedProperty);

    } catch (error) {
      console.log(error);
    }
    finally {
      this.setState({
        isLoading: false,
      });
    }
  }

  private _setHomeSite = async (): Promise<void> => {

    try {
      this.setState({
        isLoading: true,
      });

      const contextUrl: string = (await getSP().using(Caching()).site()).Url;

      await (await this.props.spHttpCtx.post(
        contextUrl + "/_api/SPHSite/SetSPHSite",
        SPHttpClient.configurations.v1,
        {
          body: JSON.stringify({
            siteUrl: this.props.site.Url
          }),
        })).json();

      this.setState({
        isNewHomeSite: true
      });

      this.props.refreshHomeSite(this.props.site.Url);
    } catch (error) {
      console.error(error);
    }
    finally {
      this.setState({
        isLoading: false,
      });
    }
  }
}
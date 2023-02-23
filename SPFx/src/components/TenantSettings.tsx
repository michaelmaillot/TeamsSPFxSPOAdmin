/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/typedef */
/* eslint-disable react/jsx-no-bind */
/* eslint-disable @microsoft/spfx/no-async-await */
import * as React from 'react';
import {
  Flex,
  Checkbox,
  Loader,
  CheckboxProps,
  Button,
  Dropdown,
} from "@fluentui/react-northstar";

import { SharingCapabilities, ImageTaggingChoice, SharingDomainRestrictionModes } from '@pnp/sp-admin';
import { MSGraphClientFactory, MSGraphClientV3 } from '@microsoft/sp-http';
import { Guid } from '@microsoft/sp-core-library';

export interface ITenantSettingsProps {
  graphClient: MSGraphClientFactory;
}

interface ITenantAdminSettings {
  allowedDomainGuidsForSyncApp: Guid;
  availableManagedPathsForSiteCreation: string[];
  deletedUserPersonalSiteRetentionPeriodInDays: number;
  excludedFileExtensionsForSyncApp: string[];
  idleSessionSignOut: any;
  imageTaggingOption: ImageTaggingChoice;
  isCommentingOnSitePagesEnabled: boolean;
  isFileActivityNotificationEnabled: boolean;
  isLegacyAuthProtocolsEnabled: boolean;
  isLoopEnabled: boolean;
  isMacSyncAppEnabled: boolean;
  isRequireAcceptingUserToMatchInvitedUserEnabled: boolean;
  isResharingByExternalUsersEnabled: boolean;
  isSharePointMobileNotificationEnabled: boolean;
  isSharePointNewsfeedEnabled: boolean;
  isSiteCreationEnabled: boolean;
  isSiteCreationUIEnabled: boolean;
  isSitePagesCreationEnabled: boolean;
  isSitesStorageLimitAutomatic: boolean;
  isSyncButtonHiddenOnPersonalSite: boolean;
  isUnmanagedSyncAppForTenantRestricted: boolean;
  personalSiteDefaultStorageLimitInMB: number;
  sharingAllowedDomainList: string[];
  sharingBlockedDomainList: string[];
  sharingCapability: string;
  sharingDomainRestrictionMode: SharingDomainRestrictionModes;
  siteCreationDefaultManagedPath: string;
  siteCreationDefaultStorageLimitInMB: number;
  tenantDefaultTimezone: string;
}

interface ITenantSettingsState {
  [key: string]: any;
  isLoading: boolean;
  isSiteCreationEnabled: boolean;
  isSiteCreationUIEnabled: boolean;
  availableManagedPathsForSiteCreation: string[];
  isCommentingOnSitePagesEnabled: boolean;
  isLoopEnabled: boolean;
  isMacSyncAppEnabled: boolean;
  sharingCapability: string;
  siteCreationDefaultManagedPath: string;
}

export default class TenantSettings extends React.Component<ITenantSettingsProps, ITenantSettingsState> {
  private _sharingCapabilities: any[];

  public constructor(props: ITenantSettingsProps) {
    super(props);

    this._sharingCapabilities = 
      Object.keys(SharingCapabilities)
      .filter(value => isNaN(Number(value)) === false)
      .map((key: any) => (SharingCapabilities[key].charAt(0).toLowerCase() + SharingCapabilities[key].slice(1)));

    this.state = {
      isLoading: true,
      isSiteCreationEnabled: false,
      isSiteCreationUIEnabled: false,
      availableManagedPathsForSiteCreation: [],
      isCommentingOnSitePagesEnabled: false,
      isLoopEnabled: false,
      isMacSyncAppEnabled: false,
      sharingCapability: this._sharingCapabilities[0],
      siteCreationDefaultManagedPath: "",
    };
  }

  public componentDidMount(): void {
    this._getTenantSettings().catch(err => console.log(err));
  }

  public render(): React.ReactElement<ITenantSettingsProps> {
    return (
      <>
        {this.state.isLoading
          ?
          <Loader label='Loading...' />
          :
          <Flex>
            <Flex column gap='gap.small' styles={{ maxWidth: "500px", paddingLeft: "15px" }}>
              <Button loading={this.state.isLoading} styles={{ maxWidth: "150px" }} primary content="Update settings" onClick={this._updateSetting} />
              <Checkbox
                labelPosition='start'
                label="Site creation (App Bar)"
                defaultChecked={this.state.isSiteCreationEnabled} checked={this.state.isSiteCreationEnabled}
                toggle
                data-tenant-setting={"isSiteCreationEnabled"}
                onChange={this._toggleSettingChange} />
              <Checkbox
                labelPosition='start'
                label="Site creation (UI)"
                defaultChecked={this.state.isSiteCreationUIEnabled}
                checked={this.state.isSiteCreationUIEnabled}
                toggle
                data-tenant-setting={"isSiteCreationUIEnabled"}
                onChange={this._toggleSettingChange} />
              <Checkbox
                labelPosition='start'
                label="Microsoft Loop"
                defaultChecked={this.state.isLoopEnabled}
                checked={this.state.isLoopEnabled}
                toggle
                data-tenant-setting={"isLoopEnabled"}
                onChange={this._toggleSettingChange} />
              <Checkbox
                labelPosition='start'
                label="Site pages commenting"
                defaultChecked={this.state.isCommentingOnSitePagesEnabled}
                checked={this.state.isCommentingOnSitePagesEnabled}
                toggle
                data-tenant-setting={"isCommentingOnSitePagesEnabled"}
                onChange={this._toggleSettingChange} />
              <Checkbox
                labelPosition='start'
                label="OneDrive sync app for Mac"
                defaultChecked={this.state.isMacSyncAppEnabled}
                checked={this.state.isMacSyncAppEnabled}
                toggle
                data-tenant-setting={"isMacSyncAppEnabled"}
                onChange={this._toggleSettingChange} />
              <Flex space='between' gap="gap.large" styles={{ maxWidth: "400px", paddingLeft: "5px", whiteSpace: "nowrap" }}>
                <Flex.Item size='size.half'>
                <span id="managed-path-label">Default managed path</span>
                </Flex.Item>
                <Flex.Item size='size.quarter' styles={{paddingLeft: "10px"}}>
                <Dropdown
                  aria-labelledby='managed-path-label'
                  items={this.state.availableManagedPathsForSiteCreation}
                  styles={{ "> div": { maxWidth: "90px" } }}
                  onChange={(event, data) => { console.log(data.value as string); this.setState({ siteCreationDefaultManagedPath: data.value as string }) }}
                  defaultValue={this.state.siteCreationDefaultManagedPath} />
                  </Flex.Item>
              </Flex>
              <Flex gap="gap.large" styles={{ maxWidth: "400px", paddingLeft: "5px", whiteSpace: "nowrap" }}>
                <span id="sharing-capability-label">Sharing capability</span>
                <Dropdown
                  aria-labelledby='sharing-capability-label'
                  items={this._sharingCapabilities}
                  styles={{ "> div": { maxWidth: "255px" } }}
                  onChange={(event, data) => { console.log(data.value as string); this.setState({ sharingCapability: data.value as string }) }}
                  defaultValue={this.state.sharingCapability} />
              </Flex>
            </Flex>
          </Flex>
        }
      </>
    );
  }

  private _toggleSettingChange = async (event: React.SyntheticEvent<HTMLElement, Event>, data: CheckboxProps): Promise<void> => {
    this.setState({
      [event.currentTarget.getAttribute("data-tenant-setting")]: data.checked as boolean,
    });
  }

  private _updateSetting = async (event: React.SyntheticEvent<HTMLElement, Event>): Promise<void> => {
    if (!this.state.isLoading) {
      try {
        this.setState({
          isLoading: true
        });

        const graphClient: MSGraphClientV3 = await this.props.graphClient.getClient("3");
        await graphClient
          .api(`https://graph.microsoft.com/beta/admin/sharepoint/settings`)
          .patch({
            isSiteCreationEnabled: this.state.isSiteCreationEnabled,
            isSiteCreationUIEnabled: this.state.isSiteCreationUIEnabled,
            availableManagedPathsForSiteCreation: this.state.availableManagedPathsForSiteCreation,
            isCommentingOnSitePagesEnabled: this.state.isCommentingOnSitePagesEnabled,
            isLoopEnabled: this.state.isLoopEnabled,
            isMacSyncAppEnabled: this.state.isMacSyncAppEnabled,
            sharingCapability: this.state.sharingCapability,
            siteCreationDefaultManagedPath: this.state.siteCreationDefaultManagedPath
          }, (error: any, results: ITenantAdminSettings) => {
            console.log(results);
          });
      } catch (error) {
        console.log(error);
      }
      finally {
        this.setState({
          isLoading: false
        });
      }
    }
  }

  private _getTenantSettings = async (): Promise<void> => {
    try {
      await this.props.graphClient.getClient("3").then(async (graphClient: MSGraphClientV3) => {
        await graphClient
          .api(`https://graph.microsoft.com/beta/admin/sharepoint/settings`)
          .get((error: any, results: ITenantAdminSettings) => {
            this.setState({
              isSiteCreationEnabled: results.isSiteCreationEnabled,
              isSiteCreationUIEnabled: results.isSiteCreationUIEnabled,
              availableManagedPathsForSiteCreation: results.availableManagedPathsForSiteCreation,
              isCommentingOnSitePagesEnabled: results.isCommentingOnSitePagesEnabled,
              isLoopEnabled: results.isLoopEnabled,
              isMacSyncAppEnabled: results.isMacSyncAppEnabled,
              sharingCapability: results.sharingCapability,
              siteCreationDefaultManagedPath: results.siteCreationDefaultManagedPath,
            });
          })
          .catch((reason: Error) => {
            throw reason;
          });
      })
        .catch((reason: Error) => {
          throw reason;
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